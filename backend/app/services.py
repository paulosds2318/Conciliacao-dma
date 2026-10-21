import io
import logging
import os
import re
import time
from pathlib import Path
import requests
import pandas as pd
from app.config import CORAL_API_URL, MODO_TESTE_API, get_coral_auth_token
from app.spreadsheet import inspecionar_planilha_xlsx, iterar_contratos_xlsx, nome_coluna_excel, normalizar_cabecalhos
from app import journal


def _logger():
    folder = Path(os.getenv("LOCALAPPDATA", Path.home())) / "ConciliacaoDMA" / "logs"
    try:
        folder.mkdir(parents=True, exist_ok=True)
        logging.basicConfig(
            filename=folder / "conciliacao.log",
            level=logging.INFO,
            format="%(asctime)s %(levelname)s %(message)s",
            encoding="utf-8",
        )
    except OSError:
        logging.basicConfig(level=logging.WARNING)
    return logging.getLogger("conciliacao_dma")


def detectar_cabecalho(df_sem_cabecalho: pd.DataFrame) -> bool:
    if df_sem_cabecalho.empty:
        return True
    valores = [valor for valor in df_sem_cabecalho.iloc[0].tolist() if not pd.isna(valor) and str(valor).strip()]
    if not valores:
        # Uma primeira linha vazia não é um cabeçalho: deve ser ignorada como
        # qualquer linha sem contrato, preservando os dados das linhas abaixo.
        return False
    padrao_cabecalho = re.compile(r"\b(contrato|contract|n[u\u00fa]mero|c[o\u00f3]digo|id)\b", re.IGNORECASE)
    if any(padrao_cabecalho.search(str(valor)) for valor in valores):
        return True
    valores_com_cara_de_dado = sum(
        bool(re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9._/ -]{2,}", str(valor).strip()))
        and any(caractere.isdigit() for caractere in str(valor))
        for valor in valores
    )
    return valores_com_cara_de_dado < max(1, (len(valores) + 1) // 2)


def _nome_coluna_excel(indice: int) -> str:
    nome = ""
    while indice >= 0:
        indice, resto = divmod(indice, 26)
        nome = chr(65 + resto) + nome
        indice -= 1
    return f"Coluna {nome}"


def _normalizar_cabecalhos(valores) -> list[str]:
    nomes = []
    usados = set()
    for indice, valor in enumerate(valores):
        nome = str(valor).strip() if not pd.isna(valor) else ""
        if not nome:
            nome = _nome_coluna_excel(indice)
        base = nome
        sufixo = 2
        while nome.casefold() in usados:
            nome = f"{base} ({sufixo})"
            sufixo += 1
        usados.add(nome.casefold())
        nomes.append(nome)
    return nomes


def _post_com_tentativas(url, payload, headers, tentativas=3):
    logger = _logger()
    for tentativa in range(1, tentativas + 1):
        try:
            response = requests.post(url, json=payload, headers=headers, timeout=(5, 20))
            transitoria = response.status_code == 429 or response.status_code >= 500
            if not transitoria or tentativa == tentativas:
                if not 200 <= response.status_code < 300:
                    logger.error("Falha HTTP %s após %s tentativa(s)", response.status_code, tentativa)
                return response
            logger.warning("HTTP %s; tentativa %s/%s", response.status_code, tentativa + 1, tentativas)
            time.sleep(tentativa)
        except requests.RequestException as error:
            if tentativa == tentativas:
                logger.error("Falha de rede após %s tentativas (%s)", tentativas, type(error).__name__)
                raise
            logger.warning("Falha de rede (%s); tentativa %s/%s", type(error).__name__, tentativa + 1, tentativas)
            time.sleep(tentativa)


def processar_e_enviar_planilha(conteudo_arquivo, acao: str, coluna_contrato: str | None = None, tem_cabecalho: bool | None = None, force_uncertain: bool = False, stream_results: bool = False):
    if acao not in ("Desconsiderar DMA", "Considerar DMA"):
        raise ValueError("Acao invalida. Escolha Desconsiderar DMA ou Considerar DMA.")

    is_xlsx = isinstance(conteudo_arquivo, (bytes, bytearray)) and bytes(conteudo_arquivo[:2]) == b"PK"
    if isinstance(conteudo_arquivo, (str, Path)):
        is_xlsx = str(conteudo_arquivo).lower().endswith((".xlsx", ".xlsm"))
    if is_xlsx:
        info = inspecionar_planilha_xlsx(conteudo_arquivo, tem_cabecalho)
        tem_cabecalho = info["tem_cabecalho"]
        names, counts = info["colunas"], info["contagens"]
        if coluna_contrato is None:
            candidates = [i for i, count in enumerate(counts) if count]
            if not candidates:
                raise ValueError("Nao encontrei contratos preenchidos na planilha.")
            index = next((i for i in candidates if "contrato" in names[i].casefold()), candidates[0])
            coluna_contrato = names[index]
        if coluna_contrato not in names:
            raise ValueError(f"A coluna '{coluna_contrato}' nao existe na planilha.")
        index = names.index(coluna_contrato)
        total_contratos = counts[index]
        linhas_dados = info["linhas_dados"]
        contratos = iterar_contratos_xlsx(conteudo_arquivo, coluna_contrato, tem_cabecalho)
    else:
        fonte = conteudo_arquivo if isinstance(conteudo_arquivo, (str, Path)) else io.BytesIO(conteudo_arquivo)
        dados_sem_cabecalho = pd.read_excel(fonte, header=None)
        if dados_sem_cabecalho.empty:
            raise ValueError("A planilha enviada esta vazia.")
        if tem_cabecalho is None:
            tem_cabecalho = detectar_cabecalho(dados_sem_cabecalho)
        if tem_cabecalho:
            nomes_colunas = _normalizar_cabecalhos(dados_sem_cabecalho.iloc[0].tolist())
            df = dados_sem_cabecalho.iloc[1:].copy()
            df.columns = nomes_colunas
        else:
            df = dados_sem_cabecalho.copy()
            df.columns = [_nome_coluna_excel(indice) for indice in range(len(df.columns))]
        if coluna_contrato is None:
            candidatas = [coluna for coluna in df.columns if df[coluna].notna().any()]
            if not candidatas:
                raise ValueError("Nao encontrei contratos preenchidos na planilha.")
            coluna_contrato = next((c for c in candidatas if "contrato" in c.casefold()), candidatas[0])
        if coluna_contrato not in df.columns:
            raise ValueError(f"A coluna '{coluna_contrato}' nao existe na planilha.")
        contratos = []
        for valor in df[coluna_contrato]:
            if pd.isna(valor):
                continue
            if isinstance(valor, float) and valor.is_integer():
                valor = int(valor)
            numero = str(valor).strip()
            if numero:
                contratos.append(numero)
        linhas_dados = len(df)
        total_contratos = len(contratos)
    if total_contratos == 0:
        raise ValueError(f"A coluna '{coluna_contrato}' nao possui contratos preenchidos.")

    if not MODO_TESTE_API:
        if (
            not CORAL_API_URL
            or not CORAL_API_URL.lower().startswith(("https://", "http://"))
            or "exemplo.com" in CORAL_API_URL.lower()
        ):
            raise ValueError("Configure uma URL válida da API Coral antes de usar o modo real.")
        if not get_coral_auth_token():
            raise ValueError("Configure o token CORAL_AUTH_TOKEN antes de usar o modo real.")

    considerar_dma = acao == "Considerar DMA"
    headers = {
        "Authorization": f"Bearer {get_coral_auth_token()}",
        "Content-Type": "application/json",
    }
    job_id = journal.job_key(
        conteudo_arquivo,
        acao,
        coluna_contrato,
        tem_cabecalho,
        "simulacao" if MODO_TESTE_API else CORAL_API_URL,
    )
    store_details = stream_results or not MODO_TESTE_API
    db = journal.connect() if store_details else None
    resultados = {
        "total": total_contratos,
        "linhas_ignoradas": linhas_dados - total_contratos,
        "coluna_contrato": coluna_contrato,
        "tem_cabecalho": tem_cabecalho,
        "sucesso": 0,
        "falhas": 0,
        "modo": "simulacao" if MODO_TESTE_API else "real",
        "detalhes": [] if not stream_results else None,
        "detalhes_db": str(journal.database_path()) if stream_results else None,
        "job_id": job_id,
        "incertos": 0,
        "forcou_incertos": force_uncertain,
    }

    for numero_contrato in contratos:
        payload = {
            "contrato": numero_contrato,
            "considerar_dma": considerar_dma,
            "desconsiderar_dma": not considerar_dma,
        }

        if MODO_TESTE_API:
            resultados["sucesso"] += 1
            resultados["detalhes"].append({**payload, "status": "OK - API simulada"})
            continue

        try:
            response = _post_com_tentativas(CORAL_API_URL, payload, headers)
            if 200 <= response.status_code < 300:
                resultados["sucesso"] += 1
                status = "OK"
            else:
                resultados["falhas"] += 1
                status = f"Falha HTTP {response.status_code}"
            resultados["detalhes"].append({**payload, "status": status})
        except requests.RequestException as error:
            resultados["falhas"] += 1
            if isinstance(error, requests.Timeout):
                status = "Tempo limite excedido; verifique a API e tente novamente"
            else:
                status = "Falha de conexão; verifique a rede e tente novamente"
            resultados["detalhes"].append({**payload, "status": status})

    return resultados


# Desktop and API callers share the streaming, checkpointed pipeline.
from app.pipeline import processar_e_enviar_planilha
