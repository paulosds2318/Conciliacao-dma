import io
import re
import requests
import pandas as pd
from app.config import CORAL_API_URL, CORAL_AUTH_TOKEN, MODO_TESTE_API


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


def processar_e_enviar_planilha(conteudo_arquivo: bytes, acao: str, coluna_contrato: str | None = None, tem_cabecalho: bool | None = None):
    if acao not in ("Desconsiderar DMA", "Considerar DMA"):
        raise ValueError("Acao invalida. Escolha Desconsiderar DMA ou Considerar DMA.")

    dados_sem_cabecalho = pd.read_excel(io.BytesIO(conteudo_arquivo), header=None)
    if dados_sem_cabecalho.empty:
        raise ValueError("A planilha enviada esta vazia.")
    if tem_cabecalho is None:
        tem_cabecalho = detectar_cabecalho(dados_sem_cabecalho)

    if tem_cabecalho:
        nomes_colunas = [str(coluna).strip() for coluna in dados_sem_cabecalho.iloc[0].tolist()]
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
    if not contratos:
        raise ValueError(f"A coluna '{coluna_contrato}' nao possui contratos preenchidos.")

    considerar_dma = acao == "Considerar DMA"
    headers = {
        "Authorization": f"Bearer {CORAL_AUTH_TOKEN}",
        "Content-Type": "application/json",
    }
    resultados = {
        "total": len(contratos),
        "linhas_ignoradas": len(df) - len(contratos),
        "coluna_contrato": coluna_contrato,
        "tem_cabecalho": tem_cabecalho,
        "sucesso": 0,
        "falhas": 0,
        "modo": "simulacao" if MODO_TESTE_API else "real",
        "detalhes": [],
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
            response = requests.post(CORAL_API_URL, json=payload, headers=headers, timeout=10)
            if response.status_code in (200, 201):
                resultados["sucesso"] += 1
                status = "OK"
            else:
                resultados["falhas"] += 1
                status = f"Erro {response.status_code}"
            resultados["detalhes"].append({**payload, "status": status})
        except Exception as error:
            resultados["falhas"] += 1
            resultados["detalhes"].append({**payload, "status": f"Falha de conexao: {error}"})

    return resultados
