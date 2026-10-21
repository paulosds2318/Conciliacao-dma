"""Streaming execution, safe resume checkpoints, and API delivery."""

import hashlib
import io
import logging
import os
import re
import sqlite3
import uuid
from pathlib import Path

import pandas as pd
import requests

from app.config import CORAL_API_URL, MODO_TESTE_API, get_coral_auth_token
from app.journal import connect, database_path, get_item, job_key, save_item
from app.spreadsheet import (
    inspecionar_planilha_xlsx,
    iterar_contratos_xlsx,
    nome_coluna_excel,
    normalizar_cabecalhos,
)


def _logger():
    root = os.getenv("CONCILIACAO_DMA_DATA_DIR") or os.getenv("LOCALAPPDATA") or Path.home()
    folder = Path(root) / "ConciliacaoDMA" / "logs"
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


def _as_path_or_bytes(source):
    return isinstance(source, (str, Path, bytes, bytearray))


def processar_e_enviar_planilha(
    source,
    acao: str,
    coluna_contrato: str | None = None,
    tem_cabecalho: bool | None = None,
    force_uncertain: bool = False,
    stream_results: bool = False,
    force_new: bool = False,
):
    if acao not in ("Desconsiderar DMA", "Considerar DMA"):
        raise ValueError("Ação inválida. Escolha Desconsiderar DMA ou Considerar DMA.")

    is_xlsx = isinstance(source, (bytes, bytearray)) and bytes(source[:2]) == b"PK"
    if isinstance(source, (str, Path)):
        is_xlsx = str(source).lower().endswith((".xlsx", ".xlsm"))

    if is_xlsx:
        info = inspecionar_planilha_xlsx(source, tem_cabecalho)
        tem_cabecalho = info["tem_cabecalho"]
        names, counts = info["colunas"], info["contagens"]
        if coluna_contrato is None:
            candidates = [i for i, count in enumerate(counts) if count]
            if not candidates:
                raise ValueError("Não encontrei contratos preenchidos na planilha.")
            index = next((i for i in candidates if "contrato" in names[i].casefold()), candidates[0])
            coluna_contrato = names[index]
        if coluna_contrato not in names:
            raise ValueError(f"A coluna '{coluna_contrato}' não existe na planilha.")
        column_index = names.index(coluna_contrato)
        total = counts[column_index]
        data_rows = info["linhas_dados"]
        contracts = iterar_contratos_xlsx(source, coluna_contrato, tem_cabecalho)
    else:
        source_for_pandas = source if isinstance(source, (str, Path)) else io.BytesIO(source)
        data = pd.read_excel(source_for_pandas, header=None)
        if data.empty:
            raise ValueError("A planilha enviada está vazia.")
        if tem_cabecalho is None:
            tem_cabecalho = _detect_header(data.iloc[0].tolist())
        if tem_cabecalho:
            names = normalizar_cabecalhos(data.iloc[0].tolist())
            frame = data.iloc[1:].copy()
            frame.columns = names
        else:
            frame = data.copy()
            frame.columns = [nome_coluna_excel(i) for i in range(len(frame.columns))]
        if coluna_contrato is None:
            candidates = [name for name in frame.columns if frame[name].notna().any()]
            if not candidates:
                raise ValueError("Não encontrei contratos preenchidos na planilha.")
            coluna_contrato = next((name for name in candidates if "contrato" in name.casefold()), candidates[0])
        if coluna_contrato not in frame.columns:
            raise ValueError(f"A coluna '{coluna_contrato}' não existe na planilha.")
        contracts_list = []
        for value in frame[coluna_contrato]:
            if pd.isna(value):
                continue
            if isinstance(value, float) and value.is_integer():
                value = int(value)
            text = str(value).strip()
            if text:
                contracts_list.append(text)
        contracts = iter(contracts_list)
        total = len(contracts_list)
        data_rows = len(frame)

    if total == 0:
        raise ValueError(f"A coluna '{coluna_contrato}' não possui contratos preenchidos.")
    if not MODO_TESTE_API:
        if not CORAL_API_URL or not CORAL_API_URL.lower().startswith(("https://", "http://")) or "exemplo.com" in CORAL_API_URL.lower():
            raise ValueError("Configure uma URL válida da API Coral antes de usar o modo real.")
        if not get_coral_auth_token():
            raise ValueError("Configure CORAL_AUTH_TOKEN no arquivo .env.")

    consider = acao == "Considerar DMA"
    endpoint_key = "simulacao" if MODO_TESTE_API else CORAL_API_URL
    run_id = job_key(source, acao, coluna_contrato, tem_cabecalho, endpoint_key)
    if force_new:
        run_id = hashlib.sha256(f"{run_id}:{uuid.uuid4()}".encode("utf-8")).hexdigest()
    use_db = stream_results or not MODO_TESTE_API
    db = connect() if use_db else None
    result = {
        "total": total,
        "linhas_ignoradas": data_rows - total,
        "coluna_contrato": coluna_contrato,
        "tem_cabecalho": tem_cabecalho,
        "sucesso": 0,
        "falhas": 0,
        "incertos": 0,
        "modo": "simulacao" if MODO_TESTE_API else "real",
        "detalhes": None if stream_results else [],
        "detalhes_db": str(database_path()) if stream_results else None,
        "job_id": run_id,
        "forcou_incertos": force_uncertain,
        "forcou_novo": force_new,
        "ja_processados": 0,
        "enviados": 0,
    }
    headers = {"Authorization": f"Bearer {get_coral_auth_token()}", "Content-Type": "application/json"}
    logger = _logger() if not MODO_TESTE_API else None

    try:
        for ordinal, contract in enumerate(contracts):
            payload = {
                "contrato": contract,
                "considerar_dma": consider,
                "desconsiderar_dma": not consider,
            }
            previous = get_item(db, run_id, ordinal) if db else None
            contract_hash = hashlib.sha256(contract.encode("utf-8")).hexdigest()
            if previous and previous["contract_hash"] != contract_hash:
                previous = None
            if previous and previous["state"] == "success":
                result["sucesso"] += 1
                result["ja_processados"] += 1
                if not stream_results:
                    result["detalhes"].append({**payload, "status": previous["status"]})
                continue
            if previous and previous["state"] in ("uncertain", "in_flight") and not force_uncertain:
                result["falhas"] += 1
                result["incertos"] += 1
                continue

            if MODO_TESTE_API:
                state, status = "success", "OK - API simulada"
            else:
                save_item(db, run_id, ordinal, contract, consider, not consider, "in_flight", "Envio iniciado")
                try:
                    response = requests.post(CORAL_API_URL, json=payload, headers=headers, timeout=(5, 20))
                    if 200 <= response.status_code < 300:
                        state, status = "success", "OK"
                    elif response.status_code >= 500:
                        state = "uncertain"
                        status = f"HTTP {response.status_code}; resultado incerto, confira na Coral antes de reenviar"
                    else:
                        state, status = "failed", f"Falha HTTP {response.status_code}"
                    if state != "success":
                        logger.error("API HTTP %s; resultado registrado sem identificador de contrato", response.status_code)
                except requests.RequestException as error:
                    state = "uncertain"
                    status = "Falha de rede; resultado incerto. Confira na Coral antes de reenviar."
                    logger.error("Falha de rede (%s); contrato omitido do log", type(error).__name__)

            if db:
                save_item(db, run_id, ordinal, contract, consider, not consider, state, status)
            if state == "success":
                result["sucesso"] += 1
            else:
                result["falhas"] += 1
                if state == "uncertain":
                    result["incertos"] += 1
            result["enviados"] += 1
            if not stream_results:
                result["detalhes"].append({**payload, "status": status})
    finally:
        if db:
            db.close()
    return result


def _detect_header(values):
    present = [value for value in values if value is not None and str(value).strip()]
    if not present:
        return False
    header_pattern = re.compile(r"\b(contrato|contract|n[uú]mero|c[oó]digo|id)\b", re.IGNORECASE)
    if any(header_pattern.search(str(value)) for value in present):
        return True
    data_like = sum(
        bool(re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9._/ -]{2,}", str(value).strip()))
        and any(char.isdigit() for char in str(value))
        for value in present
    )
    return data_like < max(1, (len(present) + 1) // 2)
