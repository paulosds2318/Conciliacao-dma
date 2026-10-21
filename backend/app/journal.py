"""Checkpoint and disk-backed result storage for desktop processing runs."""

import hashlib
import os
import sqlite3
from pathlib import Path


def database_path() -> Path:
    root = Path(os.getenv("CONCILIACAO_DMA_DATA_DIR") or os.getenv("LOCALAPPDATA") or Path.home())
    folder = root / "ConciliacaoDMA"
    folder.mkdir(parents=True, exist_ok=True)
    return folder / "processamentos.sqlite3"


def source_hash(source) -> str:
    digest = hashlib.sha256()
    if isinstance(source, (str, Path)):
        with Path(source).open("rb") as stream:
            for chunk in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(chunk)
    else:
        digest.update(bytes(source))
    return digest.hexdigest()


def job_key(source, action, column, has_header, api_url="") -> str:
    material = "\0".join((source_hash(source), action, column or "", str(bool(has_header)), api_url))
    return hashlib.sha256(material.encode("utf-8")).hexdigest()


def connect(path=None):
    db = sqlite3.connect(path or database_path(), timeout=30)
    db.row_factory = sqlite3.Row
    db.execute("PRAGMA journal_mode=WAL")
    db.execute("""CREATE TABLE IF NOT EXISTS items (
        job_id TEXT NOT NULL,
        ordinal INTEGER NOT NULL,
        contract TEXT NOT NULL,
        contract_hash TEXT NOT NULL,
        consider_dma INTEGER NOT NULL,
        disregard_dma INTEGER NOT NULL,
        state TEXT NOT NULL,
        status TEXT NOT NULL,
        updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
        PRIMARY KEY (job_id, ordinal)
    )""")
    db.execute("DELETE FROM items WHERE updated_at < datetime('now', '-30 days')")
    db.execute("CREATE INDEX IF NOT EXISTS items_job_state ON items(job_id, state, ordinal)")
    db.commit()
    return db


def get_item(db, job_id, ordinal):
    return db.execute("SELECT * FROM items WHERE job_id=? AND ordinal=?", (job_id, ordinal)).fetchone()


def save_item(db, job_id, ordinal, contract, consider_dma, disregard_dma, state, status):
    contract_hash = hashlib.sha256(contract.encode("utf-8")).hexdigest()
    db.execute(
        """INSERT INTO items (job_id, ordinal, contract, contract_hash, consider_dma, disregard_dma, state, status)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(job_id, ordinal) DO UPDATE SET
        contract=excluded.contract, contract_hash=excluded.contract_hash,
        consider_dma=excluded.consider_dma, disregard_dma=excluded.disregard_dma,
        state=excluded.state, status=excluded.status, updated_at=CURRENT_TIMESTAMP""",
        (job_id, ordinal, contract, contract_hash, int(consider_dma), int(disregard_dma), state, status),
    )
    db.commit()


def count_items(db, job_id, filter_name="Todos"):
    where = "job_id=?"
    args = [job_id]
    if filter_name == "Somente falhas":
        where += " AND state IN ('failed','uncertain','in_flight')"
    elif filter_name == "Somente sucessos":
        where += " AND state='success'"
    return db.execute(f"SELECT COUNT(*) FROM items WHERE {where}", args).fetchone()[0]


def page_items(db, job_id, filter_name="Todos", limit=200, offset=0):
    where = "job_id=?"
    args = [job_id]
    if filter_name == "Somente falhas":
        where += " AND state IN ('failed','uncertain','in_flight')"
    elif filter_name == "Somente sucessos":
        where += " AND state='success'"
    args.extend((limit, offset))
    return db.execute(
        f"SELECT contract AS contrato, consider_dma, disregard_dma, status FROM items WHERE {where} ORDER BY ordinal LIMIT ? OFFSET ?",
        args,
    ).fetchall()


def iter_items(db, job_id):
    cursor = db.execute(
        "SELECT contract, consider_dma, disregard_dma, status FROM items WHERE job_id=? ORDER BY ordinal",
        (job_id,),
    )
    yield from cursor
