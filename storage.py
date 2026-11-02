"""Persistencia compartida en PostgreSQL; SQLite para pruebas locales."""

from __future__ import annotations

from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path
import hashlib
import json
import sqlite3
from uuid import uuid4, uuid5, NAMESPACE_URL

from engine import as_json, review_defaults


class ConflictError(Exception):
    """Otro usuario guardó una revisión más reciente."""


class SourceChangedError(Exception):
    """Una nueva carga cambió un registro original ya incorporado."""


def local_url() -> str:
    path = Path(__file__).parent / "datos_locales" / "servicios.sqlite3"
    path.parent.mkdir(exist_ok=True)
    return f"sqlite:///{path}"


def _postgres(url: str) -> bool:
    return url.startswith(("postgresql://", "postgres://", "postgresql+psycopg://"))


@contextmanager
def connection(url: str):
    if _postgres(url):
        import psycopg
        from psycopg.rows import dict_row
        clean_url = url.replace("postgresql+psycopg://", "postgresql://", 1).replace("postgres://", "postgresql://", 1)
        conn = psycopg.connect(clean_url, row_factory=dict_row, connect_timeout=10)
    elif url.startswith("sqlite:///"):
        conn = sqlite3.connect(url.removeprefix("sqlite:///"), timeout=20)
        conn.row_factory = sqlite3.Row
    else:
        raise ValueError("DATABASE_URL debe ser una URL PostgreSQL o SQLite.")
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def _run(conn, url: str, sql: str, args=()):
    if _postgres(url):
        sql = sql.replace("?", "%s")
    return conn.execute(sql, args)


def init_db(url: str):
    with connection(url) as conn:
        _run(conn, url, """CREATE TABLE IF NOT EXISTS datasets (
            id TEXT PRIMARY KEY, label TEXT NOT NULL UNIQUE, filename TEXT NOT NULL,
            file_hash TEXT NOT NULL, headers_json TEXT NOT NULL, created_at TEXT NOT NULL
        )""")
        _run(conn, url, """CREATE TABLE IF NOT EXISTS records (
            id TEXT PRIMARY KEY, dataset_id TEXT NOT NULL,
            company_id TEXT NOT NULL, service_id TEXT NOT NULL, source_row INTEGER NOT NULL,
            source_json TEXT NOT NULL, edited_code TEXT NOT NULL,
            review_note TEXT NOT NULL, reviewer TEXT NOT NULL, decision TEXT NOT NULL,
            just_ciiu TEXT NOT NULL, just_cpc TEXT NOT NULL, just_match TEXT NOT NULL,
            observations TEXT NOT NULL, review_pc TEXT NOT NULL,
            revision INTEGER NOT NULL DEFAULT 0, updated_at TEXT NOT NULL,
            UNIQUE(dataset_id, company_id, service_id)
        )""")
        _run(conn, url, """CREATE TABLE IF NOT EXISTS audit (
            id TEXT PRIMARY KEY, record_id TEXT NOT NULL, previous_json TEXT NOT NULL,
            new_json TEXT NOT NULL, changed_by TEXT NOT NULL, changed_at TEXT NOT NULL
        )""")


def list_datasets(url: str) -> list[dict]:
    with connection(url) as conn:
        rows = _run(conn, url, "SELECT * FROM datasets ORDER BY created_at DESC").fetchall()
        return [dict(x) for x in rows]


DERIVED_OR_REVIEW = {
    "extrae_cpc", "Descripcion_cpc", "cpc_act_8dígitos", "ventas_tot", "id_emp",
    "CPC_FINAL", "CIIU_INICIAL_ 4 dig", "CPC_FINAL_ 4 dig", "descripcion",
    "Val CIIU servicios o excepción", "Val CPC servicios o excepción",
    "Val CIIU letra primera segunda tercera actividad",
    "Val CIIU divisiones primera segunda tercera actividad",
    "Ciiu excepción actividad principal", "Ciiu excepción actividad secundaria",
    "Ciiu excepción tercera actividad", "Al menos una actividad de la  CIIU es de excepción",
    "Verif_cpc", "Correspondencia=1",
    "Número de servicios que tienen correspondencia por empresa",
    "Al menos un servicio tiene correspondencia", "Revisar", "Nombre del responsable",
    "Decisión", "Justificación CIIU", "Justificación CPC",
    "Justificación Correspondencia", "Observaciones", "Revisiòn PC",
}


def _immutable_source(source: dict) -> dict:
    return {key: value for key, value in source.items() if key not in DERIVED_OR_REVIEW}


def import_records(url: str, label: str, filename: str, content: bytes,
                   headings: list[str], records: list[dict]) -> tuple[dict, int, int]:
    label = label.strip()
    if not label:
        raise ValueError("Indica un nombre para la ronda o lote.")
    digest = hashlib.sha256(content).hexdigest()
    now = datetime.now(timezone.utc).isoformat(timespec="seconds")
    with connection(url) as conn:
        existing = _run(conn, url, "SELECT * FROM datasets WHERE label=?", (label,)).fetchone()
        if existing:
            dataset = dict(existing)
        else:
            dataset = {"id": str(uuid4()), "label": label, "filename": filename,
                       "file_hash": digest, "headers_json": as_json(headings), "created_at": now}
            _run(conn, url, "INSERT INTO datasets VALUES (?,?,?,?,?,?)", tuple(dataset.values()))
        added, unchanged, changed = 0, 0, []
        for item in records:
            rid = str(uuid5(NAMESPACE_URL, f"{dataset['id']}|{item['company_id']}|{item['service_id']}"))
            old = _run(conn, url, "SELECT source_json FROM records WHERE id=?", (rid,)).fetchone()
            if old:
                if _immutable_source(json.loads(old["source_json"])) != _immutable_source(json.loads(as_json(item["source"]))):
                    changed.append(f"{item['company_id']} / {item['service_id']}")
                else:
                    unchanged += 1
                continue
            d = review_defaults(item["source"])
            _run(conn, url, """INSERT INTO records (
                id,dataset_id,company_id,service_id,source_row,source_json,
                edited_code,review_note,reviewer,decision,just_ciiu,just_cpc,
                just_match,observations,review_pc,revision,updated_at
            ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", (
                rid, dataset["id"], item["company_id"], item["service_id"], item["source_row"],
                as_json(item["source"]), d["edited_code"], d["review_note"], d["reviewer"],
                d["decision"], d["just_ciiu"], d["just_cpc"], d["just_match"],
                d["observations"], d["review_pc"], 0, now,
            ))
            added += 1
        if changed:
            raise SourceChangedError(
                "La carga trae cambios en registros originales ya incorporados ("
                + ", ".join(changed[:5]) + "). Use otro nombre de lote o concilie la fuente antes de reemplazarla."
            )
        return dataset, added, unchanged


def load_records(url: str, dataset_id: str) -> list[dict]:
    with connection(url) as conn:
        rows = _run(conn, url, "SELECT * FROM records WHERE dataset_id=? ORDER BY source_row", (dataset_id,)).fetchall()
        return [{**dict(r), "source": json.loads(r["source_json"])} for r in rows]


EDITABLE = ("edited_code", "review_note", "reviewer", "decision", "just_ciiu",
            "just_cpc", "just_match", "observations", "review_pc")


def save_review(url: str, record_id: str, expected_revision: int, changes: dict) -> None:
    if set(changes) != set(EDITABLE):
        raise ValueError("La revisión no contiene todos los campos esperados.")
    now = datetime.now(timezone.utc).isoformat(timespec="seconds")
    with connection(url) as conn:
        old = _run(conn, url, "SELECT * FROM records WHERE id=?", (record_id,)).fetchone()
        if old is None or old["revision"] != expected_revision:
            raise ConflictError("Otro revisor guardó este servicio. Actualiza la vista antes de continuar.")
        previous = {k: old[k] for k in EDITABLE}
        assignments = ",".join(f"{k}=?" for k in EDITABLE)
        cursor = _run(conn, url,
            f"UPDATE records SET {assignments}, revision=revision+1,updated_at=? WHERE id=? AND revision=?",
            tuple(str(changes[k]) for k in EDITABLE) + (now, record_id, expected_revision))
        if cursor.rowcount != 1:
            raise ConflictError("Otro revisor guardó este servicio. Actualiza la vista antes de continuar.")
        _run(conn, url, "INSERT INTO audit VALUES (?,?,?,?,?,?)", (
            str(uuid4()), record_id, as_json(previous), as_json(changes), changes["reviewer"], now))


def load_audit(url: str, dataset_id: str) -> list[dict]:
    with connection(url) as conn:
        rows = _run(conn, url, """SELECT a.*,r.company_id,r.service_id FROM audit a
            JOIN records r ON a.record_id=r.id WHERE r.dataset_id=? ORDER BY a.changed_at,a.id""", (dataset_id,)).fetchall()
        return [dict(r) for r in rows]
