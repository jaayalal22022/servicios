"""Carga y reglas del Tomo II, Servicios. No modifica los valores informados."""

from __future__ import annotations

from dataclasses import dataclass
from functools import lru_cache
from io import BytesIO
from pathlib import Path
import json
import re

from openpyxl import load_workbook


CATALOG_PATH = Path(__file__).parent / "catalogos" / "codigos.xlsx"
SERVICE_LETTERS = set("DEHIJKLMNPQRS")
EXCLUDED_DIVISIONS = {"K64", "K66", "Q88", "S94"}
SOURCE_FIELDS = (
    "inec_identificador_empresa", "ciiu4_actividad_principal",
    "ciiu4_actividad_secundaria", "ciiu4_tercera_actividad",
    "Descripción del servicio", "Código CPC", "id_serv", "servicio",
)
REVIEW_FIELDS = (
    "cpc_act_8dígitos", "Revisar", "Nombre del responsable", "Decisión",
    "Justificación CIIU", "Justificación CPC",
    "Justificación Correspondencia", "Observaciones", "Revisiòn PC",
)


def norm(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def as_json(value: object) -> str:
    return json.dumps(value, ensure_ascii=False, default=str)


@dataclass(frozen=True)
class Catalog:
    exceptions: frozenset[str]
    service_ciiu: frozenset[str]
    service_cpc: frozenset[str]
    combinations: frozenset[str]
    cpc_names: dict[str, str]
    exception_names: dict[str, str]


@lru_cache(maxsize=1)
def get_catalog() -> Catalog:
    book = load_workbook(CATALOG_PATH, read_only=True, data_only=True)
    rules = book["CIIU CPC validaciones"]
    exceptions, ciiu, cpc, exception_names = set(), set(), set(), {}
    for row in rules.iter_rows(min_row=3, values_only=True):
        exception, description, group, product = norm(row[1]), norm(row[2]), norm(row[5]), norm(row[8])
        if exception:
            exceptions.add(exception.upper())
            exception_names[exception.upper()] = description
        if group:
            ciiu.add(group.zfill(4))
        if product:
            cpc.add(product.zfill(4))
    combinations, cpc_names = set(), {}
    for row in book["cpc ciiu"].iter_rows(min_row=2, values_only=True):
        combined, product, name = norm(row[5]), norm(row[6]), norm(row[7])
        if re.fullmatch(r"\d{8}", combined):
            combinations.add(combined)
        if product and name and name != "#N/A":
            cpc_names.setdefault(product.zfill(4), name)
    book.close()
    return Catalog(frozenset(exceptions), frozenset(ciiu), frozenset(cpc),
                   frozenset(combinations), cpc_names, exception_names)


def parse_upload(content: bytes) -> tuple[list[dict], list[str], list[str]]:
    """Importa la hoja de servicios; las fórmulas auxiliares se recalculan en la app."""
    book = load_workbook(BytesIO(content), read_only=True, data_only=True)
    sheet = None
    for candidate in book:
        headings = [norm(c.value) for c in next(candidate.iter_rows(min_row=1, max_row=1, max_col=54))]
        if set(SOURCE_FIELDS).issubset(headings):
            sheet = candidate
            break
    if sheet is None:
        book.close()
        raise ValueError("No encontré una hoja con las variables de Servicios requeridas.")
    headings = [norm(c.value) for c in next(sheet.iter_rows(min_row=1, max_row=1, max_col=54))]
    present_headings = [x for x in headings if x]
    if len(present_headings) != len(set(present_headings)):
        raise ValueError("Hay encabezados duplicados en A:BB; revise el archivo.")
    records, warnings, seen = [], [], set()
    for row_number, cells in enumerate(sheet.iter_rows(min_row=2, max_col=54, values_only=True), 2):
        if all(v is None or norm(v) == "" for v in cells):
            continue
        source = {key: value for key, value in zip(headings, cells) if key}
        company = norm(source.get("inec_identificador_empresa"))
        service_id = norm(source.get("id_serv"))
        if not company:
            raise ValueError(f"Fila {row_number}: falta inec_identificador_empresa.")
        if not service_id:
            service_id = f"SIN_ID_FILA_{row_number}"
            warnings.append(f"Fila {row_number}: id_serv vacío. Se usó un identificador provisional para esta carga.")
        key = (company, service_id)
        if key in seen:
            raise ValueError(f"Fila {row_number}: id_serv {service_id} repetido para la empresa {company}.")
        seen.add(key)
        records.append({"company_id": company, "service_id": service_id,
                        "source_row": row_number, "source": source})
    book.close()
    if not records:
        raise ValueError("La hoja de Servicios no contiene registros.")
    return records, headings, warnings


def is_other(source: dict) -> bool:
    return norm(source.get("servicio")).casefold() == "otros" or norm(source.get("Código CPC")) == "99999999"


def review_defaults(source: dict) -> dict:
    return {
        "edited_code": norm(source.get("cpc_act_8dígitos")),
        "review_note": norm(source.get("Revisar")),
        "reviewer": norm(source.get("Nombre del responsable")),
        "decision": norm(source.get("Decisión")) or "Pendiente",
        "just_ciiu": norm(source.get("Justificación CIIU")),
        "just_cpc": norm(source.get("Justificación CPC")),
        "just_match": norm(source.get("Justificación Correspondencia")),
        "observations": norm(source.get("Observaciones")),
        "review_pc": norm(source.get("Revisiòn PC")),
    }


def evaluate(source: dict, edited_code: str, catalog: Catalog) -> dict:
    original = norm(source.get("Código CPC"))
    code = norm(edited_code) or original
    activities = [norm(source.get(field)).upper().replace(".", "") for field in
                  ("ciiu4_actividad_principal", "ciiu4_actividad_secundaria", "ciiu4_tercera_actividad")]
    activities = [a for a in activities if a]
    other = is_other(source)
    if other:
        return {"other": True, "code": code, "ciiu4": "", "cpc4": "", "cpc_name": "",
                "code_ok": True, "ciiu_ok": True, "cpc_ok": True, "pair_ok": True,
                "activity_match": False, "service_activity": False,
                "exceptions": [], "excluded_divisions": [], "alerts": []}
    formatted = bool(re.fullmatch(r"\d{8}", code))
    ciiu4, cpc4 = (code[:4], code[4:]) if formatted else ("", "")
    exceptions = [a for a in activities if a in catalog.exceptions]
    excluded = [a for a in activities if a[:3] in EXCLUDED_DIVISIONS]
    service_activity = any(a[0] in SERVICE_LETTERS for a in activities)
    match = formatted and any(ciiu4 == a[1:5] for a in activities if len(a) >= 5)
    ciiu_ok = formatted and ciiu4 in catalog.service_ciiu
    cpc_ok = formatted and cpc4 in catalog.service_cpc
    pair_ok = formatted and code in catalog.combinations
    alerts = []
    if not formatted:
        alerts.append("Código de servicio: se requieren 8 dígitos")
    else:
        if code == "99999999":
            alerts.append("99999999 se reserva para la fila Otros")
        if not ciiu_ok:
            alerts.append("CIIU del código fuera de servicios o excepciones")
        if not cpc_ok:
            alerts.append("CPC fuera de servicios o excepciones")
        if not pair_ok:
            alerts.append("Combinación CIIU–CPC no encontrada en el catálogo (revisar)")
    if not service_activity and not exceptions:
        alerts.append("Ninguna actividad de servicio o excepción identificada")
    if excluded:
        alerts.append("Actividad en división no investigada: " + ", ".join(excluded))
    if not match:
        alerts.append("CIIU del servicio no coincide con las actividades")
    return {"other": False, "code": code, "ciiu4": ciiu4, "cpc4": cpc4,
            "cpc_name": catalog.cpc_names.get(cpc4, "Sin descripción en el catálogo"),
            "code_ok": formatted, "ciiu_ok": ciiu_ok, "cpc_ok": cpc_ok,
            "pair_ok": pair_ok, "activity_match": match,
            "service_activity": service_activity, "exceptions": exceptions,
            "excluded_divisions": excluded, "alerts": alerts}


def summarize_company(items: list[dict]) -> dict:
    active = [r for r in items if not r["validation"]["other"]]
    count = sum(bool(r["validation"]["activity_match"]) for r in active)
    return {"total": len(active), "matches": count,
            "minimum_met": count > 0 if active else None,
            "reviewed": sum(r["decision"] in ("Aceptado", "Cambio") for r in active)}
