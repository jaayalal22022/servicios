"""Exportación de la revisión preservando los campos informados."""

from io import BytesIO
import json

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

from engine import norm, summarize_company

VALIDATION_COLUMNS = (
    "Val CIIU servicios o excepción", "Val CPC servicios o excepción",
    "Val CIIU letra primera segunda tercera actividad",
    "Val CIIU divisiones primera segunda tercera actividad",
    "Ciiu excepción actividad principal", "Ciiu excepción actividad secundaria",
    "Ciiu excepción tercera actividad", "Al menos una actividad de la  CIIU es de excepción",
    "Verif_cpc", "Correspondencia=1",
    "Número de servicios que tienen correspondencia por empresa",
    "Al menos un servicio tiene correspondencia",
)


def export_excel(dataset: dict, items: list[dict], audit: list[dict]) -> bytes:
    headings = json.loads(dataset["headers_json"])
    wb = Workbook()
    base = wb.active
    base.title = "Servicios revisados"
    base.append(headings)
    by_company = {}
    for item in items:
        by_company.setdefault(item["company_id"], []).append(item)
    summaries = {company: summarize_company(rows) for company, rows in by_company.items()}
    details = wb.create_sheet("Revision_app")
    details.append(["empresa", "id_serv", "tipo_servicio", "descripción", "código_original",
                    "código_corregido", "código_vigente", "CIIU4", "CPC4", "descripción_CPC",
                    "CIIU_admitido", "CPC_admitido", "par_en_catálogo", "coincide_actividad",
                    "actividades_excepción", "divisiones_no_investigadas", "alertas",
                    "correspondencias_empresa", "cumple_mínimo_empresa", "decisión", "responsable"])
    for item in items:
        src, result = item["source"], item["validation"]
        row = dict(src)
        row.update({
            "cpc_act_8dígitos": item["edited_code"],
            "Revisar": item["review_note"],
            "Nombre del responsable": item["reviewer"],
            "Decisión": item["decision"],
            "Justificación CIIU": item["just_ciiu"],
            "Justificación CPC": item["just_cpc"],
            "Justificación Correspondencia": item["just_match"],
            "Observaciones": item["observations"],
            "Revisiòn PC": item["review_pc"],
            "extrae_cpc": norm(src.get("Código CPC"))[-4:],
            "Descripcion_cpc": result["cpc_name"],
            "id_emp": item["company_id"],
            "CPC_FINAL": result["code"],
            "CIIU_INICIAL_ 4 dig": result["ciiu4"],
            "CPC_FINAL_ 4 dig": result["cpc4"],
            "descripcion": result["cpc_name"],
        })
        if not result["other"]:
            summary = summaries[item["company_id"]]
            row.update({
                "Val CIIU servicios o excepción": "" if result["ciiu_ok"] else "CIIU no corresponde a servicios o excepción",
                "Val CPC servicios o excepción": "" if result["cpc_ok"] else "CPC no corresponde a servicios o excepción",
                "Val CIIU letra primera segunda tercera actividad": "" if result["service_activity"] else "CIIU no corresponde a servicios",
                "Val CIIU divisiones primera segunda tercera actividad": "Excepción división CIIU" if result["excluded_divisions"] else "",
                "Ciiu excepción actividad principal": "CIIU de excepción" if norm(src.get("ciiu4_actividad_principal")).upper().replace(".", "") in result["exceptions"] else "",
                "Ciiu excepción actividad secundaria": "CIIU de excepción" if norm(src.get("ciiu4_actividad_secundaria")).upper().replace(".", "") in result["exceptions"] else "",
                "Ciiu excepción tercera actividad": "CIIU de excepción" if norm(src.get("ciiu4_tercera_actividad")).upper().replace(".", "") in result["exceptions"] else "",
                "Al menos una actividad de la  CIIU es de excepción": "CIIU de excepción" if result["exceptions"] else "",
                "Verif_cpc": "" if result["activity_match"] else "CPC no corresponde a actividades",
                "Correspondencia=1": 1 if result["activity_match"] else 2,
                "Número de servicios que tienen correspondencia por empresa": summary["matches"],
                "Al menos un servicio tiene correspondencia": "Al menos un servicio tiene correspondencia entre CIIU CPC" if summary["minimum_met"] else "",
            })
        else:
            for col in VALIDATION_COLUMNS:
                row[col] = ""
        base.append([row.get(h) for h in headings])
        if result["other"]:
            continue
        summary = summaries[item["company_id"]]
        details.append([item["company_id"], item["service_id"], norm(src.get("servicio")),
                        norm(src.get("Descripción del servicio")), norm(src.get("Código CPC")),
                        item["edited_code"], result["code"], result["ciiu4"], result["cpc4"],
                        result["cpc_name"], result["ciiu_ok"], result["cpc_ok"], result["pair_ok"],
                        result["activity_match"], ", ".join(result["exceptions"]),
                        ", ".join(result["excluded_divisions"]), " | ".join(result["alerts"]),
                        summary["matches"], summary["minimum_met"], item["decision"], item["reviewer"]])
    history = wb.create_sheet("Historial")
    history.append(["empresa", "id_serv", "fecha_UTC", "responsable", "antes", "después"])
    for entry in audit:
        history.append([entry["company_id"], entry["service_id"], entry["changed_at"],
                        entry["changed_by"], entry["previous_json"], entry["new_json"]])
    for sheet in wb:
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        sheet.row_dimensions[1].height = 36
        for cell in sheet[1]:
            cell.fill = PatternFill("solid", fgColor="17324D")
            cell.font = Font(color="FFFFFF", bold=True)
            cell.alignment = Alignment(wrap_text=True, vertical="center")
        for col in sheet.columns:
            letter = col[0].column_letter
            width = min(38, max(13, len(str(col[0].value or "")) + 3))
            sheet.column_dimensions[letter].width = width
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()
