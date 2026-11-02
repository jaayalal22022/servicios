"""Tablero ENESEM Tomo II: revisión de Servicios."""

from collections import defaultdict
import os

import pandas as pd
import streamlit as st

from engine import evaluate, get_catalog, is_other, norm, parse_upload, summarize_company
from exporter import export_excel
from storage import (ConflictError, SourceChangedError, import_records, init_db,
                     list_datasets, load_audit, load_records, local_url, save_review)


st.set_page_config(page_title="ENESEM · Servicios", layout="wide", page_icon="🔎")


def database_url() -> str:
    value = os.environ.get("DATABASE_URL", "")
    if value:
        return value
    try:
        return st.secrets.get("DATABASE_URL", "") or local_url()
    except (FileNotFoundError, KeyError):
        return local_url()


def frame(items: list[dict]) -> pd.DataFrame:
    rows = []
    for item in items:
        src, val = item["source"], item["validation"]
        rows.append({
            "Empresa": item["company_id"], "id_serv": item["service_id"],
            "Tipo": norm(src.get("servicio")),
            "Servicio informado": norm(src.get("Descripción del servicio")),
            "CPC original": norm(src.get("Código CPC")), "CPC vigente": val["code"],
            "CPC4": val["cpc4"], "Descripción CPC": val["cpc_name"],
            "CIIU coincide": "Sí" if val["activity_match"] else "No",
            "Alertas": " | ".join(val["alerts"]) or "Sin alertas",
            "Decisión": item["decision"], "Responsable": item["reviewer"],
        })
    return pd.DataFrame(rows)


def record_picker(items: list[dict], key: str) -> dict:
    chosen = st.selectbox(
        "Elegir servicio para revisar",
        items,
        format_func=lambda r: (f"{r['company_id']} · {r['service_id']} · "
                               f"{norm(r['source'].get('Descripción del servicio'))[:90]}"),
        key=key,
    )
    return chosen


def review_form(item: dict, db_url: str, catalog, prefix: str):
    src, current = item["source"], item["validation"]
    st.caption(f"Empresa {item['company_id']} · id_serv {item['service_id']} · revisión {item['revision']}")
    st.write("**Actividades de la empresa**")
    st.write("Principal:", norm(src.get("ciiu4_actividad_principal")), "·", norm(src.get("desc_actividad_principal")))
    st.write("Secundaria:", norm(src.get("ciiu4_actividad_secundaria")), "·", norm(src.get("desc_actividad_secundaria")))
    st.write("Tercera:", norm(src.get("ciiu4_tercera_actividad")), "·", norm(src.get("desc_tercera_actividad")))
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("CPC original", norm(src.get("Código CPC")))
    col2.metric("Código vigente", current["code"])
    col3.metric("Correspondencia CIIU", "Sí" if current["activity_match"] else "No")
    col4.metric("Alertas", len(current["alerts"]))
    st.caption(f"CPC4 {current['cpc4'] or '—'} · {current['cpc_name'] or '—'}")
    if current["exceptions"]:
        st.info("Actividades de excepción: " + ", ".join(current["exceptions"]))
    if current["alerts"]:
        for alert in current["alerts"]:
            st.warning(alert)
    else:
        st.success("Sin alertas automáticas para este servicio.")
    with st.form(f"review_{prefix}_{item['id']}"):
        edited = st.text_input("CPC corregido (8 dígitos; el original se conserva)",
                               value=item["edited_code"], key=f"code_{prefix}_{item['id']}")
        choice = st.selectbox("Decisión", ["Pendiente", "Aceptado", "Cambio"],
                              index=["Pendiente", "Aceptado", "Cambio"].index(item["decision"])
                              if item["decision"] in ("Pendiente", "Aceptado", "Cambio") else 0,
                              key=f"decision_{prefix}_{item['id']}")
        reviewer = st.text_input("Nombre del responsable", value=item["reviewer"] or st.session_state.get("my_name", ""))
        note = st.text_area("Revisar / acción necesaria", value=item["review_note"])
        c1, c2 = st.columns(2)
        just_ciiu = c1.text_area("Justificación CIIU", value=item["just_ciiu"])
        just_cpc = c2.text_area("Justificación CPC", value=item["just_cpc"])
        just_match = st.text_area("Justificación de correspondencia", value=item["just_match"])
        observations = st.text_area("Observaciones", value=item["observations"])
        review_pc = st.text_input("Revisión PC", value=item["review_pc"])
        submitted = st.form_submit_button("Guardar revisión", type="primary")
    if not submitted:
        return
    if not reviewer.strip():
        st.error("Indica el nombre del responsable antes de guardar.")
        return
    candidate = evaluate(src, edited, catalog)
    if edited.strip() and not candidate["code_ok"]:
        st.error("El CPC corregido debe contener exactamente ocho dígitos.")
        return
    if edited.strip() == "99999999" and not is_other(src):
        st.error("99999999 se reserva para la fila 'Otros'.")
        return
    if choice == "Cambio" and (not edited.strip() or edited.strip() == norm(src.get("Código CPC"))):
        st.error("Para marcar Cambio, escribe un CPC distinto del original.")
        return
    if choice in ("Aceptado", "Cambio") and candidate["alerts"] and not any(
        x.strip() for x in (just_ciiu, just_cpc, just_match, observations)
    ):
        st.error("Hay alertas. Registra una justificación u observación para cerrar la revisión.")
        return
    changes = {
        "edited_code": edited.strip(), "review_note": note.strip(), "reviewer": reviewer.strip(),
        "decision": choice, "just_ciiu": just_ciiu.strip(), "just_cpc": just_cpc.strip(),
        "just_match": just_match.strip(), "observations": observations.strip(),
        "review_pc": review_pc.strip(),
    }
    try:
        save_review(db_url, item["id"], item["revision"], changes)
    except ConflictError as exc:
        st.error(str(exc))
        return
    st.session_state.pop("download_bytes", None)
    st.success("Revisión guardada.")
    st.rerun()


st.title("ENESEM · Tomo II · Servicios")
st.caption("Validación CIIU–CPC · fuente original intacta · 'Otros' fuera de la revisión")

url = database_url()
try:
    init_db(url)
    catalog = get_catalog()
except Exception as exc:
    st.error(f"No se pudo abrir la base de datos o el catálogo: {exc}")
    st.stop()

with st.sidebar:
    st.header("Datos de revisión")
    st.text_input("Tu nombre", key="my_name", help="Se registra junto a cada decisión; la autenticación se configura para producción.")
    if url.startswith("sqlite:///"):
        st.info("Modo local de prueba. Para trabajo compartido, configura DATABASE_URL de Supabase.")
    else:
        st.success("Guardado compartido en PostgreSQL")
    with st.expander("Cargar un Excel de Servicios", expanded=not bool(list_datasets(url))):
        label = st.text_input("Ronda o lote", value="ENESEM 2025")
        upload = st.file_uploader("Base de Servicios (.xlsx)", type=["xlsx"])
        if st.button("Incorporar base", disabled=upload is None):
            try:
                content = upload.getvalue()
                parsed, headings, warnings = parse_upload(content)
                dataset, added, unchanged = import_records(url, label, upload.name,
                                                            content, headings, parsed)
                st.session_state["dataset_id"] = dataset["id"]
                st.session_state["import_notice"] = f"{added} servicios incorporados; {unchanged} ya existían."
                st.session_state["import_warnings"] = warnings
                st.session_state.pop("download_bytes", None)
                st.rerun()
            except (ValueError, SourceChangedError) as exc:
                st.error(str(exc))
    datasets = list_datasets(url)
    if not datasets:
        st.info("Carga una base para iniciar.")
        st.stop()
    options = {d["id"]: d for d in datasets}
    selected = st.selectbox("Ronda o lote abierto", list(options),
                            index=list(options).index(st.session_state["dataset_id"])
                            if st.session_state.get("dataset_id") in options else 0,
                            format_func=lambda x: options[x]["label"])
    dataset = options[selected]
    st.caption(f"Fuente inicial: {dataset['filename']}")

if st.session_state.get("export_dataset_id") != dataset["id"]:
    st.session_state.pop("download_bytes", None)

notice = st.session_state.pop("import_notice", None)
if notice:
    st.success(notice)
for warning in st.session_state.pop("import_warnings", []):
    st.warning(warning)

items = load_records(url, dataset["id"])
for item in items:
    item["validation"] = evaluate(item["source"], item["edited_code"], catalog)
active = [r for r in items if not r["validation"]["other"]]
if not active:
    st.info("La base no contiene servicios evaluables; las filas 'Otros' se excluyen.")
    st.stop()

firms = defaultdict(list)
for row in active:
    firms[row["company_id"]].append(row)
minimums = {k: summarize_company(v) for k, v in firms.items()}
total_alerts = sum(bool(r["validation"]["alerts"]) for r in active)
c1, c2, c3, c4 = st.columns(4)
c1.metric("Servicios evaluables", len(active))
c2.metric("Empresas", len(firms))
c3.metric("Servicios revisados", sum(r["decision"] in ("Aceptado", "Cambio") for r in active))
c4.metric("Con alertas", total_alerts)
st.caption(f"{sum(x['minimum_met'] for x in minimums.values())} de {len(firms)} empresas cumplen la correspondencia mínima de un servicio.")

tab_s, tab_e, tab_c = st.tabs(["Por servicio", "Por empresa", "Por código CPC"])
with tab_s:
    f1, f2, f3 = st.columns(3)
    company_q = f1.text_input("Buscar empresa o descripción", key="service_search")
    filter_alert = f2.selectbox("Alertas", ["Todas", "Con alertas", "Sin alertas"])
    filter_decision = f3.selectbox("Decisión", ["Todas", "Pendiente", "Aceptado", "Cambio"])
    visible = [r for r in active if
               (not company_q or company_q.casefold() in (r["company_id"] + " " + norm(r["source"].get("Descripción del servicio"))).casefold())
               and (filter_alert == "Todas" or bool(r["validation"]["alerts"]) == (filter_alert == "Con alertas"))
               and (filter_decision == "Todas" or r["decision"] == filter_decision)]
    st.dataframe(frame(visible), hide_index=True, use_container_width=True, height=380)
    if visible:
        review_form(record_picker(visible, "service_picker"), url, catalog, "service")
    else:
        st.info("No hay servicios para estos filtros.")

with tab_e:
    company = st.selectbox("Empresa", sorted(firms), key="company_picker")
    summary = minimums[company]
    st.subheader(f"Empresa {company}")
    st.write(f"Servicios con correspondencia: **{summary['matches']} de {summary['total']}** · "
             f"mínimo: **{'Cumple' if summary['minimum_met'] else 'No cumple'}** · "
             f"revisados: **{summary['reviewed']}**")
    st.dataframe(frame(firms[company]), hide_index=True, use_container_width=True)
    review_form(record_picker(firms[company], "company_record_picker"), url, catalog, "company")

with tab_c:
    grouped = defaultdict(list)
    for row in active:
        grouped[row["validation"]["cpc4"] or "Sin código válido"].append(row)
    group_rows = [{"CPC4": code, "Descripción": catalog.cpc_names.get(code, "Sin descripción"),
                   "Servicios": len(group), "Empresas": len({r["company_id"] for r in group}),
                   "Con alertas": sum(bool(r["validation"]["alerts"]) for r in group)}
                  for code, group in sorted(grouped.items())]
    st.dataframe(pd.DataFrame(group_rows), hide_index=True, use_container_width=True, height=340)
    code = st.selectbox("Explorar CPC de cuatro dígitos", sorted(grouped), key="cpc_picker")
    st.caption(f"{code} · {catalog.cpc_names.get(code, 'Sin descripción en el catálogo')}")
    st.dataframe(frame(grouped[code]), hide_index=True, use_container_width=True)
    review_form(record_picker(grouped[code], "cpc_record_picker"), url, catalog, "cpc")

with st.expander("Descargar avances e historial"):
    st.write("El archivo incluye los campos de la matriz, validaciones recalculadas, una vista de revisión y el historial de cambios.")
    if st.button("Preparar Excel actualizado"):
        st.session_state["download_bytes"] = export_excel(dataset, items, load_audit(url, dataset["id"]))
        st.session_state["export_dataset_id"] = dataset["id"]
    if "download_bytes" in st.session_state:
        st.download_button("Descargar Excel", data=st.session_state["download_bytes"],
                           file_name=f"servicios_revisados_{dataset['label'].replace(' ', '_')}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
