"""Vista Streamlit del análisis conciliado de riesgos materializados."""

import pandas as pd
import streamlit as st

from config.risk_catalog import EXCLUSION_CATEGORIES, RISK_CATALOG
from repositories.incident_risk_repository import (
    available_incident_years, fetch_available_problems, fetch_classification_overrides,
    fetch_incidents_period, fetch_risk_problem_link_history, fetch_risk_problem_links,
    remove_risk_problem_link, save_risk_problem_link,
)
from services.incident_risk_service import build_analysis, classify_incidents
from utils.risk_excel_export import build_risk_workbook

MONTHS = {1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",11:"Noviembre",12:"Diciembre"}


def render_riesgos_materializados():
    st.divider(); st.subheader("Análisis conciliado de riesgos materializados")
    years = available_incident_years()
    if not years:
        st.info("No hay años disponibles en la fuente de incidentes."); return
    current = pd.Timestamp.now(tz="America/Bogota").year
    default = years.index(current) if current in years else 0
    cols = st.columns([1,1,1])
    year = cols[0].selectbox("Año", years, index=default, key="risk_year")
    whole_year = cols[1].checkbox("Todo el año", value=True, key="risk_whole_year")
    if whole_year:
        month_from, month_to = 1, 12
        cols[2].caption("Periodo: enero a diciembre")
    else:
        month_cols = st.columns(2)
        month_from = month_cols[0].selectbox("Mes desde", MONTHS, format_func=MONTHS.get, key="risk_month_from")
        month_to = month_cols[1].selectbox("Mes hasta", MONTHS, index=11, format_func=MONTHS.get, key="risk_month_to")
        if month_from > month_to:
            st.error("Mes desde no puede ser posterior a Mes hasta."); return
    with st.spinner("Consultando y clasificando incidentes..."):
        source = fetch_incidents_period(year, month_from, month_to)
        if source.empty:
            st.info("No hay incidentes en el periodo seleccionado."); return
        overrides = fetch_classification_overrides(source["numero"].tolist())
        detail = classify_incidents(source, overrides)
        problem_links = fetch_risk_problem_links()
    fcols = st.columns(5)
    risk_filter = fcols[0].selectbox("Riesgo", ["Todos"] + [r["risk_id"] for r in RISK_CATALOG])
    owner_filter = fcols[1].selectbox("Dueño del riesgo", ["Todos"] + sorted({r["owner"] for r in RISK_CATALOG}))
    status_filter = fcols[2].selectbox("Estado de clasificación", ["Todos", "RISK", "EXCLUSION", "PENDING"])
    exclusion_filter = fcols[3].selectbox("Categoría de exclusión", ["Todos"] + list(EXCLUSION_CATEGORIES))
    search = fcols[4].text_input("Buscar incidente", placeholder="INC...").strip().upper()
    filtered = detail.copy()
    if risk_filter != "Todos": filtered = filtered[filtered["risk_id"] == risk_filter]
    if owner_filter != "Todos":
        ids = [r["risk_id"] for r in RISK_CATALOG if r["owner"] == owner_filter]; filtered = filtered[filtered["risk_id"].isin(ids)]
    if status_filter != "Todos": filtered = filtered[filtered["classification_type"] == status_filter]
    if exclusion_filter != "Todos": filtered = filtered[filtered["exclusion_category"] == exclusion_filter]
    if search: filtered = filtered[filtered["numero"].str.contains(search, regex=False)]
    analysis = build_analysis(filtered, range(month_from, month_to + 1), problem_links)
    analysis["link_history"] = fetch_risk_problem_link_history()
    validation = analysis["validation"]
    cards = st.columns(5)
    values = [("Total incidentes", validation["total"]), ("Riesgos diferentes", filtered.loc[filtered["classification_type"] == "RISK", "risk_id"].nunique()),
              ("Materializados", validation["risk"]), ("Exclusiones", validation["exclusion"]), ("Pendientes", validation["pending"])]
    for col, (label, value) in zip(cards, values): col.metric(label, value)
    event_count = int(analysis["events"]["event_id"].nunique()) if not analysis["events"].empty else 0
    treated_risks = int(analysis["risks"]["Problemas asociados"].fillna("").ne("").sum()) if not analysis["risks"].empty else 0
    untreated_recurrent = int(((analysis["risks"]["Estado"] == "🔥 REINCIDENTE") & analysis["risks"]["Problemas asociados"].fillna("").eq("")).sum()) if not analysis["risks"].empty else 0
    executive_cards = st.columns(3)
    executive_cards[0].metric("Eventos materiales estimados", event_count)
    executive_cards[1].metric("Riesgos con problema", treated_risks)
    executive_cards[2].metric("Reincidentes sin problema", untreated_recurrent)
    if validation["reconciled"]: st.success("100% conciliado")
    else:
        st.error("Error de conciliación de incidentes")
        if not validation["conflicts"].empty: st.dataframe(validation["conflicts"], hide_index=True)
    st.markdown("#### Riesgos materializados")
    if analysis["risks"].empty: st.info("No hay incidentes asociados a riesgos con los filtros actuales.")
    else:
        main_cols = ["ID","Riesgo Materializado","Dueño del Riesgo","Impacto Escala","Cantidad tickets asociados","Tickets Asociados","Asignación Operativa RACI","Estado"] + [c for c in ("Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic") if c in analysis["risks"]]
        st.dataframe(analysis["risks"][main_cols], use_container_width=True, hide_index=True)
        with st.expander("Ver detalle completo de tickets por riesgo"):
            for _, row in analysis["risks"].iterrows():
                st.write(f"**{row['ID']} · {row['Estado']}** — {row['Tickets Asociados']}")
    st.markdown("#### Naturaleza y causa raíz de los incidentes")
    st.caption("Esta vista permite seleccionar un problema coherente con la naturaleza del evento. La causa confirmada prevalece sobre la causa automática; cuando no existe evidencia suficiente se indica que continúa en análisis.")
    cause_columns = {
        "risk_id": "Riesgo", "tickets": "INC asociados", "incident_nature": "Naturaleza de los INC",
        "component": "Componente general", "root_cause": "Causa raíz", "root_cause_status": "Estado de la causa",
        "event_status": "Nivel de recurrencia",
    }
    event_causes = analysis["events"][[column for column in cause_columns if column in analysis["events"]]].rename(columns=cause_columns)
    if not event_causes.empty:
        event_causes["Causa raíz"] = event_causes["Causa raíz"].replace("", "En proceso de análisis")
    st.dataframe(event_causes, use_container_width=True, hide_index=True)
    st.markdown("#### Tratamiento mediante problemas")
    treatment_columns = ["ID", "Naturaleza consolidada de los INC", "Causa raíz consolidada", "Estado causa raíz", "Eventos reales", "Problemas asociados", "Estado del problema"]
    treatment = analysis["risks"][treatment_columns].copy() if not analysis["risks"].empty else pd.DataFrame(columns=treatment_columns)
    st.dataframe(treatment, use_container_width=True, hide_index=True)
    if untreated_recurrent:
        st.warning(f"Hay {untreated_recurrent} riesgos reincidentes sin problema asociado ni tratamiento trazable.")
    with st.expander("Ver eventos materiales agrupados"):
        st.caption("Agrupación automática inicial: mismo riesgo y componente dentro de una ventana de 6 horas. Debe revisarse cuando exista evidencia de eventos distintos.")
        st.dataframe(analysis["events"], use_container_width=True, hide_index=True)
    with st.expander("Vincular o retirar un problema (usuarios autorizados)"):
        active_risks = sorted(filtered.loc[filtered["classification_type"] == "RISK", "risk_id"].dropna().unique())
        problems = fetch_available_problems()
        if not active_risks:
            st.info("No hay riesgos materializados con los filtros actuales.")
        elif problems.empty:
            st.info("No hay problemas registrados. Créalo primero desde el módulo Problemas.")
        else:
            risk_choice = st.selectbox("Riesgo materializado", active_risks, key="link_risk")
            problem_options = problems["numero"].astype(str).tolist()
            labels = dict(zip(problems["numero"].astype(str), problems["declaracion_problema"].fillna("").astype(str)))
            problem_choice = st.selectbox("Problema existente", problem_options, key="link_problem",
                                          format_func=lambda number: f"{number} · {labels.get(number, '')}")
            link_note = st.text_area("Justificación de la asociación", key="link_problem_note")
            col_link, col_unlink = st.columns(2)
            if col_link.button("Vincular problema", type="primary"):
                try:
                    save_risk_problem_link(risk_choice, problem_choice, link_note, st.session_state.get("user"))
                except (PermissionError, ValueError) as error:
                    st.error(str(error))
                else:
                    st.success("Problema vinculado con trazabilidad de usuario y fecha."); st.rerun()
            if col_unlink.button("Retirar vínculo"):
                try:
                    remove_risk_problem_link(risk_choice, problem_choice, st.session_state.get("user"))
                except (PermissionError, ValueError) as error:
                    st.error(str(error))
                else:
                    st.success("Vínculo retirado y registrado en el historial."); st.rerun()
    st.markdown("#### Conciliación"); st.dataframe(analysis["reconciliation"], use_container_width=True, hide_index=True)
    st.markdown("#### Exclusiones"); st.dataframe(analysis["exclusions"], use_container_width=True, hide_index=True)
    st.markdown("#### Detalle explicable por incidente")
    detail_cols = ["numero","creado_dt","mes","breve_descripcion","categoria","servicio_negocio","risk_id","classification_type","exclusion_category","classification_reason","confidence","estado","grupo_asignacion","prioridad"]
    st.dataframe(filtered[[c for c in detail_cols if c in filtered]], use_container_width=True, hide_index=True)
    if validation["reconciled"]:
        data = build_risk_workbook(analysis, year, month_from, month_to)
        suffix = str(year) if whole_year else f"{year}_{month_from:02d}-{month_to:02d}"
        st.download_button("📥 Descargar análisis en Excel", data=data, file_name=f"Riesgos_Materializados_{suffix}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
