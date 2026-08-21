"""Motor explicable y excluyente de clasificación de incidentes corporativos."""

import re
import unicodedata
import hashlib
from collections import Counter

import pandas as pd

from config.risk_catalog import EXCLUSION_CATEGORIES, RISK_BY_ID, RISK_CATALOG, recurrence_status

TEXT_FIELDS = ("categoria", "servicio_negocio", "tipo_falla", "breve_descripcion", "descripcion",
               "grupo_asignacion", "observaciones_trabajo", "observaciones_adicionales", "actualizaciones",
               "lista_notas_trabajo", "impacto", "tipificacion_original", "causa_raiz_original",
               "tipificacion_auto", "tipo_incidente_auto", "causa_raiz_auto", "es_alerta_auto")

EXCLUSION_RULES = (
    ("Pruebas y Simulacros", ("prueba noc", "prueba interna", "simulacro", "prueba controlada", "ticket generado intencionalmente")),
    ("Falsas Alarmas (Monitoreo NOC/SOC)", ("falsa alarma", "sin afectacion", "alerta normalizada", "alarma sin indisponibilidad", "intermitencia sin afectacion")),
    ("Eventos de Seguridad Contenidos", ("evento contenido", "amenaza contenida", "incidente contenido")),
    ("Requerimientos y Soporte Rutinario (BAU)", ("olvido de contrasena", "instalacion normal", "configuracion de usuario", "error de usuario", "acompanamiento", "consulta", "solicitud operativa", "soporte rutinario")),
)


def normalize_text(value):
    value = "" if value is None or (not isinstance(value, (list, tuple, dict)) and pd.isna(value)) else str(value)
    value = unicodedata.normalize("NFKD", value).encode("ascii", "ignore").decode().casefold()
    return re.sub(r"\s+", " ", value).strip()


def normalize_incident_number(value):
    return re.sub(r"\s+", "", str(value or "").upper())


def _field_texts(incident):
    return {field: normalize_text(incident.get(field, "")) for field in TEXT_FIELDS}


def _evidence(fields, terms):
    return [(field, term) for field, text in fields.items() for term in terms if term in text]


def _risk_candidate(risk, fields, all_text):
    if any(term in all_text for term in risk.get("exclusions", ())):
        return None
    rules = risk["classification_rules"]
    strong = [term for term in rules.get("strong", ()) if term in all_text]
    product_terms = tuple(rules.get("product", ())) + tuple(rules.get("technical", ()))
    product = _evidence(fields, product_terms)
    failure = _evidence(fields, rules.get("failure", ()))
    distinct_fields = len({field for field, _ in product + failure})
    if not strong and (not product or not failure or distinct_fields < 2):
        return None
    score = min(0.99, 0.72 + 0.05 * min(len(product) + len(failure), 4) + (0.12 if strong else 0))
    evidence = []
    for field, term in product + failure:
        label = f"{field}: {term}"
        if label not in evidence:
            evidence.append(label)
    if strong:
        evidence.insert(0, f"evidencia fuerte: {strong[0]}")
    return {"risk": risk, "confidence": round(score, 2), "evidence": evidence[:4]}


def classify_incident(incident):
    number = normalize_incident_number(incident.get("numero", incident.get("incident_number", "")))
    fields = _field_texts(incident)
    all_text = " | ".join(fields.values())
    material_terms = ("caida", "caido", "indisponibilidad", "no responde", "degradacion", "lentitud", "timeout", "error 500", "error 503", "error 504", "falla", "no genero alerta", "no detecto")
    bau_terms = ("instalacion", "activacion", "configuracion de usuario", "nuevo pc", "guia de uso", "tramite", "orden")
    alert_terms = ("alerta", "alarma", "monitoreo", "noc")
    has_material_evidence = any(term in all_text for term in material_terms)
    if any(term in all_text for term in bau_terms) and not has_material_evidence:
        return {"incident_number": number, "classification_type": "EXCLUSION", "risk_id": "",
                "exclusion_category": "Requerimientos y Soporte Rutinario (BAU)",
                "classification_reason": "Actividad de instalación, activación, trámite o uso sin evidencia de falla.",
                "confidence": 0.97, "classification_source": "AUTO", "candidate_risks": ""}
    if any(term in all_text for term in alert_terms) and not has_material_evidence:
        return {"incident_number": number, "classification_type": "EXCLUSION", "risk_id": "",
                "exclusion_category": "Falsas Alarmas (Monitoreo NOC/SOC)",
                "classification_reason": "Alerta o consulta de monitoreo sin afectación material confirmada.",
                "confidence": 0.95, "classification_source": "AUTO", "candidate_risks": ""}
    # Exclusiones inequívocas se evalúan primero. BAU se evalúa después de los
    # riesgos específicos para que "consulta de listas" o soporte ERP no oculten R164/R132.
    for category, phrases in EXCLUSION_RULES[:-1]:
        matches = [phrase for phrase in phrases if phrase in all_text]
        if matches:
            # Las exclusiones explícitas prevalecen; no se deducen de una palabra aislada.
            return {"incident_number": number, "classification_type": "EXCLUSION", "risk_id": "",
                    "exclusion_category": category,
                    "classification_reason": f"{category}: evidencia explícita '{matches[0]}'.",
                    "confidence": 0.96, "classification_source": "AUTO", "candidate_risks": ""}
    candidates = [candidate for risk in sorted(RISK_CATALOG, key=lambda r: r["priority"])
                  if (candidate := _risk_candidate(risk, fields, all_text))]
    if candidates:
        candidates.sort(key=lambda c: (c["risk"]["priority"], -c["confidence"]))
        selected = candidates[0]
        risk_id = selected["risk"]["risk_id"]
        return {"incident_number": number, "classification_type": "RISK", "risk_id": risk_id,
                "exclusion_category": "", "classification_reason": f"{risk_id} – " + " + ".join(selected["evidence"]),
                "confidence": selected["confidence"], "classification_source": "AUTO",
                "candidate_risks": ", ".join(c["risk"]["risk_id"] for c in candidates)}
    category, phrases = EXCLUSION_RULES[-1]
    matches = [phrase for phrase in phrases if phrase in all_text]
    if matches:
        return {"incident_number": number, "classification_type": "EXCLUSION", "risk_id": "",
                "exclusion_category": category,
                "classification_reason": f"{category}: evidencia explícita '{matches[0]}'.",
                "confidence": 0.93, "classification_source": "AUTO", "candidate_risks": ""}
    return {"incident_number": number, "classification_type": "PENDING", "risk_id": "",
            "exclusion_category": "", "classification_reason": "Evidencia insuficiente para asignar de forma confiable un riesgo o exclusión.",
            "confidence": 0.0, "classification_source": "AUTO", "candidate_risks": ""}


def normalize_incidents(df):
    work = df.copy()
    for column in TEXT_FIELDS:
        if column not in work:
            work[column] = ""
        work[column] = work[column].fillna("").astype(str).str.strip()
    work["numero"] = work.get("numero", pd.Series(dtype=str)).apply(normalize_incident_number)
    work["creado_dt"] = pd.to_datetime(work.get("creado"), errors="coerce")
    work = work[work["numero"].ne("")].copy()
    # La PK debería garantizarlo; ante un origen anómalo se conserva un registro lógico por INC.
    return work.sort_values("creado_dt", na_position="last").drop_duplicates("numero", keep="last")


def apply_overrides(detail, overrides=None):
    detail = detail.copy()
    detail["automatic_classification_type"] = detail["classification_type"]
    detail["automatic_risk_id"] = detail["risk_id"]
    detail["automatic_exclusion_category"] = detail["exclusion_category"]
    if overrides is None or overrides.empty:
        detail["modified_by"] = ""; detail["modified_at"] = pd.NaT; detail["change_reason"] = ""
        return detail
    ov = overrides.drop_duplicates("incident_number", keep="last").copy()
    detail = detail.merge(ov, left_on="numero", right_on="incident_number", how="left", suffixes=("", "_override"))
    mask = detail["final_classification_type"].fillna("").isin(("RISK", "EXCLUSION", "PENDING"))
    for target, source in (("classification_type", "final_classification_type"), ("risk_id", "final_risk_id"), ("exclusion_category", "final_exclusion_category")):
        detail.loc[mask, target] = detail.loc[mask, source].fillna("")
    detail.loc[mask, "classification_source"] = "MANUAL"
    detail.loc[mask, "classification_reason"] = "Override manual: " + detail.loc[mask, "change_reason"].fillna("sin motivo registrado")
    return detail


def classify_incidents(df, overrides=None):
    work = normalize_incidents(df)
    classifications = pd.DataFrame([classify_incident(row) for row in work.to_dict("records")])
    detail = work.merge(classifications, left_on="numero", right_on="incident_number", how="left", validate="one_to_one")
    detail["mes_num"] = detail["creado_dt"].dt.month
    detail["mes"] = detail["creado_dt"].dt.strftime("%b").fillna("")
    return apply_overrides(detail, overrides)


def group_materialized_events(detail, window_hours=6):
    """Agrupa tickets próximos del mismo riesgo/componente en un evento auditable."""
    risk_rows = detail[detail["classification_type"] == "RISK"].copy()
    columns = ["event_id", "risk_id", "component", "incident_nature", "event_start", "event_end", "ticket_count",
               "tickets", "affected_clients", "event_status", "root_cause_status", "root_cause"]
    if risk_rows.empty:
        return pd.DataFrame(columns=columns)
    risk_rows["component"] = risk_rows["servicio_negocio"].fillna("").apply(normalize_text).replace("", "sin componente")
    risk_rows = risk_rows.sort_values(["risk_id", "component", "creado_dt", "numero"])
    events = []
    for (risk_id, component), group in risk_rows.groupby(["risk_id", "component"], dropna=False):
        current = []
        previous = None
        for _, row in group.iterrows():
            moment = row["creado_dt"]
            if current and (pd.isna(moment) or pd.isna(previous) or moment - previous > pd.Timedelta(hours=window_hours)):
                events.append(_event_record(risk_id, component, current))
                current = []
            current.append(row)
            previous = moment
        if current:
            events.append(_event_record(risk_id, component, current))
    result = pd.DataFrame(events, columns=columns)
    event_counts = result.groupby("risk_id")["event_id"].transform("nunique")
    result["event_status"] = event_counts.map(lambda count: "REINCIDENTE" if count >= 2 else "AISLADO")
    return result


def _event_record(risk_id, component, rows):
    frame = pd.DataFrame(rows)
    tickets = sorted(frame["numero"].unique())
    start, end = frame["creado_dt"].min(), frame["creado_dt"].max()
    seed = f"{risk_id}|{component}|{start}|{'|'.join(tickets)}"
    causes = frame.get("causa_raiz_original", pd.Series(dtype=str)).fillna("").astype(str).str.strip()
    causes = causes[~causes.str.casefold().isin(("", "sin inferencia"))]
    inferred = frame.get("causa_raiz_auto", pd.Series(dtype=str)).fillna("").astype(str).str.strip()
    inferred = inferred[~inferred.str.casefold().isin(("", "sin inferencia"))]
    root_cause = causes.iloc[0] if not causes.empty else (inferred.iloc[0] if not inferred.empty else "")
    status = "Causa confirmada" if not causes.empty else ("Causa probable" if root_cause else "Pendiente de investigación")
    nature_values = []
    for column in ("tipificacion_original", "tipificacion_auto", "tipo_incidente_auto"):
        values = frame.get(column, pd.Series(dtype=str)).fillna("").astype(str).str.strip()
        for value in values:
            if value and value.casefold() not in ("sin inferencia", "no determinado") and value not in nature_values:
                nature_values.append(value)
    incident_nature = " | ".join(nature_values[:3]) if nature_values else component
    clients = frame.get("empresa", pd.Series(dtype=str)).fillna("").astype(str).str.strip()
    return {"event_id": "EVT-" + hashlib.sha1(seed.encode()).hexdigest()[:10].upper(), "risk_id": risk_id,
            "component": component, "incident_nature": incident_nature, "event_start": start, "event_end": end, "ticket_count": len(tickets),
            "tickets": ", ".join(tickets), "affected_clients": clients[clients.ne("")].nunique(),
            "event_status": "", "root_cause_status": status, "root_cause": root_cause}


def validate_reconciliation(detail):
    conflicts = []
    duplicated = detail[detail["numero"].duplicated(False)]
    for number in duplicated["numero"].unique():
        conflicts.append({"incident": number, "conflict": "ID duplicado"})
    invalid = detail[~detail["classification_type"].isin(("RISK", "EXCLUSION", "PENDING"))]
    for _, row in invalid.iterrows():
        conflicts.append({"incident": row["numero"], "conflict": "Tipo de clasificación inválido"})
    both = detail[(detail["risk_id"].fillna("") != "") & (detail["exclusion_category"].fillna("") != "")]
    for _, row in both.iterrows():
        conflicts.append({"incident": row["numero"], "conflict": f"Riesgo {row['risk_id']} y exclusión {row['exclusion_category']}"})
    total = detail["numero"].nunique()
    risk = detail.loc[detail["classification_type"] == "RISK", "numero"].nunique()
    exclusion = detail.loc[detail["classification_type"] == "EXCLUSION", "numero"].nunique()
    pending = detail.loc[detail["classification_type"] == "PENDING", "numero"].nunique()
    reconciled = not conflicts and total == risk + exclusion + pending and len(detail) == total
    return {"total": total, "risk": risk, "exclusion": exclusion, "pending": pending,
            "reconciled": reconciled, "percentage": 1.0 if reconciled else ((risk + exclusion + pending) / total if total else 1.0),
            "conflicts": pd.DataFrame(conflicts)}


def build_analysis(detail, months, problem_links=None):
    validation = validate_reconciliation(detail)
    month_labels = {1:"Ene",2:"Feb",3:"Mar",4:"Abr",5:"May",6:"Jun",7:"Jul",8:"Ago",9:"Sep",10:"Oct",11:"Nov",12:"Dic"}
    events = group_materialized_events(detail)
    links = problem_links.copy() if problem_links is not None else pd.DataFrame()
    risks = []
    risk_rows = detail[detail["classification_type"] == "RISK"]
    for risk_id, group in risk_rows.groupby("risk_id"):
        risk = RISK_BY_ID.get(risk_id, {"name":"Riesgo no catalogado", "owner":"", "impact":"", "responsible_r":"", "accountable_a":"", "consulted_c":"", "informed_i":""})
        tickets = sorted(group["numero"].unique())
        risk_events = events[events["risk_id"] == risk_id]
        row = {"ID": risk_id, "Riesgo Materializado": risk["name"], "Dueño del Riesgo": risk["owner"],
               "Impacto Escala": risk["impact"], "Cantidad tickets asociados": len(tickets),
               "Tickets Asociados": ", ".join(tickets), "Asignación Operativa RACI": f"R: {risk['responsible_r']} | A: {risk['accountable_a']} | C: {risk['consulted_c']} | I: {risk['informed_i']}",
               "Estado": recurrence_status(risk_events["event_id"].nunique()), "R": risk["responsible_r"], "A": risk["accountable_a"], "C": risk["consulted_c"], "I": risk["informed_i"]}
        risk_links = links[links["risk_id"] == risk_id] if not links.empty and "risk_id" in links else pd.DataFrame()
        row["Eventos reales"] = int(risk_events["event_id"].nunique())
        row["Problemas asociados"] = ", ".join(sorted(risk_links["problem_number"].astype(str).unique())) if not risk_links.empty else ""
        row["Estado del problema"] = ", ".join(sorted(risk_links["problem_status"].fillna("Sin estado").astype(str).unique())) if not risk_links.empty else "Sin problema asociado"
        row["Estado causa raíz"] = ", ".join(sorted(risk_events["root_cause_status"].unique())) if not risk_events.empty else "Pendiente de investigación"
        natures = [value for value in risk_events.get("incident_nature", pd.Series(dtype=str)).fillna("").astype(str).unique() if value]
        causes = [value for value in risk_events.get("root_cause", pd.Series(dtype=str)).fillna("").astype(str).unique() if value]
        row["Naturaleza consolidada de los INC"] = " | ".join(natures) if natures else "En proceso de caracterización"
        row["Causa raíz consolidada"] = " | ".join(causes) if causes else "En proceso de análisis"
        for month in months:
            row[month_labels[month]] = int(group.loc[group["mes_num"] == month, "numero"].nunique())
        risks.append(row)
    risk_columns = ["ID", "Riesgo Materializado", "Dueño del Riesgo", "Impacto Escala",
                    "Cantidad tickets asociados", "Tickets Asociados", "Asignación Operativa RACI", "Estado",
                    *[month_labels[m] for m in months], "Eventos reales", "Problemas asociados", "Estado del problema",
                    "Naturaleza consolidada de los INC", "Causa raíz consolidada", "Estado causa raíz", "R", "A", "C", "I"]
    risks_df = pd.DataFrame(risks, columns=risk_columns)
    if not risks_df.empty:
        risks_df = risks_df.sort_values("Cantidad tickets asociados", ascending=False)
    reconciliation_rows = [
        {"Categoría de los Tickets": f"{risk['risk_id']} – {risk['name']}",
         "Cantidad tickets asociados": int(risk_rows.loc[risk_rows["risk_id"] == risk["risk_id"], "numero"].nunique()),
         "Impacto en Matriz": "Sí"}
        for risk in RISK_CATALOG
        if risk["active"]
    ]
    reconciliation_rows.extend([
        {"Categoría de los Tickets": "Incidentes clasificados como exclusión", "Cantidad tickets asociados": validation["exclusion"], "Impacto en Matriz": "No"},
        {"Categoría de los Tickets": "Pendientes de clasificación", "Cantidad tickets asociados": validation["pending"], "Impacto en Matriz": "Pendiente de análisis"},
    ])
    reconciliation = pd.DataFrame(reconciliation_rows)
    exclusions = []
    for category, meta in EXCLUSION_CATEGORIES.items():
        group = detail[(detail["classification_type"] == "EXCLUSION") & (detail["exclusion_category"] == category)]
        examples = [f"{r.numero}: {r.breve_descripcion or r.descripcion}" for r in group.head(3).itertuples()]
        exclusions.append({"Categoría de Exclusión": category, "Cantidad tickets asociados": group["numero"].nunique(),
                           "Justificación Metodológica y Ejemplos": meta["description"] + ((" Ejemplos: " + " | ".join(examples)) if examples else "")})
    stats = {"rules": Counter(detail.loc[detail["classification_type"] != "PENDING", "classification_reason"].str.split(":").str[0]),
             "pending": validation["pending"], "manual_overrides": int((detail["classification_source"] == "MANUAL").sum()),
             "ambiguous": detail[detail["candidate_risks"].fillna("").str.contains(",")]}
    return {"detail": detail, "risks": risks_df, "events": events, "problem_links": links,
            "reconciliation": reconciliation,
            "exclusions": pd.DataFrame(exclusions), "pending": detail[detail["classification_type"] == "PENDING"].copy(),
            "validation": validation, "statistics": stats}
