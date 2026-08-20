"""Consultas de una sola pasada para el análisis de riesgos de incidentes."""

import pandas as pd

from repositories.database import db_execute, get_conn
from repositories.tables import exigir_contexto_consulta

INCIDENT_COLUMNS = (
    "numero", "creado", "cerrado", "breve_descripcion", "descripcion", "categoria",
    "servicio_negocio", "prioridad", "estado", "grupo_asignacion", "asignado_a",
    "tipo_falla", "empresa", "solicitante", "creado_por", "observaciones_trabajo",
    "observaciones_adicionales", "actualizaciones", "lista_notas_trabajo", "impacto",
    "tipificacion_original", "causa_raiz_original", "origen_auto", "tipificacion_auto",
    "tipo_incidente_auto", "causa_raiz_auto", "es_alerta_auto",
)


def available_incident_years():
    exigir_contexto_consulta()
    conn = get_conn()
    try:
        rows = db_execute(conn, """
            SELECT DISTINCT SUBSTRING(creado, 1, 4)::int AS year
            FROM incidents WHERE creado ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}' ORDER BY year DESC
        """).fetchall()
        return [row[0] for row in rows if row[0] is not None]
    finally:
        conn.close()


def fetch_incidents_period(year, month_from=1, month_to=12):
    exigir_contexto_consulta()
    year, month_from, month_to = int(year), int(month_from), int(month_to)
    start = pd.Timestamp(year=year, month=month_from, day=1)
    end = pd.Timestamp(year=year + (month_to == 12), month=1 if month_to == 12 else month_to + 1, day=1)
    conn = get_conn()
    try:
        sql = f"""SELECT {', '.join(INCIDENT_COLUMNS)} FROM incidents
                  WHERE creado >= ? AND creado < ? ORDER BY creado, numero"""
        cursor = db_execute(conn, sql, (start.isoformat(), end.isoformat()))
        return pd.DataFrame(cursor.fetchall(), columns=[c[0] for c in cursor.description])
    finally:
        conn.close()


def fetch_classification_overrides(incident_numbers):
    exigir_contexto_consulta()
    numbers = sorted({str(number).strip().upper() for number in incident_numbers if str(number).strip()})
    if not numbers:
        return pd.DataFrame()
    conn = get_conn()
    try:
        cursor = db_execute(conn, """
            SELECT incident_number, final_classification_type, final_risk_id,
                   final_exclusion_category, modified_by, modified_at, change_reason
            FROM incident_risk_classification_overrides
            WHERE incident_number = ANY(?)
        """, (numbers,))
        return pd.DataFrame(cursor.fetchall(), columns=[c[0] for c in cursor.description])
    except Exception as exc:
        if getattr(exc, "sqlstate", None) == "42P01":
            return pd.DataFrame()
        raise
    finally:
        conn.close()
