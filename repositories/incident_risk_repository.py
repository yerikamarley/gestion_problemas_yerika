"""Consultas de una sola pasada para el análisis de riesgos de incidentes."""

import pandas as pd

from core.permissions import ACTION_WRITE_PROBLEMS, puede_ejecutar
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


def fetch_available_problems():
    exigir_contexto_consulta()
    conn = get_conn()
    try:
        cursor = db_execute(conn, """
            SELECT numero, declaracion_problema, estado, prioridad, asignado_a,
                   creado, comentarios, notas_trabajo
            FROM problems ORDER BY creado DESC, numero DESC
        """)
        return pd.DataFrame(cursor.fetchall(), columns=[c[0] for c in cursor.description])
    finally:
        conn.close()


def fetch_risk_problem_links():
    exigir_contexto_consulta()
    conn = get_conn()
    try:
        cursor = db_execute(conn, """
            SELECT l.risk_id, l.problem_number, l.link_status, l.notes,
                   l.created_by, l.created_at, l.updated_at,
                   p.declaracion_problema, p.estado AS problem_status,
                   p.prioridad AS problem_priority, p.asignado_a AS problem_owner
            FROM risk_problem_links l
            JOIN problems p ON p.numero = l.problem_number
            ORDER BY l.risk_id, l.updated_at DESC
        """)
        return pd.DataFrame(cursor.fetchall(), columns=[c[0] for c in cursor.description])
    except Exception as exc:
        if getattr(exc, "sqlstate", None) == "42P01":
            return pd.DataFrame()
        raise
    finally:
        conn.close()


def fetch_risk_problem_link_history():
    exigir_contexto_consulta()
    conn = get_conn()
    try:
        cursor = db_execute(conn, """SELECT risk_id, problem_number, action, notes,
                   changed_by, changed_at FROM risk_problem_link_history ORDER BY changed_at DESC""")
        return pd.DataFrame(cursor.fetchall(), columns=[c[0] for c in cursor.description])
    except Exception as exc:
        if getattr(exc, "sqlstate", None) == "42P01":
            return pd.DataFrame()
        raise
    finally:
        conn.close()


def _authorize_problem_link(conn, actor_email):
    row = db_execute(conn, "SELECT role, active FROM app_users WHERE email = ?", (str(actor_email or "").strip().lower(),)).fetchone()
    if not row or not bool(row[1]) or not puede_ejecutar(row[0], ACTION_WRITE_PROBLEMS):
        raise PermissionError("Tu rol permite consultar, pero no modificar asociaciones de problemas.")


def save_risk_problem_link(risk_id, problem_number, notes, actor_email):
    risk_id, problem_number = str(risk_id or "").strip().upper(), str(problem_number or "").strip().upper()
    if not risk_id or not problem_number:
        raise ValueError("Selecciona un riesgo y un problema.")
    conn = get_conn()
    try:
        _authorize_problem_link(conn, actor_email)
        if not db_execute(conn, "SELECT 1 FROM problems WHERE numero = ?", (problem_number,)).fetchone():
            raise ValueError("El problema seleccionado ya no existe.")
        db_execute(conn, """
            INSERT INTO risk_problem_links
                (risk_id, problem_number, link_status, notes, created_by, created_at, updated_at)
            VALUES (?, ?, 'CONFIRMED', ?, ?, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP)
            ON CONFLICT (risk_id, problem_number) DO UPDATE
            SET link_status = 'CONFIRMED', notes = EXCLUDED.notes,
                updated_at = CURRENT_TIMESTAMP
        """, (risk_id, problem_number, str(notes or "").strip(), str(actor_email).strip().lower()))
        db_execute(conn, """INSERT INTO risk_problem_link_history
            (risk_id, problem_number, action, notes, changed_by)
            VALUES (?, ?, 'LINKED', ?, ?)""", (risk_id, problem_number, str(notes or "").strip(), str(actor_email).strip().lower()))
        conn.commit()
    except Exception:
        conn.rollback(); raise
    finally:
        conn.close()


def remove_risk_problem_link(risk_id, problem_number, actor_email):
    conn = get_conn()
    try:
        _authorize_problem_link(conn, actor_email)
        db_execute(conn, "DELETE FROM risk_problem_links WHERE risk_id = ? AND problem_number = ?", (risk_id, problem_number))
        db_execute(conn, """INSERT INTO risk_problem_link_history
            (risk_id, problem_number, action, changed_by) VALUES (?, ?, 'UNLINKED', ?)""",
            (risk_id, problem_number, str(actor_email).strip().lower()))
        conn.commit()
    except Exception:
        conn.rollback(); raise
    finally:
        conn.close()
