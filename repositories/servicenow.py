"""Persistencia de snapshots JSON junto con la fila usada por los tableros."""

import json
from collections import defaultdict

from repositories.database import db_execute, db_placeholders, validar_identificador_sql

CONFLICTO_SNAPSHOT = """ON CONFLICT (tabla, sys_id) DO UPDATE SET
    numero = EXCLUDED.numero,
    documento = CASE
        WHEN jsonb_array_length(COALESCE(EXCLUDED.documento->'historial_conversacion',
                                       EXCLUDED.documento->'traza_unificada', '[]'::jsonb)) = 0
        THEN EXCLUDED.documento || jsonb_build_object(
            'historial_conversacion', COALESCE(servicenow_records.documento->'historial_conversacion',
                                               servicenow_records.documento->'traza_unificada', '[]'::jsonb))
        ELSE EXCLUDED.documento END,
    sincronizado_por = EXCLUDED.sincronizado_por, sincronizado_en = CURRENT_TIMESTAMP"""


def ensure_servicenow_schema(conn):
    db_execute(conn, """CREATE TABLE IF NOT EXISTS servicenow_records (
        tabla TEXT NOT NULL,
        sys_id TEXT NOT NULL,
        numero TEXT NOT NULL,
        documento JSONB NOT NULL,
        sincronizado_en TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP,
        sincronizado_por TEXT NOT NULL,
        PRIMARY KEY (tabla, sys_id),
        UNIQUE (tabla, numero)
    )""")
    # Supabase expone tablas públicas por REST: solo el backend PostgreSQL
    # (propietario/BYPASSRLS) puede consultar estos snapshots sin políticas.
    db_execute(conn, "ALTER TABLE servicenow_records ENABLE ROW LEVEL SECURITY")
    db_execute(conn, """CREATE TABLE IF NOT EXISTS servicenow_sync_state (
        tipo TEXT PRIMARY KEY,
        ultima_correcta TIMESTAMPTZ,
        ultimo_intento TIMESTAMPTZ,
        estado TEXT NOT NULL DEFAULT 'pendiente',
        mensaje TEXT NOT NULL DEFAULT '',
        registros INTEGER NOT NULL DEFAULT 0
    )""")
    db_execute(conn, """ALTER TABLE servicenow_sync_state
        ADD COLUMN IF NOT EXISTS habilitada BOOLEAN NOT NULL DEFAULT FALSE,
        ADD COLUMN IF NOT EXISTS actor_email TEXT NOT NULL DEFAULT '',
        ADD COLUMN IF NOT EXISTS ultima_datos TIMESTAMPTZ""")
    db_execute(conn, "ALTER TABLE servicenow_sync_state ENABLE ROW LEVEL SECURITY")


def guardar_documento(conn, documento, tabla_local, row, actor_email):
    if tabla_local not in {"cases", "incidents"}:
        raise ValueError("Tabla local no permitida.")
    columns = [validar_identificador_sql(column) for column in row]
    updates = ", ".join(f"{c} = EXCLUDED.{c}" for c in columns if c != "numero")
    db_execute(conn, f"INSERT INTO {tabla_local} ({', '.join(columns)}) "
               f"VALUES ({db_placeholders(len(columns))}) ON CONFLICT (numero) DO UPDATE SET {updates}",
               tuple(row[c] for c in columns))
    snapshot = dict(documento)
    if "historial_conversacion" in snapshot:
        snapshot.pop("traza_unificada", None)  # Un solo ejemplar de los mensajes en JSON.
    db_execute(conn, """INSERT INTO servicenow_records (tabla, sys_id, numero, documento, sincronizado_por)
        VALUES (?, ?, ?, ?::jsonb, ?) """ + CONFLICTO_SNAPSHOT,
        (documento["tabla"], documento["case_id"], row["numero"],
         json.dumps(snapshot, ensure_ascii=False), actor_email))


def guardar_documentos_pagina(conn, documentos, tabla_local, filas, actor_email):
    """Upserts por página, sin viajes por fila ni pipeline del pooler."""
    if tabla_local not in {"cases", "incidents"}:
        raise ValueError("Tabla local no permitida.")
    if len(documentos) != len(filas):
        raise ValueError("Cada documento debe tener una fila analítica.")
    grupos = defaultdict(list)
    snapshots = []
    for documento, row in zip(documentos, filas):
        columnas = tuple(sorted(validar_identificador_sql(c) for c in row))
        grupos[columnas].append(row)
        snapshot = dict(documento)
        if "historial_conversacion" in snapshot:
            snapshot.pop("traza_unificada", None)
        snapshots.append({"tabla": documento["tabla"], "sys_id": documento["case_id"],
                          "numero": row["numero"], "documento": snapshot, "sincronizado_por": actor_email})
    # Agrupar por las claves presentes mantiene intactos los campos históricos
    # omitidos por la API; JSON null y cadenas vacías siguen siendo actualizaciones.
    for columnas, grupo in grupos.items():
        campos = ", ".join(columnas)
        updates = ", ".join(f"{c} = EXCLUDED.{c}" for c in columnas if c != "numero")
        conflicto = f"DO UPDATE SET {updates}" if updates else "DO NOTHING"
        db_execute(conn, f"INSERT INTO {tabla_local} ({campos}) SELECT {campos} "
                   f"FROM jsonb_populate_recordset(NULL::{tabla_local}, ?::jsonb) WHERE TRUE "
                   f"ON CONFLICT (numero) {conflicto}",
                   (json.dumps(grupo, ensure_ascii=False),))
    if snapshots:
        db_execute(conn, """INSERT INTO servicenow_records (tabla, sys_id, numero, documento, sincronizado_por)
            SELECT tabla, sys_id, numero, documento, sincronizado_por
            FROM jsonb_to_recordset(?::jsonb) AS lote
                (tabla text, sys_id text, numero text, documento jsonb, sincronizado_por text) WHERE TRUE """ + CONFLICTO_SNAPSHOT,
            (json.dumps(snapshots, ensure_ascii=False),))


def leer_documento(conn, tabla, numero):
    result = db_execute(conn, "SELECT documento FROM servicenow_records WHERE tabla = ? AND numero = ?",
                        (tabla, numero)).fetchone()
    if not result:
        return None
    documento = dict(result[0])
    if "historial_conversacion" in documento:
        documento["traza_unificada"] = documento["historial_conversacion"]
    return documento


def leer_historial_conversacion(conn, tabla, numero):
    row = db_execute(conn, """SELECT COALESCE(documento->'historial_conversacion', documento->'traza_unificada', '[]'::jsonb)
        FROM servicenow_records WHERE tabla = ? AND numero = ?""", (tabla, numero)).fetchone()
    return row[0] if row else None


def leer_campos_casos(conn, numeros):
    """Una consulta por selección; no descarga nuevamente los cuerpos de correo."""
    if not numeros:
        return []
    return db_execute(conn, """SELECT numero, documento->'campos_raw', documento->'slas'
        FROM servicenow_records
        WHERE tabla = 'sn_customerservice_case' AND numero = ANY(?::text[])""",
        (list(dict.fromkeys(numeros)),)).fetchall()


def leer_evidencia_casos(conn, numeros):
    """Una lectura local por selección para conservar autor y dirección del correo."""
    if not numeros:
        return []
    return db_execute(conn, """SELECT numero, documento->'campos_raw'->'description',
        COALESCE(documento->'historial_conversacion', documento->'traza_unificada', '[]'::jsonb)
        FROM servicenow_records WHERE tabla = 'sn_customerservice_case' AND numero = ANY(?::text[])""",
        (list(dict.fromkeys(numeros)),)).fetchall()


def activar_sincronizacion_casos(conn, actor_email):
    # La activación sobrevive al reinicio del servidor y no guarda credenciales.
    db_execute(conn, """INSERT INTO servicenow_sync_state (tipo, habilitada, actor_email)
        VALUES ('casos', TRUE, ?) ON CONFLICT (tipo) DO UPDATE SET
        habilitada = TRUE, actor_email = EXCLUDED.actor_email""", (actor_email,))


def leer_activacion_casos(conn):
    row = db_execute(conn, "SELECT habilitada, actor_email FROM servicenow_sync_state WHERE tipo = 'casos'").fetchone()
    return (bool(row[0]), row[1]) if row else (False, "")


def leer_estado_sincronizacion(conn):
    row = db_execute(conn, """SELECT ultima_correcta, ultimo_intento, estado, mensaje, registros, ultima_datos
        FROM servicenow_sync_state WHERE tipo = 'casos'""").fetchone()
    return dict(zip(("ultima_correcta", "ultimo_intento", "estado", "mensaje", "registros", "ultima_datos"), row)) if row else None


def guardar_estado_sincronizacion(conn, *, estado, intento, correcta=None, mensaje="", registros=0):
    db_execute(conn, """INSERT INTO servicenow_sync_state
        (tipo, ultima_correcta, ultimo_intento, estado, mensaje, registros, ultima_datos)
        VALUES ('casos', ?, ?, ?, ?, ?, ?)
        ON CONFLICT (tipo) DO UPDATE SET
            ultima_correcta = COALESCE(EXCLUDED.ultima_correcta, servicenow_sync_state.ultima_correcta),
            ultimo_intento = EXCLUDED.ultimo_intento, estado = EXCLUDED.estado,
            mensaje = EXCLUDED.mensaje, registros = EXCLUDED.registros,
            ultima_datos = CASE WHEN EXCLUDED.ultima_correcta IS NOT NULL AND EXCLUDED.registros > 0
                THEN EXCLUDED.ultima_correcta ELSE servicenow_sync_state.ultima_datos END""",
        (correcta, intento, estado, mensaje, registros, correcta if registros > 0 else None))


def guardar_estado_sincronizacion_manual(conn, *, estado, intento, mensaje="", registros=0):
    """Actualiza el estado visible sin mover el cursor de la ingesta automática."""
    db_execute(conn, """INSERT INTO servicenow_sync_state
        (tipo, ultimo_intento, estado, mensaje, registros, ultima_datos)
        VALUES ('casos', ?, ?, ?, ?, CASE WHEN ? > 0 THEN CURRENT_TIMESTAMP ELSE NULL END)
        ON CONFLICT (tipo) DO UPDATE SET
            ultimo_intento = EXCLUDED.ultimo_intento, estado = EXCLUDED.estado,
            mensaje = EXCLUDED.mensaje, registros = EXCLUDED.registros,
            ultima_datos = COALESCE(EXCLUDED.ultima_datos, servicenow_sync_state.ultima_datos)""",
        (intento, estado, mensaje, registros, registros))
