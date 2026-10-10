"""Persistencia de snapshots JSON junto con la fila usada por los tableros."""

import json

from repositories.database import db_execute, db_placeholders, validar_identificador_sql


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
        VALUES (?, ?, ?, ?::jsonb, ?)
        ON CONFLICT (tabla, sys_id) DO UPDATE SET
            numero = EXCLUDED.numero,
            documento = CASE
                WHEN jsonb_array_length(COALESCE(EXCLUDED.documento->'historial_conversacion',
                                               EXCLUDED.documento->'traza_unificada', '[]'::jsonb)) = 0
                THEN EXCLUDED.documento || jsonb_build_object(
                    'historial_conversacion', COALESCE(servicenow_records.documento->'historial_conversacion',
                                                       servicenow_records.documento->'traza_unificada', '[]'::jsonb))
                ELSE EXCLUDED.documento END,
            sincronizado_por = EXCLUDED.sincronizado_por, sincronizado_en = CURRENT_TIMESTAMP""",
        (documento["tabla"], documento["case_id"], row["numero"],
         json.dumps(snapshot, ensure_ascii=False), actor_email))


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
