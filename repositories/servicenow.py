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
    db_execute(conn, "ALTER TABLE servicenow_sync_state ENABLE ROW LEVEL SECURITY")


def guardar_documento(conn, documento, tabla_local, row, actor_email):
    if tabla_local not in {"cases", "incidents"}:
        raise ValueError("Tabla local no permitida.")
    columns = [validar_identificador_sql(column) for column in row]
    updates = ", ".join(f"{c} = EXCLUDED.{c}" for c in columns if c != "numero")
    db_execute(conn, f"INSERT INTO {tabla_local} ({', '.join(columns)}) "
               f"VALUES ({db_placeholders(len(columns))}) ON CONFLICT (numero) DO UPDATE SET {updates}",
               tuple(row[c] for c in columns))
    db_execute(conn, """INSERT INTO servicenow_records (tabla, sys_id, numero, documento, sincronizado_por)
        VALUES (?, ?, ?, ?::jsonb, ?)
        ON CONFLICT (tabla, sys_id) DO UPDATE SET
            numero = EXCLUDED.numero, documento = EXCLUDED.documento,
            sincronizado_por = EXCLUDED.sincronizado_por, sincronizado_en = CURRENT_TIMESTAMP""",
        (documento["tabla"], documento["case_id"], row["numero"],
         json.dumps(documento, ensure_ascii=False), actor_email))


def leer_documento(conn, tabla, numero):
    result = db_execute(conn, "SELECT documento FROM servicenow_records WHERE tabla = ? AND numero = ?",
                        (tabla, numero)).fetchone()
    return result[0] if result else None


def leer_estado_sincronizacion(conn):
    row = db_execute(conn, """SELECT ultima_correcta, ultimo_intento, estado, mensaje, registros
        FROM servicenow_sync_state WHERE tipo = 'casos'""").fetchone()
    return dict(zip(("ultima_correcta", "ultimo_intento", "estado", "mensaje", "registros"), row)) if row else None


def guardar_estado_sincronizacion(conn, *, estado, intento, correcta=None, mensaje="", registros=0):
    db_execute(conn, """INSERT INTO servicenow_sync_state
        (tipo, ultima_correcta, ultimo_intento, estado, mensaje, registros)
        VALUES ('casos', ?, ?, ?, ?, ?)
        ON CONFLICT (tipo) DO UPDATE SET
            ultima_correcta = COALESCE(EXCLUDED.ultima_correcta, servicenow_sync_state.ultima_correcta),
            ultimo_intento = EXCLUDED.ultimo_intento, estado = EXCLUDED.estado,
            mensaje = EXCLUDED.mensaje, registros = EXCLUDED.registros""",
        (correcta, intento, estado, mensaje, registros))
