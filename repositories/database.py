"""Primitivas de conexión y ejecución para PostgreSQL/Supabase."""

import re
import time

from core.settings import SUPABASE_DATABASE_URL


DB_TRANSIENT_SQLSTATES = {"40P01", "40001"}
DB_MAX_REINTENTOS = 3
DB_REINTENTO_ESPERA_SEGUNDOS = 0.35
DB_BLOQUEO_CARGA_CASOS = "gestion_problemas_yerika:carga_casos"
DB_BATCH_SIZE = 1000
DB_LOOKUP_BATCH_SIZE = 5000


def db_placeholder():
    return "%s"


def db_placeholders(count):
    return ", ".join([db_placeholder()] * count)


def db_sql(sql):
    """Convierte placeholders neutrales a la sintaxis usada por psycopg."""
    return sql.replace("?", "%s")


def db_bool(value):
    return bool(value)


def get_conn():
    if not SUPABASE_DATABASE_URL:
        raise RuntimeError("Configura SUPABASE_DATABASE_URL en Secrets para conectar la base de datos.")

    import psycopg

    return psycopg.connect(SUPABASE_DATABASE_URL, connect_timeout=30)


def db_execute(conn, sql, params=()):
    return conn.execute(db_sql(sql), params)


def db_executemany(conn, sql, rows):
    if not rows:
        return
    with conn.cursor() as cursor:
        cursor.executemany(db_sql(sql), rows)


def emitir_progreso(progress_callback, valor, mensaje):
    if progress_callback:
        progress_callback(max(0, min(1, float(valor))), mensaje)


def lotes(valores, tamano):
    for inicio in range(0, len(valores), tamano):
        yield valores[inicio : inicio + tamano]


def es_error_db_transitorio(error):
    sqlstate = getattr(error, "sqlstate", None) or getattr(error, "pgcode", None)
    return sqlstate in DB_TRANSIENT_SQLSTATES


def ejecutar_con_reintentos_db(operacion, intentos=DB_MAX_REINTENTOS):
    for intento in range(intentos):
        try:
            return operacion()
        except Exception as error:
            if not es_error_db_transitorio(error) or intento >= intentos - 1:
                raise
            time.sleep(DB_REINTENTO_ESPERA_SEGUNDOS * (2**intento))


def validar_identificador_sql(nombre):
    if not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", str(nombre or "")):
        raise ValueError(f"Identificador SQL no valido: {nombre}")
    return nombre

