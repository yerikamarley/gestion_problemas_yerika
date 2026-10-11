"""Candidatos al seguimiento: dos consultas locales, sin llamadas a ServiceNow."""

import pandas as pd

from repositories.database import db_execute, db_placeholders, get_conn
from repositories.tables import columnas_existentes, exigir_contexto_consulta
from services.proveedores import CAMPOS_TEXTO, catalogo_ans


def leer_candidatos_proveedores(tabla, columnas):
    exigir_contexto_consulta()
    if tabla not in {"cases", "incidents"}:
        raise ValueError("Tabla no permitida")
    conn = get_conn()
    try:
        existentes = columnas_existentes(conn, tabla)
        seleccion = [c for c in columnas if c in existentes]
        campos = [c for c in (*CAMPOS_TEXTO, "historial_conversacion") if c in existentes]
        if not seleccion or not campos:
            return pd.DataFrame(columns=columnas)
        # No se filtra por fecha de creación: un escalamiento puede ser posterior.
        # Solo se transportan candidatos y las columnas de análisis, sin snapshots.
        contenido = "concat_ws(' ', " + ", ".join(f"COALESCE({c}::text, '')" for c in campos) + ")"
        aliases = sorted({a.casefold() for p in catalogo_ans()["proveedores"].values() for a in p["aliases"]})
        condiciones = [f"{contenido} ILIKE ANY (ARRAY[{db_placeholders(len(aliases))}]::text[])"]
        params = ["%" + alias + "%" for alias in aliases]
        if "nombre_proveedor" in existentes:
            condiciones.append("NULLIF(TRIM(nombre_proveedor::text), '') IS NOT NULL")
        if "escalado_proveedor" in existentes:
            condiciones.append("LOWER(TRIM(escalado_proveedor::text)) IN ('1', 'true', 'verdadero', 'si', 'sí', 'yes')")
        cursor = db_execute(conn, f"SELECT {', '.join(seleccion)} FROM {tabla} WHERE (" + " OR ".join(condiciones) + ")", params)
        df = pd.DataFrame(cursor.fetchall(), columns=seleccion)
        return df.reindex(columns=columnas)
    finally:
        conn.close()
