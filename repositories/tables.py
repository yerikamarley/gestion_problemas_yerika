"""Consultas genéricas y seguras sobre tablas de la aplicación."""

import re

import pandas as pd

from repositories.database import db_execute, get_conn, validar_identificador_sql


def read_table(table_name):
    table_name = validar_identificador_sql(table_name)
    conn = get_conn()
    cursor = db_execute(conn, f"SELECT * FROM {table_name}")
    rows = cursor.fetchall()
    columns = [column[0] for column in cursor.description]
    conn.close()
    return pd.DataFrame(rows, columns=columns)


def columnas_existentes(conn, table_name):
    table_name = validar_identificador_sql(table_name)
    return {
        row[0]
        for row in db_execute(
            conn,
            """
            SELECT column_name
            FROM information_schema.columns
            WHERE table_schema = 'public' AND table_name = ?
            """,
            (table_name,),
        ).fetchall()
    }


def columnas_select_seguras(conn, table_name, columns):
    if not columns:
        return "*", None, []
    existentes = columnas_existentes(conn, table_name)
    selected_columns = [validar_identificador_sql(column) for column in columns if column in existentes]
    missing_columns = [column for column in columns if column not in existentes]
    if not selected_columns:
        return "", columns, missing_columns
    return ", ".join(selected_columns), columns, missing_columns


def limites_periodo(anio=None, mes=None):
    if mes not in (None, "", "Todos"):
        if isinstance(mes, str) and "-" in mes:
            partes = mes.split("-")
            anio = int(partes[0])
            mes = int(partes[1])
        else:
            if anio in (None, "", "Todos"):
                return None
            anio = int(anio)
            mes = int(mes)
        siguiente_anio = anio + 1 if mes == 12 else anio
        siguiente_mes = 1 if mes == 12 else mes + 1
        return f"{anio:04d}-{mes:02d}", f"{siguiente_anio:04d}-{siguiente_mes:02d}"
    if anio not in (None, "", "Todos"):
        anio = int(anio)
        return f"{anio:04d}", f"{anio + 1:04d}"
    return None


def agregar_condicion_periodo(where, params, anio=None, mes=None):
    limites = limites_periodo(anio, mes)
    if limites:
        inicio, fin = limites
        where.append("(creado >= ? AND creado < ?)")
        params.extend([inicio, fin])


def valor_filtro_activo(valor):
    return valor not in (None, "", "Todos")


def read_table_filtered(table_name, columns=None, anio=None, mes=None, equals=None, likes=None, limit=None):
    table_name = validar_identificador_sql(table_name)
    conn = get_conn()
    column_sql, requested_columns, missing_columns = columnas_select_seguras(conn, table_name, columns)
    if not column_sql:
        conn.close()
        df = pd.DataFrame(columns=columns or [])
        df.attrs["missing_columns"] = missing_columns
        return df

    where = []
    params = []
    agregar_condicion_periodo(where, params, anio, mes)

    existentes = columnas_existentes(conn, table_name)
    for columna, valor in (equals or {}).items():
        if not valor_filtro_activo(valor) or columna not in existentes:
            continue
        columna = validar_identificador_sql(columna)
        where.append(f"{columna} = ?")
        params.append(valor)

    for columna, valor in (likes or {}).items():
        if not valor_filtro_activo(valor) or columna not in existentes:
            continue
        columna = validar_identificador_sql(columna)
        where.append(f"LOWER(COALESCE({columna}, '')) LIKE ?")
        params.append(f"%{str(valor).lower()}%")

    where_sql = f" WHERE {' AND '.join(where)}" if where else ""
    limit_sql = " LIMIT ?" if limit else ""
    if limit:
        params.append(int(limit))

    cursor = db_execute(
        conn,
        f"SELECT {column_sql} FROM {table_name}{where_sql} ORDER BY creado DESC{limit_sql}",
        params,
    )
    rows = cursor.fetchall()
    result_columns = [column[0] for column in cursor.description]
    conn.close()
    df = pd.DataFrame(rows, columns=result_columns)
    if requested_columns:
        for column in missing_columns:
            df[column] = pd.NA
        df = df[requested_columns]
    df.attrs["missing_columns"] = missing_columns
    return df


def obtener_meses_disponibles(table_name):
    table_name = validar_identificador_sql(table_name)
    conn = get_conn()
    cursor = db_execute(
        conn,
        f"""
        SELECT DISTINCT substr(creado, 1, 7) AS mes
        FROM {table_name}
        WHERE creado IS NOT NULL AND creado <> '' AND length(creado) >= 7
        ORDER BY mes
        """,
    )
    meses = [row[0] for row in cursor.fetchall() if row[0] and re.fullmatch(r"\d{4}-\d{2}", row[0])]
    conn.close()
    return meses


def obtener_ultimo_mes_disponible(table_name):
    meses = obtener_meses_disponibles(table_name)
    return meses[-1] if meses else ""


def read_table_years(table_name, years, columns=None):
    table_name = validar_identificador_sql(table_name)
    years = [str(year).strip() for year in years if str(year).strip()]
    if not years:
        return pd.DataFrame()
    conn = get_conn()
    column_sql, requested_columns, missing_columns = columnas_select_seguras(conn, table_name, columns)
    if not column_sql or (columns and "creado" not in column_sql.split(", ")):
        conn.close()
        df = pd.DataFrame(columns=columns or [])
        df.attrs["missing_columns"] = missing_columns
        return df
    condiciones = []
    params = []
    for year in years:
        siguiente = str(int(year) + 1)
        condiciones.append("(creado >= ? AND creado < ?)")
        params.extend([year, siguiente])
    cursor = db_execute(
        conn,
        f"SELECT {column_sql} FROM {table_name} WHERE {' OR '.join(condiciones)}",
        params,
    )
    rows = cursor.fetchall()
    result_columns = [column[0] for column in cursor.description]
    conn.close()
    df = pd.DataFrame(rows, columns=result_columns)
    if requested_columns:
        for column in missing_columns:
            df[column] = pd.NA
        df = df[requested_columns]
    df.attrs["missing_columns"] = missing_columns
    return df

