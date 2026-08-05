"""Resúmenes analíticos reutilizables para los dashboards de casos."""

import unicodedata

import pandas as pd

from config.equipo_soporte import (
    EQUIPO_SOPORTE_CASOS,
    SEGMENTO_EQUIPO_SOPORTE,
    SEGMENTO_OTROS_RESPONSABLES,
    SEGMENTO_SIN_ASIGNACION,
)


COL_SEGMENTO_ASIGNACION = "Segmento de asignación"
VALORES_SIN_ASIGNACION = {"", "sin asignar", "no asignado", "unassigned", "none", "nan"}


def _normalizar_nombre(valor):
    if valor is None or pd.isna(valor):
        valor = ""
    texto = unicodedata.normalize("NFKD", str(valor))
    texto = "".join(caracter for caracter in texto if not unicodedata.combining(caracter))
    return " ".join(texto.casefold().strip().split())


EQUIPO_SOPORTE_NORMALIZADO = {_normalizar_nombre(nombre) for nombre in EQUIPO_SOPORTE_CASOS}


def segmento_asignacion(valor):
    """Clasifica un responsable sin mezclar otros equipos en las métricas."""
    normalizado = _normalizar_nombre(valor)
    if normalizado in VALORES_SIN_ASIGNACION:
        return SEGMENTO_SIN_ASIGNACION
    if normalizado in EQUIPO_SOPORTE_NORMALIZADO:
        return SEGMENTO_EQUIPO_SOPORTE
    return SEGMENTO_OTROS_RESPONSABLES


def agregar_segmento_asignacion(df, columna_asignado="asignado"):
    trabajo = df.copy()
    if columna_asignado not in trabajo.columns:
        trabajo[COL_SEGMENTO_ASIGNACION] = SEGMENTO_SIN_ASIGNACION
    else:
        trabajo[COL_SEGMENTO_ASIGNACION] = trabajo[columna_asignado].apply(segmento_asignacion)
    return trabajo


def segmentar_casos_por_asignacion(df, columna_asignado="asignado"):
    trabajo = agregar_segmento_asignacion(df, columna_asignado)
    return {
        "todos": trabajo,
        "equipo": trabajo[trabajo[COL_SEGMENTO_ASIGNACION] == SEGMENTO_EQUIPO_SOPORTE].copy(),
        "otros": trabajo[trabajo[COL_SEGMENTO_ASIGNACION] == SEGMENTO_OTROS_RESPONSABLES].copy(),
        "sin_asignacion": trabajo[trabajo[COL_SEGMENTO_ASIGNACION] == SEGMENTO_SIN_ASIGNACION].copy(),
    }


def top_categorias(df, columna, etiqueta, top_n=5, valor_vacio="Sin información"):
    """Agrupa una categoría y devuelve las de mayor volumen con su cantidad."""
    columnas = [etiqueta, "Cantidad"]
    if df.empty or columna not in df.columns:
        return pd.DataFrame(columns=columnas)

    serie = df[columna].replace("", pd.NA).fillna(valor_vacio).astype(str).str.strip()
    serie = serie.replace("", valor_vacio)
    return (
        serie.value_counts(dropna=False)
        .head(top_n)
        .rename_axis(etiqueta)
        .reset_index(name="Cantidad")
    )


def _serie_fechas(df, columna):
    if columna not in df.columns:
        return pd.Series(pd.NaT, index=df.index, dtype="datetime64[ns]")
    return pd.to_datetime(df[columna], errors="coerce")


def _estado_esperando_cliente(valor):
    estado = _normalizar_nombre(valor)
    return "esper" in estado or (
        "pendiente" in estado
        and any(palabra in estado for palabra in ("cliente", "usuario", "solicitante", "respuesta"))
    )


def _estado_cerrado(valor):
    estado = _normalizar_nombre(valor)
    return any(palabra in estado for palabra in ("cerrado", "closed", "resuelto", "resolved", "solucionado", "finalizado", "completado"))


def resumen_diario_soporte(df, anio, mes, alcance="soporte_y_sin_asignacion"):
    """Agrupa por día los casos creados en el mes y su estado actual."""
    columnas = [
        "Fecha", "Total del día", "Abiertos", "Esperando cliente",
        "Cerrados", "Sin asignación",
    ]
    if df.empty:
        return pd.DataFrame(columns=columnas)

    trabajo = agregar_segmento_asignacion(df)
    if alcance == "soporte":
        trabajo = trabajo[trabajo[COL_SEGMENTO_ASIGNACION] == SEGMENTO_EQUIPO_SOPORTE]
    elif alcance == "sin_asignacion":
        trabajo = trabajo[trabajo[COL_SEGMENTO_ASIGNACION] == SEGMENTO_SIN_ASIGNACION]
    else:
        trabajo = trabajo[
            trabajo[COL_SEGMENTO_ASIGNACION].isin([SEGMENTO_EQUIPO_SOPORTE, SEGMENTO_SIN_ASIGNACION])
        ]

    creados = _serie_fechas(trabajo, "creado")
    en_mes = creados.dt.year.eq(int(anio)) & creados.dt.month.eq(int(mes))
    trabajo = trabajo[en_mes].copy()
    creados = creados[en_mes]
    estados = trabajo.get("estado", pd.Series("", index=trabajo.index))
    esperando = estados.apply(_estado_esperando_cliente)
    cerrados = estados.apply(_estado_cerrado) | _serie_fechas(trabajo, "cerrado").notna()
    esperando = esperando & ~cerrados
    abiertos = ~cerrados & ~esperando
    sin_asignar = trabajo[COL_SEGMENTO_ASIGNACION].eq(SEGMENTO_SIN_ASIGNACION)
    filas = []
    for fecha in pd.date_range(f"{int(anio):04d}-{int(mes):02d}-01", periods=pd.Period(f"{anio}-{mes:02d}").days_in_month):
        del_dia = creados.dt.normalize().eq(fecha)
        filas.append({
            "Fecha": fecha,
            "Total del día": int(del_dia.sum()),
            "Abiertos": int((del_dia & abiertos).sum()),
            "Esperando cliente": int((del_dia & esperando).sum()),
            "Cerrados": int((del_dia & cerrados).sum()),
            "Sin asignación": int((del_dia & sin_asignar).sum()),
        })
    return pd.DataFrame(filas, columns=columnas)
