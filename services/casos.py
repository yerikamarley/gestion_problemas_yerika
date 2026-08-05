"""Resúmenes analíticos reutilizables para los dashboards de casos."""

import calendar
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
    return "cliente" in estado and any(palabra in estado for palabra in ("esper", "pendiente", "respuesta"))


def resumen_diario_soporte(df, anio, mes, alcance="soporte_y_sin_asignacion"):
    """Resume el mes usando fechas reales y la asignación/estado actuales."""
    columnas = [
        "Fecha", "Nuevos", "Cerrados", "Abiertos al cierre",
        "Esperando cliente*", "Sin asignación*", "Balance diario",
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
    cerrados = _serie_fechas(trabajo, "cerrado")
    estados = trabajo.get("estado", pd.Series("", index=trabajo.index))
    esperando = estados.apply(_estado_esperando_cliente)
    sin_asignar = trabajo[COL_SEGMENTO_ASIGNACION].eq(SEGMENTO_SIN_ASIGNACION)
    filas = []
    for dia in range(1, calendar.monthrange(int(anio), int(mes))[1] + 1):
        fecha = pd.Timestamp(int(anio), int(mes), dia)
        fin_dia = fecha + pd.Timedelta(days=1)
        nuevos_dia = creados.ge(fecha) & creados.lt(fin_dia)
        cerrados_dia = cerrados.ge(fecha) & cerrados.lt(fin_dia)
        pendientes = creados.lt(fin_dia) & (cerrados.isna() | cerrados.ge(fin_dia))
        nuevos = int(nuevos_dia.sum())
        cerrados_total = int(cerrados_dia.sum())
        filas.append({
            "Fecha": fecha,
            "Nuevos": nuevos,
            "Cerrados": cerrados_total,
            "Abiertos al cierre": int(pendientes.sum()),
            "Esperando cliente*": int((pendientes & esperando).sum()),
            "Sin asignación*": int((pendientes & sin_asignar).sum()),
            "Balance diario": nuevos - cerrados_total,
        })
    return pd.DataFrame(filas, columns=columnas)
