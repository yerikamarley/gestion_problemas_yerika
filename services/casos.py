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
