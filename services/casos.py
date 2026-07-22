"""Resúmenes analíticos reutilizables para los dashboards de casos."""

import pandas as pd


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

