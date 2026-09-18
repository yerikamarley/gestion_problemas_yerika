"""SLA de casos según el vencimiento exportado por la plataforma de origen."""

import re

import pandas as pd


ZONA_PLATAFORMA = "America/Bogota"
COL_ESTADO_SLA = "estado_sla_caso"

# Orden de la exportación sn_customerservice_case compartida por el usuario.
COLUMNAS_CASOS_PLATAFORMA = {
    "numero": "Número",
    "estado": "Estado",
    "descripcion": "Breve descripción",
    "contacto": "Contacto",
    "creado_por": "Creado por",
    "cuenta": "Cuenta",
    "creado": "Creado",
    "sla_creado": "SLA creado",
    "fecha_vencimiento_sla": "Fecha de vencimiento del SLA",
    "cerrado": "Cerrado",
    "asignado": "Asignado a",
    "canal": "Canal",
    "prioridad": "Prioridad",
    "producto": "Producto",
    "notas_resolucion": "Notas de resolución",
    "observaciones_adicionales": "Observaciones adicionales",
    "observaciones_trabajo": "Observaciones y notas de trabajo",
    "notas_trabajo": "Notas del trabajo",
    "causa": "Causa Raiz*",
    "escalado_proveedor": "Escalado a proveedor",
    "nombre_proveedor": "Nombre del proveedor",
    "numero_caso_externo": "Numero de caso externo",
    "actualizado": "Actualizado",
    "producto_adicional": "Producto.1",
    "caso_referencia": "Caso",
    "codigo_resolucion": "Código de resolución",
}


def fecha_plataforma(valor):
    """Interpreta fechas sin zona como hora de Colombia, sin depender del servidor."""
    if valor is None or pd.isna(valor) or str(valor).strip() == "":
        return pd.NaT
    try:
        texto = str(valor).strip()
        fecha = pd.to_datetime(valor, errors="coerce", dayfirst=bool(re.match(r"^\d{1,2}[/\-]", texto)))
        if pd.isna(fecha):
            return pd.NaT
        return fecha.tz_localize(ZONA_PLATAFORMA) if fecha.tzinfo is None else fecha.tz_convert(ZONA_PLATAFORMA)
    except (ValueError, TypeError, OverflowError):
        return pd.NaT


def evaluar_sla_caso(row, ahora):
    vencimiento = fecha_plataforma(row.get("fecha_vencimiento_sla"))
    creado = fecha_plataforma(row.get("creado"))
    cierre_original = row.get("cerrado")
    cerrado = fecha_plataforma(cierre_original)
    estado = str(row.get("estado", "")).casefold()
    es_cerrado = pd.notna(cerrado) or bool(re.search(
        r"\b(?:cerrado|closed|resuelto|resolved|solucionado|finalizado|completado)\b", estado
    ))
    resultado = {COL_ESTADO_SLA: "Sin fecha de vencimiento", "horas_para_vencer": None}
    if pd.isna(vencimiento):
        original = row.get("fecha_vencimiento_sla")
        if original is not None and pd.notna(original) and str(original).strip():
            resultado[COL_ESTADO_SLA] = "Fechas inconsistentes"
        return resultado
    cierre_invalido = cierre_original is not None and pd.notna(cierre_original) and str(cierre_original).strip() and pd.isna(cerrado)
    creado_original = row.get("creado")
    creado_invalido = creado_original is not None and pd.notna(creado_original) and str(creado_original).strip() and pd.isna(creado)
    if cierre_invalido or creado_invalido or (pd.notna(creado) and (
        vencimiento < creado or (pd.notna(cerrado) and cerrado < creado) or (not es_cerrado and creado > ahora)
    )):
        resultado[COL_ESTADO_SLA] = "Fechas inconsistentes"
        return resultado
    if es_cerrado and pd.isna(cerrado):
        resultado[COL_ESTADO_SLA] = "Sin fecha de cierre"
        return resultado
    referencia = cerrado if es_cerrado else ahora
    resultado["horas_para_vencer"] = (vencimiento - referencia).total_seconds() / 3600
    resultado[COL_ESTADO_SLA] = (
        ("Cumple" if cerrado <= vencimiento else "No cumple") if es_cerrado
        else ("Dentro del plazo" if ahora <= vencimiento else "Vencido")
    )
    return resultado


def guardar_fecha_plataforma(valor):
    fecha = fecha_plataforma(valor)
    if pd.notna(fecha):
        return fecha.strftime("%Y-%m-%d %H:%M:%S")
    return "" if valor is None or pd.isna(valor) else str(valor).strip()


def agregar_sla_casos(df, ahora=None):
    trabajo = df.copy()
    instante = fecha_plataforma(ahora) if ahora is not None else pd.Timestamp.now(tz=ZONA_PLATAFORMA)
    resultados = [evaluar_sla_caso(row, instante) for _, row in trabajo.iterrows()]
    trabajo[COL_ESTADO_SLA] = pd.Series([r[COL_ESTADO_SLA] for r in resultados], index=trabajo.index, dtype="object")
    trabajo["horas_para_vencer"] = pd.Series([r["horas_para_vencer"] for r in resultados], index=trabajo.index, dtype="float64")
    return trabajo


def resumen_sla_casos(df):
    estados = agregar_sla_casos(df)[COL_ESTADO_SLA]
    cumple = int(estados.eq("Cumple").sum())
    no_cumple = int(estados.eq("No cumple").sum())
    evaluados = cumple + no_cumple
    return {
        "cumple": cumple,
        "no_cumple": no_cumple,
        "evaluados": evaluados,
        "porcentaje": round(cumple * 100 / evaluados, 2) if evaluados else None,
    }


def tabla_casos_plataforma(df, adicionales=()):
    """Mantiene las 26 columnas de origen y añade los resultados de la herramienta."""
    extras = [col for col in adicionales if col in df.columns and col not in COLUMNAS_CASOS_PLATAFORMA]
    return df.reindex(columns=[*COLUMNAS_CASOS_PLATAFORMA, *extras]).rename(
        columns={**COLUMNAS_CASOS_PLATAFORMA, COL_ESTADO_SLA: "Estado SLA", "horas_para_vencer": "Horas restantes / margen al cierre"}
    )
