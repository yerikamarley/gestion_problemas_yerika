"""Comparación de cohortes mensuales de soporte con un día de corte común.

Estado, asignación y SLA corresponden a la última carga disponible, no a una
reconstrucción histórica de los estados al cierre de cada día.
"""

from calendar import monthrange
from datetime import date, timedelta
import re

import pandas as pd

from services.casos import segmentar_casos_por_asignacion
from services.casos_sla import agregar_sla_casos, fecha_plataforma, COL_ESTADO_SLA


def periodos_corte(corte):
    corte = pd.Timestamp(corte).date()
    inicio = corte.replace(day=1)
    ultimo_anterior = inicio - timedelta(days=1)
    fin_anterior = date(ultimo_anterior.year, ultimo_anterior.month,
                        min(corte.day, monthrange(ultimo_anterior.year, ultimo_anterior.month)[1]))
    return {"anterior": (ultimo_anterior.replace(day=1), fin_anterior), "actual": (inicio, corte)}


def _texto(valor):
    return "" if valor is None or pd.isna(valor) else str(valor).strip()


def preparar_base_ejecutiva(df, inicio, fin, clasificar_causa):
    trabajo = df.copy()
    for col in ["numero", "asignado", "creado", "cerrado", "estado", "actualizado"]:
        if col not in trabajo:
            trabajo[col] = ""
    trabajo["numero"] = trabajo["numero"].map(_texto)
    trabajo = trabajo[trabajo["numero"].ne("")]
    # Si llega un duplicado, la asignación de la última actualización prevalece.
    trabajo["_actualizado"] = trabajo["actualizado"].map(fecha_plataforma)
    trabajo = trabajo.sort_values("_actualizado", kind="stable", na_position="first").drop_duplicates("numero", keep="last")
    trabajo = segmentar_casos_por_asignacion(trabajo)["equipo"]
    fechas = trabajo["creado"].map(lambda valor: fecha_plataforma(valor))
    dias = fechas.map(lambda valor: valor.date() if pd.notna(valor) else None)
    trabajo = trabajo[dias.map(lambda dia: dia is not None and inicio <= dia <= fin).astype(bool)].copy()
    trabajo = agregar_sla_casos(trabajo)
    trabajo["cerrado_reporte"] = trabajo.apply(
        lambda row: bool(_texto(row["cerrado"])) or bool(re.search(
            r"\b(?:cerrado|closed|resuelto|resolved|solucionado|finalizado|completado)\b",
            _texto(row["estado"]).casefold())), axis=1,
    ) if not trabajo.empty else pd.Series(dtype=bool)
    causas = [_texto(clasificar_causa(row)) for _, row in trabajo.iterrows()]
    trabajo["causa_agrupada"] = [
        "Sin clasificar" if causa.casefold() in ("", "sin causa comun", "sin causa común", "sin dato") else causa
        for causa in causas
    ]
    return trabajo.drop(columns="_actualizado")


def metricas_ejecutivas(base):
    estados = base[COL_ESTADO_SLA]
    cumple = int(estados.eq("Cumple").sum())
    no_cumple = int(estados.eq("No cumple").sum())
    evaluados = cumple + no_cumple
    cerrados = int(base["cerrado_reporte"].sum())
    return {"total": len(base), "abiertos": len(base) - cerrados, "cerrados": cerrados,
            "cumple": cumple, "no_cumple": no_cumple, "evaluados": evaluados,
            "sin_evaluar": cerrados - evaluados,
            "sla": cumple * 100 / evaluados if evaluados else None}


def agrupar_causas(base):
    if base.empty:
        return pd.DataFrame(columns=["Causa", "Casos", "Porcentaje"])
    grupos = base.groupby("causa_agrupada").size().rename("Casos").reset_index().rename(columns={"causa_agrupada": "Causa"})
    grupos = grupos.sort_values(["Casos", "Causa"], ascending=[False, True]).reset_index(drop=True)
    grupos["Porcentaje"] = grupos["Casos"] * 100 / len(base)
    return grupos


def causas_para_lamina(causas):
    """Tres barras como en el estándar, sin omitir casos del denominador."""
    if len(causas) <= 3:
        return causas.copy()
    resto = causas.iloc[2:]
    return pd.concat([causas.head(2), pd.DataFrame([{
        "Causa": f"Otras causas ({len(resto)} grupos)",
        "Casos": int(resto["Casos"].sum()), "Porcentaje": float(resto["Porcentaje"].sum()),
    }])], ignore_index=True)


def construir_resumen_ejecutivo(anterior, actual, corte, clasificar_causa):
    periodos = periodos_corte(corte)
    bases = {nombre: preparar_base_ejecutiva(df, *periodos[nombre], clasificar_causa)
             for nombre, df in [("anterior", anterior), ("actual", actual)]}
    metricas = {nombre: metricas_ejecutivas(base) for nombre, base in bases.items()}
    a, b = metricas["anterior"], metricas["actual"]
    causas = agrupar_causas(bases["actual"])
    return {"periodos": periodos, "bases": bases, "metricas": metricas, "causas": causas,
            "causas_lamina": causas_para_lamina(causas),
            "diferencia": b["total"] - a["total"],
            "variacion": (b["total"] - a["total"]) * 100 / a["total"] if a["total"] else None,
            "diferencia_sla": b["sla"] - a["sla"] if a["sla"] is not None and b["sla"] is not None else None}
