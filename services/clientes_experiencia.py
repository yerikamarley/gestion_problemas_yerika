"""Seguimiento de clientes clave: pendientes primero, sin scores inferidos."""

import re

import pandas as pd

from services.casos_historial import COL_CLASIFICACION_INICIAL, agregar_analisis_historial
from services.casos_sla import COL_ESTADO_SLA, ZONA_PLATAFORMA, agregar_sla_casos, fecha_plataforma


def texto(valor):
    return "" if valor is None or pd.isna(valor) else str(valor).strip()


def registros_unicos(df):
    trabajo = df.copy()
    if "numero" in trabajo:
        # Los registros sin número no deben desaparecer por compartir un vacío.
        identificados = trabajo["numero"].map(texto).ne("")
        repetidos = trabajo.loc[identificados, "numero"].duplicated(keep="last")
        trabajo = trabajo.drop(index=repetidos[repetidos].index)
    return trabajo.reset_index(drop=True)


def cerrado(row):
    return bool(texto(row.get("cerrado"))) or bool(re.search(
        r"\b(?:cerrado|closed|resuelto|resolved|solucionado|finalizado|completado)\b",
        texto(row.get("estado")).casefold(),
    ))


def ordenar_abiertos(df):
    """En el histórico conserva primero abiertos antiguos, después cerrados recientes."""
    trabajo = df.copy()
    if trabajo.empty:
        return trabajo
    trabajo["_cerrado"] = [cerrado(row) for _, row in trabajo.iterrows()]
    trabajo["_fecha"] = [fecha_plataforma(row.get("creado")) for _, row in trabajo.iterrows()]
    abiertos = trabajo[~trabajo["_cerrado"]].sort_values("_fecha", na_position="last", kind="stable")
    cerrados = trabajo[trabajo["_cerrado"]].sort_values("_fecha", ascending=False, na_position="last", kind="stable")
    return pd.concat([abiertos, cerrados]).drop(columns=["_cerrado", "_fecha"])


def seguimiento_clientes(casos, incidentes, ahora=None):
    """Recibe el período filtrado; cuenta una vez cada número de caso e incidente.

    El vencimiento de incidentes utiliza únicamente la fecha de origen, sin
    inventar una fecha a partir de un objetivo de ANS. La antigüedad usa Colombia.
    """
    instante = fecha_plataforma(ahora) if ahora is not None else pd.Timestamp.now(tz=ZONA_PLATAFORMA)
    casos = agregar_sla_casos(registros_unicos(casos), ahora=instante)
    incidentes = registros_unicos(incidentes)
    filas, historico = [], []
    for tipo, base in [("Caso", casos), ("Incidente", incidentes)]:
        for _, row in base.iterrows():
            cliente = texto(row.get("cliente_clave"))
            es_cerrado = cerrado(row)
            creado = fecha_plataforma(row.get("creado"))
            vencimiento = fecha_plataforma(row.get("fecha_vencimiento_sla"))
            if tipo == "Caso":
                estado_sla = row[COL_ESTADO_SLA]
            elif pd.isna(vencimiento):
                estado_sla = "Sin fecha de vencimiento"
            elif pd.notna(creado) and (vencimiento < creado or creado > instante):
                estado_sla = "Fechas inconsistentes"
            else:
                estado_sla = "Vencido" if vencimiento < instante else "Dentro del plazo"
            historico.append({"Cliente": cliente, "Tipo": tipo, "Abierto": not es_cerrado,
                              "Vencido": not es_cerrado and estado_sla == "Vencido"})
            if es_cerrado:
                continue
            prioridad = texto(row.get("prioridad"))
            numero_prioridad = re.match(r"^([1-5])\b", prioridad)
            orden_prioridad = int(numero_prioridad[1]) if numero_prioridad else {
                "crítica": 1, "critica": 1, "alta": 2, "moderada": 3, "baja": 4,
            }.get(prioridad.casefold(), 6)
            filas.append({
                "Cliente": cliente, "Tipo": tipo, "Número": texto(row.get("numero")),
                "Estado": texto(row.get("estado")), "Prioridad": prioridad,
                "Responsable": texto(row.get("asignado" if tipo == "Caso" else "asignado_a")) or "Sin asignar",
                "Solicitud": texto(row.get("descripcion" if tipo == "Caso" else "breve_descripcion")) or texto(row.get("descripcion")),
                "Creado": creado, "Vencimiento": vencimiento, "Estado SLA": estado_sla,
                "Días abiertos": round(max(0, (instante - creado).total_seconds() / 86400), 1) if pd.notna(creado) and creado <= instante else None,
                "_vencido": estado_sla == "Vencido", "_prioridad": orden_prioridad,
            })
    columnas = ["Cliente", "Tipo", "Número", "Estado", "Prioridad", "Estado SLA", "Días abiertos",
                "Responsable", "Solicitud", "Creado", "Vencimiento"]
    pendientes = pd.DataFrame(filas, columns=columnas + ["_vencido", "_prioridad"])
    pendientes = pendientes.sort_values(["_vencido", "_prioridad", "Creado", "Número"],
                                       ascending=[False, True, True, True], na_position="last", kind="stable")
    pendientes = pendientes[columnas].reset_index(drop=True)
    resumen = []
    for cliente in sorted({row["Cliente"] for row in historico}):
        registros = [row for row in historico if row["Cliente"] == cliente]
        abiertos = pendientes[pendientes["Cliente"].eq(cliente)]
        resumen.append({
            "Cliente": cliente,
            "Pendientes": sum(row["Abierto"] for row in registros),
            "Casos abiertos": sum(row["Abierto"] and row["Tipo"] == "Caso" for row in registros),
            "Casos vencidos": sum(row["Vencido"] and row["Tipo"] == "Caso" for row in registros),
            "Incidentes abiertos": sum(row["Abierto"] and row["Tipo"] == "Incidente" for row in registros),
            "Incidentes vencidos": sum(row["Vencido"] and row["Tipo"] == "Incidente" for row in registros),
            "Mayor antigüedad (días)": abiertos["Días abiertos"].max(),
            "Casos del período": sum(row["Tipo"] == "Caso" for row in registros),
            "Incidentes del período": sum(row["Tipo"] == "Incidente" for row in registros),
        })
    clientes = pd.DataFrame(resumen, columns=["Cliente", "Pendientes", "Casos abiertos", "Casos vencidos",
        "Incidentes abiertos", "Incidentes vencidos", "Mayor antigüedad (días)", "Casos del período", "Incidentes del período"])
    clientes["_vencidos"] = clientes["Casos vencidos"] + clientes["Incidentes vencidos"]
    clientes = clientes.sort_values(["_vencidos", "Pendientes", "Mayor antigüedad (días)", "Cliente"],
                                   ascending=[False, False, False, True], na_position="last").drop(columns="_vencidos")
    abiertos = casos.loc[[not cerrado(row) for _, row in casos.iterrows()]].copy() if not casos.empty else casos
    analisis = agregar_analisis_historial(abiertos)
    motivos = analisis.groupby(COL_CLASIFICACION_INICIAL).size().rename("Casos abiertos").reset_index()
    motivos = motivos.rename(columns={COL_CLASIFICACION_INICIAL: "Motivo de contacto"}).sort_values("Casos abiertos", ascending=False)
    return {"casos": ordenar_abiertos(casos), "incidentes": ordenar_abiertos(incidentes),
            "pendientes": pendientes, "clientes": clientes.reset_index(drop=True), "motivos": motivos}
