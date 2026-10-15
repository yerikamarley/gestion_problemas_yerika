"""Categorías y relojes de soporte de la tabla 1 de SER-PR-001 v003.

Los límites se comparan en segundos, sin redondear. Las horas sin zona son
locales de Colombia. No se deduce una primera atención de notas o actualizaciones
ni una resolución del cierre administrativo.
"""

import json
import re
import unicodedata
from datetime import datetime, time, timedelta
from pathlib import Path

import pandas as pd

from services.casos_sla import fecha_plataforma


INCIDENTE = "Incidente"
SEGURIDAD = "Incidente Seguridad"
CONSULTA = "Consulta / Ayuda"
SIN_CATEGORIA = "Sin clasificar"
CATEGORIAS = (INCIDENTE, SEGURIDAD, CONSULTA)
PRIORIDADES = ("Critico", "Alto", "Moderado", "Bajo")
ATENCION_MINUTOS = dict(zip(PRIORIDADES, (15, 20, 30, 30)))
# Cantidad y unidad originales: un día se convierte con el horario de soporte.
RESOLUCION = {
    INCIDENTE: dict(zip(PRIORIDADES, ((4, "horas"), (8, "horas"), (4, "dias"), (8, "dias")))),
    SEGURIDAD: dict(zip(PRIORIDADES, ((4, "horas"), (8, "horas"), (24, "horas"), (48, "horas")))),
    CONSULTA: {"Moderado": (4, "dias")},
}
COORDINADORES = {
    INCIDENTE: "Director de Implementación y Servicio",
    SEGURIDAD: "Especialista en seguridad de la información y continuidad de negocio",
    CONSULTA: "Coordinación de redes y ciberseguridad",
}
COLUMNAS_FECHAS = ("fecha_notificacion", "fecha_atencion", "fecha_resolucion")


def texto(valor):
    if valor is None or pd.isna(valor):
        return ""
    return "".join(c for c in unicodedata.normalize("NFKD", str(valor).strip().lower())
                   if not unicodedata.combining(c))


def categoria_gestion(valor):
    categoria = re.sub(r"^\d+\s*[-.:]?\s*", "", texto(valor))
    if "consulta" in categoria or "ayuda" in categoria:
        return CONSULTA
    if "seguridad" in categoria:
        return SEGURIDAD
    if categoria in {"incidente", "incidentes", "incident", "incidente operativo"}:
        return INCIDENTE
    return SIN_CATEGORIA


def prioridad_gestion(valor):
    prioridad = texto(valor)
    prefijo = re.match(r"^(?:p\s*)?([1-4])(?:\s*[-.:]\s*.*)?$", prioridad)
    if prefijo:
        return PRIORIDADES[int(prefijo[1]) - 1]
    nombres = {"critico": "Critico", "critica": "Critico", "critical": "Critico",
               "alto": "Alto", "alta": "Alto", "high": "Alto",
               "moderado": "Moderado", "moderada": "Moderado", "media": "Moderado",
               "medio": "Moderado", "moderate": "Moderado", "bajo": "Bajo", "baja": "Bajo", "low": "Bajo"}
    return nombres.get(prioridad, "Sin prioridad")


def cargar_horario():
    ruta = Path(__file__).resolve().parent.parent / "config" / "incidentes_ans.json"
    return validar_horario(json.loads(ruta.read_text(encoding="utf-8")))


def validar_horario(horario):
    horario = dict(horario)
    dias = horario.get("dias_semana", [])
    franjas = horario.get("franjas", [])
    horas_dia = horario.get("horas_por_dia", 0)
    if (not dias or any(type(d) is not int or d not in range(7) for d in dias)
            or not franjas or not 0 < float(horas_dia) <= 24
            or horario.get("zona_horaria") != "America/Bogota"):
        raise ValueError("Horario de soporte de incidentes inválido")
    fin_previo = None
    for inicio, fin in franjas:
        inicio, fin = time.fromisoformat(inicio), time.fromisoformat(fin)
        if inicio >= fin or (fin_previo is not None and inicio < fin_previo):
            raise ValueError("Las franjas de soporte no pueden superponerse")
        fin_previo = fin
    for festivo in horario.get("festivos", []):
        datetime.strptime(festivo, "%Y-%m-%d")
    return horario


def ventanas_soporte(dia, horario):
    if dia.weekday() not in horario["dias_semana"] or dia.isoformat() in horario.get("festivos", []):
        return
    for inicio, fin in horario["franjas"]:
        yield (pd.Timestamp(datetime.combine(dia, time.fromisoformat(inicio)), tz=horario["zona_horaria"]),
               pd.Timestamp(datetime.combine(dia, time.fromisoformat(fin)), tz=horario["zona_horaria"]))


def segundos_soporte(inicio, fin, horario=None):
    horario = horario or cargar_horario()
    inicio, fin = fecha_plataforma(inicio), fecha_plataforma(fin)
    if pd.isna(inicio) or pd.isna(fin) or fin < inicio:
        return None
    segundos = 0.0
    dia = inicio.date()
    while dia <= fin.date():
        for desde, hasta in ventanas_soporte(dia, horario):
            if min(fin, hasta) > max(inicio, desde):
                segundos += (min(fin, hasta) - max(inicio, desde)).total_seconds()
        dia += timedelta(days=1)
    return segundos


def vencimiento_soporte(inicio, segundos, horario=None):
    horario = horario or cargar_horario()
    inicio = fecha_plataforma(inicio)
    if pd.isna(inicio) or segundos is None or segundos < 0:
        return pd.NaT
    restante = float(segundos)
    dia = inicio.date()
    # Los objetivos de la tabla caben holgadamente en este límite.
    for _ in range(3660):
        for desde, hasta in ventanas_soporte(dia, horario):
            desde = max(desde, inicio)
            if hasta < desde:
                continue
            disponible = (hasta - desde).total_seconds()
            if restante <= disponible:
                return desde + pd.Timedelta(seconds=restante)
            restante -= disponible
        dia += timedelta(days=1)
    raise ValueError("El horario no permite calcular el vencimiento")


def objetivos(categoria, prioridad, horario=None):
    horario = horario or cargar_horario()
    regla = RESOLUCION.get(categoria, {}).get(prioridad)
    if regla is None:
        return None, None
    cantidad, unidad = regla
    horas = cantidad * float(horario["horas_por_dia"]) if unidad == "dias" else cantidad
    return ATENCION_MINUTOS[prioridad], horas


def tabla_objetivos(horario=None):
    horario = horario or cargar_horario()
    filas = []
    for categoria in CATEGORIAS:
        for prioridad, (cantidad, unidad) in RESOLUCION[categoria].items():
            minutos, horas = objetivos(categoria, prioridad, horario)
            filas.append({"Categoría": categoria, "Cargo que coordina": COORDINADORES[categoria],
                          "Prioridad": prioridad.replace("Critico", "Crítico"),
                          "Atención máxima": f"{minutos} minutos",
                          "Resolución máxima": f"{cantidad} {'días' if unidad == 'dias' else 'horas'}",
                          "Horas de soporte": horas})
    return pd.DataFrame(filas)


def es_finalizado(row):
    estado = texto(row.get("estado"))
    return (bool(re.search(r"\b(cerrado|closed|resuelto|resolved|solucionado|finalizado|completado)\b", estado))
            or pd.notna(fecha_plataforma(row.get("cerrado")))
            or pd.notna(fecha_plataforma(row.get("fecha_resolucion"))))


def evaluar_incidente(row, ahora=None, horario=None):
    horario = horario or cargar_horario()
    ahora = fecha_plataforma(ahora) if ahora is not None else pd.Timestamp.now(tz=horario["zona_horaria"])
    categoria = categoria_gestion(row.get("categoria"))
    prioridad_original = prioridad_gestion(row.get("prioridad"))
    prioridad = "Moderado" if categoria == CONSULTA else prioridad_original
    atencion, resolucion = objetivos(categoria, prioridad, horario)
    # Una fecha explícita inválida no se sustituye silenciosamente por otra.
    notificacion = texto(row.get("fecha_notificacion"))
    inicio = fecha_plataforma(row.get("fecha_notificacion") if notificacion else row.get("creado"))
    fuente_inicio = "Notificación / apertura" if notificacion else "Creación del registro"
    finalizado = es_finalizado(row)
    validacion = ""
    if categoria == SIN_CATEGORIA:
        validacion = "Categoría pendiente de validar"
    elif categoria != CONSULTA and prioridad_original == "Sin prioridad":
        validacion = "Prioridad pendiente de validar"
    salida = {"categoria_gestion": categoria, "prioridad_gestion": prioridad,
              "coordinador_gestion": COORDINADORES.get(categoria, "Pendiente de clasificación"),
              "validacion_clasificacion": validacion, "inicio_sla": inicio,
              "fuente_inicio_sla": fuente_inicio, "finalizado_gestion": finalizado,
              "objetivo_atencion_minutos": atencion, "objetivo_resolucion_horas": resolucion}
    for etapa, campo, objetivo in (("atencion", "fecha_atencion", None if atencion is None else atencion * 60),
                                    ("resolucion", "fecha_resolucion", None if resolucion is None else resolucion * 3600)):
        original = row.get(campo)
        fin = fecha_plataforma(original)
        vencimiento = vencimiento_soporte(inicio, objetivo, horario)
        segundos = None
        motivo = ""
        if objetivo is None:
            estado = "Sin clasificación"
            motivo = "Falta una categoría o prioridad válida para definir el plazo."
        elif pd.isna(inicio) or inicio > ahora:
            estado = "Fechas inconsistentes" if texto(row.get("fecha_notificacion")) or texto(row.get("creado")) else "Sin fecha inicial"
            motivo = "La fecha inicial es inválida o futura." if estado == "Fechas inconsistentes" else "Falta la fecha de notificación o creación."
        elif texto(original) and (pd.isna(fin) or fin < inicio or fin > ahora):
            estado = "Fechas inconsistentes"
            motivo = f"La fecha de {'primera atención' if etapa == 'atencion' else 'resolución'} es inválida, anterior al inicio o futura."
        elif pd.notna(fin):
            segundos = segundos_soporte(inicio, fin, horario)
            estado = "Cumple" if segundos <= objetivo else "No cumple"
        elif etapa == "atencion":
            # Ausencia de evidencia de atención no demuestra que siga sin atenderse.
            estado = "Sin fecha de atención"
            motivo = "No hay fecha de primera atención disponible para medir este tiempo."
        elif finalizado:
            estado = "Sin fecha de resolución"
            motivo = "El caso está resuelto o cerrado, pero no hay fecha de resolución disponible."
        else:
            transcurrido = segundos_soporte(inicio, ahora, horario)
            estado = "Vencido" if transcurrido > objetivo else "Dentro del plazo"
        salida[f"{etapa}_estado"] = estado
        salida[f"{etapa}_segundos"] = segundos
        salida[f"{etapa}_vencimiento"] = vencimiento
        salida[f"{etapa}_motivo"] = motivo
    return salida


def agregar_evaluacion(df, ahora=None, horario=None):
    horario = horario or cargar_horario()
    ahora = ahora if ahora is not None else pd.Timestamp.now(tz=horario["zona_horaria"])
    resultado = df.copy()
    columnas = evaluar_incidente({}, ahora, horario)
    if df.empty:
        for columna in columnas:
            resultado[columna] = pd.Series(index=df.index, dtype="object")
        return resultado
    evaluados = pd.DataFrame([evaluar_incidente(row, ahora, horario) for row in df.to_dict("records")], index=df.index)
    for columna in evaluados:
        resultado[columna] = evaluados[columna]
    return resultado


def resumen_etapa(df, etapa):
    estados = df.get(f"{etapa}_estado", pd.Series(index=df.index, dtype="object"))
    cumple = int(estados.eq("Cumple").sum())
    no_cumple = int(estados.eq("No cumple").sum())
    evaluados = cumple + no_cumple
    pendientes = int(estados.isin(["Dentro del plazo", "Vencido"]).sum())
    return {"cumple": cumple, "no_cumple": no_cumple, "evaluados": evaluados,
            "porcentaje": round(100 * cumple / evaluados, 2) if evaluados else None,
            "pendientes": pendientes, "vencidos": int(estados.eq("Vencido").sum()),
            "sin_datos": len(df) - evaluados - pendientes}
