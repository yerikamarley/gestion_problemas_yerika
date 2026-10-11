"""Columnas completas de casos desde los snapshots locales, sin consultar la API."""

import json
from datetime import datetime, timezone
from zoneinfo import ZoneInfo

import pandas as pd

from services.casos_sla import COLUMNAS_CASOS_PLATAFORMA
from services.servicenow import CASE_FIELDS, valor


ETIQUETAS_CAMPOS = {
    "upon_reject": "Después del rechazo", "knowledge": "Conocimiento",
    "assigned_on": "Asignado el", "impact": "Impacto", "active": "Activo",
    "sys_domain_path": "Ruta de dominio", "sync_driver": "Controlador de sincronización",
    "notify": "Notificar", "sys_class_name": "Tipo de tarea", "closed_by": "Cerrado por",
    "reassignment_count": "Número de reasignación",
    "notes_to_comments": "Agregar notas de resolución a comentarios",
    "escalation": "Escalado", "upon_approval": "Después de la aprobación",
    "u_activaci_n_firma": "Activación firma", "task_effective_number": "Número efectivo",
    "sys_updated_by": "Actualizado por", "opened_by": "Abierto por",
    "sys_domain": "Dominio", "follow_the_sun": '"Follow the sun"',
    "opened_at": "Abierto", "case_report": "Informe de caso", "subcategory": "Subcategoría",
    "assignment_group": "Grupo de asignación", "proactive": "Proactivo",
    "auto_close": "Cierre automático", "urgency": "Urgencia",
    "u_approval_number_1": "Número de aprobaciones", "activity_due": "Vencimiento de la actividad",
    "u_boolean_3": "Escalado a proveedor (u_boolean_3)", "approval": "Aprobación",
    "sys_mod_count": "Actualizaciones", "category": "Categoría", "sys_id": "ID del caso",
    "description": "Descripción", "comments": "Comentarios",
}
for _local, _campos in CASE_FIELDS.items():
    if _local in COLUMNAS_CASOS_PLATAFORMA:
        for _campo in _campos:
            ETIQUETAS_CAMPOS.setdefault(_campo, COLUMNAS_CASOS_PLATAFORMA[_local])

FECHAS_API = {"sys_created_on", "sys_updated_on", "closed_at", "sla_due", "assigned_on",
              "opened_at", "activity_due", "due_date", "start_time", "planned_end_time"}
PREFIJO_CAMPO = "sn__"
PREFIJO_SLA = "sn_sla__"
ETIQUETAS_SLA = {"sla": "Nombre", "stage": "Estado", "start_time": "Inicio",
                 "planned_end_time": "Vencimiento", "has_breached": "Incumplido", "sys_id": "ID"}
CAMPO_PORTAL = {"active", "opened_at", "opened_by", "assigned_on", "case_report", "description",
                "u_boolean_3", "sys_mod_count", "sys_updated_by"}
COLUMNAS_CASOS_PORTAL = (
    "Número", "Estado", "Breve descripción", "Contacto", "Creado por", "Cuenta", "Creado",
    "Fecha de vencimiento del SLA", "Cerrado", "Asignado a", "Prioridad", "Producto",
    "Notas de resolución", "Observaciones adicionales", "Observaciones y notas de trabajo",
    "Notas del trabajo", "Causa Raiz*", "Escalado a proveedor", "Nombre del proveedor", "Actualizado",
    "Estado SLA", "Horas restantes / margen al cierre", "mes", "Grupo cliente clave",
    "Cliente clave detectado", "Fuente identificación VIP", "Segmento de asignación", "Tipologia caso",
    "tipificacion", "tiempo_respuesta", "Activo", "Abierto", "Abierto por",
    "Asignado el", "Informe de caso", "Descripción", "Escalado a proveedor (u_boolean_3)",
    "Actualizaciones", "Actualizado por", "SLA 1 · Estado", "SLA 1 · Inicio", "SLA 1 · Vencimiento",
    "Historial de la conversación",
)


def valor_columna(campo, nombre):
    contenido = valor(campo, legible=nombre not in FECHAS_API)
    if contenido is None:
        return ""
    if nombre in FECHAS_API and contenido:
        try:
            fecha = datetime.fromisoformat(str(contenido).replace("Z", "+00:00"))
            if fecha.tzinfo is None:
                fecha = fecha.replace(tzinfo=timezone.utc)
            return fecha.astimezone(ZoneInfo("America/Bogota")).strftime("%Y-%m-%d %H:%M:%S")
        except (ValueError, TypeError):
            pass  # UNKNOWN y otros valores especiales se conservan.
    if isinstance(contenido, bool) or str(contenido).lower() in {"true", "false"}:
        return "verdadero" if str(contenido).lower() == "true" else "falso"
    if isinstance(contenido, (dict, list)):
        return json.dumps(contenido, ensure_ascii=False)
    return str(contenido)


def ampliar_columnas_casos(df, registros, *, solo_portal=False):
    """Conserva filas e índices; expande campos y todos los SLA sin duplicar casos."""
    if df.empty or not registros:
        return df.copy()
    por_numero = {numero: (raw or {}, slas or []) for numero, raw, slas in registros}
    filas = []
    sustituciones = {col: [] for col in COLUMNAS_CASOS_PLATAFORMA}
    for _, row in df.iterrows():
        raw, slas = por_numero.get(row.get("numero"), ({}, []))
        usados = set()
        for col in sustituciones:
            presentes = [campo for campo in CASE_FIELDS.get(col, ()) if campo in raw]
            campo = next((c for c in presentes if valor(raw[c]) not in (None, "")),
                         presentes[0] if presentes else None)
            contenido = valor_columna(raw[campo], campo) if campo else ""
            sustituciones[col].append(contenido if contenido != "" else row.get(col, ""))
            if campo and ETIQUETAS_CAMPOS.get(campo) == COLUMNAS_CASOS_PLATAFORMA[col]:
                usados.add(campo)
        fila = {PREFIJO_CAMPO + campo: valor_columna(contenido, campo)
                for campo, contenido in raw.items()
                if campo not in usados and (not solo_portal or campo in CAMPO_PORTAL)}
        if (raw or slas) and not solo_portal:
            fila[PREFIJO_SLA + "cantidad"] = str(len(slas))
        for posicion, sla in enumerate(slas[:1] if solo_portal else slas, 1):
            for campo in ("stage", "start_time", "planned_end_time") if solo_portal else ETIQUETAS_SLA:
                fila[f"{PREFIJO_SLA}{posicion}__{campo}"] = valor_columna(sla.get(campo), campo)
        filas.append(fila)
    trabajo = df.copy()
    for col, contenidos in sustituciones.items():
        trabajo[col] = contenidos
    extras = pd.DataFrame(filas, index=df.index).fillna("")
    return pd.concat([trabajo, extras], axis=1)


def etiquetas_columnas_adicionales(columnas):
    resultado = {}
    existentes = set(COLUMNAS_CASOS_PLATAFORMA.values())
    for columna in columnas:
        if columna.startswith(PREFIJO_CAMPO):
            tecnico = columna[len(PREFIJO_CAMPO):]
            etiqueta = ETIQUETAS_CAMPOS.get(tecnico, tecnico)
            if etiqueta in existentes:
                etiqueta = f"{etiqueta} ({tecnico})"
        elif columna.startswith(PREFIJO_SLA):
            partes = columna[len(PREFIJO_SLA):].split("__", 1)
            etiqueta = "Cantidad de SLAs" if partes == ["cantidad"] else (
                f"SLA {partes[0]} · {ETIQUETAS_SLA[partes[1]]}")
        else:
            continue
        resultado[columna] = etiqueta
        existentes.add(etiqueta)
    return resultado


def tabla_casos_portal(df):
    from services.casos_sla import COL_ESTADO_SLA, tabla_casos_plataforma

    tabla = tabla_casos_plataforma(df, adicionales=[COL_ESTADO_SLA, "horas_para_vencer", *COLUMNAS_CASOS_PORTAL])
    return tabla.reindex(columns=COLUMNAS_CASOS_PORTAL)
