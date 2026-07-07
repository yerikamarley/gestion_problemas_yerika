import html
import json
import re
from io import BytesIO
from textwrap import dedent

import pandas as pd
import plotly.express as px
import streamlit as st
import streamlit.components.v1 as components
from PIL import Image, ImageDraw, ImageFont

from app_logic import (
    agregar_campos_sla_incidentes,
    agregar_campos_sla_respuesta,
    analizar_reincidencias_y_problemas,
    autenticar_usuario,
    calcular_disponibilidad_por_mes,
    contar_incidentes,
    eliminar_usuario,
    es_error_db_transitorio,
    guardar_casos,
    guardar_incidentes,
    guardar_usuario,
    init_db,
    limpiar_incidentes,
    listar_usuarios,
    load_casos,
    load_casos_anios,
    load_casos_filtrados,
    load_incidentes,
    load_incidentes_anios,
    load_incidentes_filtrados,
    obtener_meses_disponibles,
    obtener_ultimo_mes_disponible,
    resumir_disponibilidad_mes,
    duracion_sla_horas_incidente,
    estado_sla_incidente,
    familia_sla_incidente,
    normalizar_fecha,
    normalizar_prioridad_incidente,
    normalizar_texto,
    normalizar_email,
    sla_objetivo_horas_incidente,
)


# Constantes genericas para literales repetidos detectados por SonarQube
TEXT_CANTIDAD = 'Cantidad'
TEXT_TIPIFICACION_AUTO = 'tipificacion_auto'
TEXT_CREADO_DT = 'creado_dt'
TEXT_TIPIFICACION = 'Tipificacion'
TEXT_ESTADO_SLA = 'estado_sla'
TEXT_CLIENTE = 'Cliente'
TEXT_CASOS = 'Casos'
TEXT_TODOS = 'Todos'
TEXT_CREADO_DT_DASHBOARD = '_creado_dt_dashboard'
TEXT_COERCE = 'coerce'
TEXT_ABIERTOS = 'Abiertos'
TEXT_HORAS_PARA_VENCER = '_horas_para_vencer'
TEXT_CLIENTE_CLAVE = 'cliente_clave'
TEXT_ESTADO = 'estado'
TEXT_NUMERO = 'numero'
TEXT_CLIENTE_NORM_AGENDA = '_cliente_norm_agenda'
TEXT_CREADO = 'creado'
TEXT_ABIERTO = '_abierto'
TEXT_TIPIFICACION_2 = 'tipificacion'
TEXT_HORAS_ABIERTO = '_horas_abierto'
TEXT_PRIORIDAD = 'prioridad'
TEXT_DESCRIPCION = 'Descripcion'
TEXT_CUENTA = 'cuenta'
TEXT_YELLOW = 'yellow'
TEXT_FECHA = 'Fecha'
TEXT_INCIDENTES = 'Incidentes'
TEXT_OUTSIDE = 'outside'
TEXT_SLA_OBJETIVO_HORAS = 'sla_objetivo_horas'
TEXT_NIVEL = 'Nivel'
TEXT_DESCRIPCION_2 = 'descripcion'
TEXT_FUENTE_CLIENTE = 'fuente_cliente'
TEXT_OBJECT = 'object'
TEXT_PURPLE = 'purple'
TEXT_ESTADO_2 = 'Estado'
TEXT_PRIORIDAD_2 = 'Prioridad'
TEXT_CAUSA_RAIZ_AUTO = 'causa_raiz_auto'
TEXT_CERRADO = 'cerrado'
TEXT_DURACION_HORAS_NUM = 'duracion_horas_num'
TEXT_ES_ALERTA_AUTO = 'es_alerta_auto'
TEXT_PRIMARY = 'primary'
TEXT_PRODUCTO = 'producto'
TEXT_CLASIFICACION = 'Clasificacion'
TEXT_CREADO_2 = 'Creado'
TEXT_LECTURA = 'Lectura'
TEXT_RESUMEN = 'Resumen'
TEXT_SCORE = 'Score'
TEXT_CERRADO_2 = '_cerrado'
TEXT_CREADO_DT_2 = '_creado_dt'
TEXT_CANAL = 'canal'
TEXT_ATENCIONES = 'Atenciones'
TEXT_EVIDENCIA = 'Evidencia'
TEXT_NUMERO_2 = 'Numero'
TEXT_PREGUNTA = 'Pregunta'
TEXT_RESPONSABLE = 'Responsable'
TEXT_TIPOLOGIA = 'Tipologia'
TEXT_TOTAL = 'Total'
TEXT_PROXIMO_VENCER = '_proximo_vencer'
TEXT_VENCIDO = '_vencido'
TEXT_ADMIN = 'admin'
TEXT_CAUSA = 'causa'
TEXT_CAUSA_COMUN = 'causa_comun'
TEXT_DURACION_SLA_HORAS = 'duracion_sla_horas'
TEXT_LAVENDER = 'lavender'
TEXT_PASSWORD = 'password'
TEXT_F18701 = '#f18701'
TEXT_AMARILLO = 'Amarillo'
TEXT_CERRADOS = 'Cerrados'
TEXT_PERIODO = 'Periodo: '
TEXT_VENCIDOS = 'Vencidos'
TEXT_APLICA_SLA_INCIDENTE = 'aplica_sla_incidente'
TEXT_ASIGNADO = 'asignado'
TEXT_CERRADO_DT = 'cerrado_dt'
TEXT_DURACION_HORAS = 'duracion_horas'
TEXT_EMPRESA = 'empresa'
TEXT_FECHA_VENCIMIENTO_SLA = 'fecha_vencimiento_sla'
TEXT_FLOAT = 'float'
TEXT_OBSERVACIONES_ADICIONALES = 'observaciones_adicionales'
TEXT_OBSERVACIONES_TRABAJO = 'observaciones_trabajo'
TEXT_PRIORIDAD_NORMALIZADA = 'prioridad_normalizada'
TEXT_TIEMPO_RESPUESTA = 'tiempo_respuesta'
TEXT_TIEMPO_RESPUESTA_H = 'tiempo_respuesta_h'
TEXT_TIPO_INCIDENTE_AUTO = 'tipo_incidente_auto'
TEXT_DIAS = ' dias'
TEXT_NO_CUMPLEN = ' | No cumplen: '
TEXT_F35B04 = '#f35b04'
TEXT_CANAL_2 = 'Canal'
TEXT_SEGMENTO = 'Segmento'
TEXT_ALERTA = '_alerta'
TEXT_CLIENTE_EXTERNO = '_cliente_externo'
TEXT_PRIORIDAD_ALTA = '_prioridad_alta'
TEXT_REINCIDENTE = '_reincidente'
TEXT_FIRMA_REINCIDENCIA = '_firma_reincidencia'
TEXT_FIRMA_REINCIDENCIA_LABEL = '_firma_reincidencia_label'
TEXT_ASIGNADO_A = 'asignado_a'
TEXT_BREVE_DESCRIPCION = 'breve_descripcion'
TEXT_CASOS_2 = 'casos'
TEXT_CAUSA_TECNICA = 'causa_tecnica'
TEXT_CODIGO_RESOLUCION = 'codigo_resolucion'
TEXT_COUNT = 'count'
TEXT_DURACION_DIAS_NUM = 'duracion_dias_num'
TEXT_ORANGE = 'orange'
TEXT_TIPO_FALLA = 'tipo_falla'
TEXT_SERVICIO_NEGOCIO = 'servicio_negocio'
TEXT_SOLICITANTE = 'solicitante'
TEXT_TOTAL_CASOS = 'total_casos'
TEXT_TOTAL_INCIDENTES = 'total_incidentes'
TEXT_VIEWER = 'viewer'

UI_PALETTE = {
    "bg": "#ffffff",
    "bg_soft": "#ffffff",
    "surface": "#ffffff",
    "surface_alt": "#ffffff",
    "border": "#e5e7eb",

    "text": "#141414",
    "muted": "#5a5151",

    TEXT_PRIMARY: TEXT_F35B04,
    "primary_hover": TEXT_F18701,
    TEXT_ORANGE: TEXT_F18701,
    TEXT_YELLOW: "#f7b801",
    "yellow_soft": "#ffe0a1",
    "mustard": "#d6a21a",
    TEXT_LAVENDER: "#9683ec",
    TEXT_PURPLE: "#5d16a6",

    "red": TEXT_F35B04,
    "red_soft": TEXT_F18701,
    "green": TEXT_F35B04,
    "green_soft": TEXT_F18701,
    "blue": "#9683ec",
    "blue_soft": "#5d16a6",
}

CHART_COLORS = [
    UI_PALETTE[TEXT_PRIMARY],
    UI_PALETTE[TEXT_YELLOW],
    UI_PALETTE[TEXT_ORANGE],
    UI_PALETTE[TEXT_LAVENDER],
    UI_PALETTE[TEXT_PURPLE],
    UI_PALETTE["text"],
]

PRODUCT_PIE_COLORS = [
    "#0b3a78",
    "#5a7f45",
    "#971531",
    "#ff5a00",
    "#d89a00",
    "#8f9294",
]

CACHE_TTL_SEGUNDOS = 300
DATAFRAME_DISPLAY_LIMIT = 1000
DATAFRAME_PAGE_SIZE = 50
SLA_CASOS_HORAS = 36


def limpiar_cache_datos():
    st.cache_data.clear()


def dataframe_liviano(df, limite=DATAFRAME_DISPLAY_LIMIT, height=None):
    if len(df) > limite:
        st.caption(
            f"Mostrando {limite} de {len(df)} registros para mantener la vista agil. "
            "Aplica filtros para revisar un subconjunto mas pequeno."
        )
        df = df.head(limite)
    kwargs = {
        "use_container_width": True,
        "hide_index": True,
    }
    if height is not None:
        kwargs["height"] = height
    st.dataframe(df, **kwargs)


def cambiar_pagina_dataframe(page_key, delta):
    st.session_state[page_key] = st.session_state.get(page_key, 0) + delta


def dataframe_paginado(df, key, filas_por_pagina=DATAFRAME_PAGE_SIZE, height=None, reset_token=None):
    total = len(df)
    page_key = f"{key}_pagina"
    token_key = f"{key}_token"
    if st.session_state.get(token_key) != reset_token:
        st.session_state[token_key] = reset_token
        st.session_state[page_key] = 0

    total_paginas = max(1, (total + filas_por_pagina - 1) // filas_por_pagina)
    pagina_actual = min(max(int(st.session_state.get(page_key, 0)), 0), total_paginas - 1)
    st.session_state[page_key] = pagina_actual

    inicio = pagina_actual * filas_por_pagina
    fin = min(inicio + filas_por_pagina, total)
    visible = df.iloc[inicio:fin] if total else df

    kwargs = {
        "use_container_width": True,
        "hide_index": True,
    }
    if height is not None:
        kwargs["height"] = height
    st.dataframe(visible, **kwargs)

    col_info, col_anterior, col_siguiente = st.columns([2.4, 1, 1])
    with col_info:
        if total:
            st.caption(
                f"Mostrando {inicio + 1}-{fin} de {total} registros | "
                f"Pagina {pagina_actual + 1} de {total_paginas}"
            )
        else:
            st.caption("Sin registros para mostrar.")
    with col_anterior:
        st.button(
            "Anterior",
            key=f"{key}_anterior",
            disabled=pagina_actual <= 0,
            on_click=cambiar_pagina_dataframe,
            args=(page_key, -1),
            use_container_width=True,
        )
    with col_siguiente:
        st.button(
            "Siguiente",
            key=f"{key}_siguiente",
            disabled=pagina_actual >= total_paginas - 1,
            on_click=cambiar_pagina_dataframe,
            args=(page_key, 1),
            use_container_width=True,
        )


@st.cache_data(show_spinner=False)
def dataframe_csv_bytes(df):
    return df.to_csv(index=False).encode("utf-8-sig")


def nombre_archivo_seguro(texto):
    limpio = re.sub(r"[^A-Za-z0-9_-]+", "_", str(texto).strip().lower())
    limpio = re.sub(r"_+", "_", limpio).strip("_")
    return limpio or "datos"


def render_descarga_dataframe(df, key, nombre_base, periodo_label):
    if df.empty:
        return
    archivo_periodo = nombre_archivo_seguro(periodo_label)
    st.download_button(
        "Descargar data completa",
        data=dataframe_csv_bytes(df),
        file_name=f"{nombre_base}_{archivo_periodo}.csv",
        mime="text/csv",
        key=key,
        help="Incluye todos los registros del periodo y filtros actuales, no solo la pagina visible.",
    )


@st.cache_data(ttl=CACHE_TTL_SEGUNDOS, show_spinner=False)
def cargar_casos_cache():
    return load_casos()


@st.cache_data(ttl=CACHE_TTL_SEGUNDOS, show_spinner=False)
def cargar_incidentes_cache():
    return load_incidentes()


@st.cache_data(ttl=CACHE_TTL_SEGUNDOS, show_spinner=False)
def cargar_incidentes_sla_cache():
    return agregar_campos_sla_incidentes(cargar_incidentes_cache())


@st.cache_data(ttl=CACHE_TTL_SEGUNDOS, show_spinner=False)
def cargar_casos_anios_cache(anios):
    return load_casos_anios(list(anios))


@st.cache_data(ttl=CACHE_TTL_SEGUNDOS, show_spinner=False)
def cargar_incidentes_anios_cache(anios):
    return load_incidentes_anios(list(anios))


@st.cache_data(ttl=CACHE_TTL_SEGUNDOS, show_spinner=False)
def cargar_meses_disponibles_cache(tabla):
    return obtener_meses_disponibles(tabla)


@st.cache_data(ttl=CACHE_TTL_SEGUNDOS, show_spinner=False)
def cargar_ultimo_mes_disponible_cache(tabla):
    return obtener_ultimo_mes_disponible(tabla)


@st.cache_data(ttl=CACHE_TTL_SEGUNDOS, show_spinner=False)
def cargar_meses_disponibles_multi_cache(tablas):
    meses = set()
    for tabla in tablas:
        meses.update(obtener_meses_disponibles(tabla))
    return sorted(meses)


@st.cache_data(ttl=CACHE_TTL_SEGUNDOS, show_spinner=False)
def cargar_casos_filtrados_cache(anio=None, mes=None, cliente="", estado="", servicio="", tipificacion=""):
    return load_casos_filtrados(
        anio=anio,
        mes=mes,
        cliente=cliente,
        estado=estado,
        servicio=servicio,
        tipificacion=tipificacion,
    )


@st.cache_data(ttl=CACHE_TTL_SEGUNDOS, show_spinner=False)
def cargar_incidentes_filtrados_cache(
    anio=None,
    mes=None,
    cliente="",
    estado="",
    servicio="",
    tipificacion="",
    es_alerta="",
):
    return load_incidentes_filtrados(
        anio=anio,
        mes=mes,
        cliente=cliente,
        estado=estado,
        servicio=servicio,
        tipificacion=tipificacion,
        es_alerta=es_alerta,
    )


@st.cache_data(ttl=CACHE_TTL_SEGUNDOS, show_spinner=False)
def cargar_incidentes_sla_filtrados_cache(
    anio=None,
    mes=None,
    cliente="",
    estado="",
    servicio="",
    tipificacion="",
    es_alerta="",
):
    return agregar_campos_sla_incidentes(
        cargar_incidentes_filtrados_cache(anio, mes, cliente, estado, servicio, tipificacion, es_alerta)
    )

# Constantes de etiquetas usadas en tablas, filtros y graficas
TIPIFICACION_REDIRECCIONAMIENTO_AGENDA = "9 - Redireccionamiento Agenda"
TIPIFICACION_CLIENTE_NO_ASISTIO = "10 - Cliente no asistio"
TIPIFICACION_INCIDENTE_CLIENTE_EXTERNO = "Incidente Cliente Externo"
TIPIFICACION_INCIDENTE_INTERNO = "Incidente Interno"
TIPIFICACION_CASO_CLIENTE_EXTERNO = "Caso Cliente Externo"

CLIENTE_SICOV = "SICOV"
CLIENTE_TELEFONICA = "TELEFONICA"
CLIENTE_TUYA = "TUYA"
CLIENTE_SUFI_BANCOLOMBIA = "SUFI BANCOLOMBIA"
CLIENTE_RCI_COLOMBIA = "RCI COLOMBIA S.A COMPAÃ‘ÃA DE FINANCIAMIENTO"
CLIENTE_PORVENIR = "PORVENIR"
CLIENTE_MIBANCO = "MIBANCO S.A."
CLIENTE_BBVA = "BBVA"
CLIENTE_BANCOOMEVA = "BANCOOMEVA"
CLIENTE_BANCOLOMBIA = "BANCOLOMBIA"
CLIENTE_BANCO_POPULAR = "BANCO POPULAR"
CLIENTE_BANCO_FALABELLA = "BANCO FALABELLA"
CLIENTE_BANCO_DE_OCCIDENTE = "BANCO DE OCCIDENTE"
CLIENTE_BANCO_DAVIVIENDA = "BANCO DAVIVIENDA S.A."
CLIENTE_BANCO_CAJA_SOCIAL = "BANCO CAJA SOCIAL"
CLIENTE_AV_VILLAS = "AV VILLAS"
CLIENTE_FALLABELLA = "FALLABELLA"
CLIENTE_COLPENSIONES = "COLPENSIONES"
CLIENTE_CLARO = "CLARO"

PANDAS_DATETIME_DTYPE = "datetime64[ns]"
SIN_DATO = "Sin dato"
SIN_DURACION = "Sin duracion"
SIN_CUENTA = "Sin cuenta"
SIN_SERVICIO = "Sin servicio"
PRIORIDAD_ALTA = "Prioridad alta"
ESTADO_SLA_CUMPLE = "Cumple"
ESTADO_SLA_NO_CUMPLE = "No cumple"

COL_ESTADO_ATENCION = "Estado atencion"
COL_TOTAL_ATENCIONES = "Total atenciones"
COL_SLA_CASOS_PCT = "SLA casos %"
COL_SLA_INCIDENTES_PCT = "SLA incidentes %"
COL_ALERTAS_INCIDENTES = "Alertas incidentes"
COL_CASOS_SIN_CAUSA = "Casos sin causa"
COL_PRODUCTO_PRINCIPAL = "Producto principal"
COL_CAUSA_INCIDENTE_PRINCIPAL = "Causa incidente principal"
COL_ULTIMA_ATENCION = "Ultima atencion"
COL_VENCIMIENTO_SLA = "Vencimiento SLA"
COL_CANAL_AGRUPADO = "Canal agrupado"
COL_MOTIVO_INFERIDO = "Motivo inferido"
COL_CLIENTE_AGENDA = "Cliente agenda"
COL_CICLO_CLIENTE = "Ciclo cliente"
COL_CLIENTE_RECURRENTE_AGENDA = "Cliente recurrente agenda"
COL_CASOS_HISTORICOS_CLIENTE = "Casos historicos cliente"
COL_AGENDAS_HISTORICAS_CLIENTE = "Agendas historicas cliente"
COL_PRIMERA_ATENCION_CLIENTE = "Primera atencion cliente"
COL_ACCION_SUGERIDA = "Accion sugerida"
COL_CASOS_AGENDA = "Casos agenda"
COL_PRIMERA_AGENDA = "Primera agenda"
COL_ULTIMA_AGENDA = "Ultima agenda"
COL_LECTURA_EJECUTIVA = "Lectura ejecutiva"
COL_FAMILIA_INCIDENTE = "Familia incidente"
COL_EVIDENCIA_INCIDENTE = "Evidencia observada"
COL_CUMPLE_SLA = "Cumple SLA"
COL_NO_CUMPLE_SLA = "No cumple SLA"
COL_PROM_HORAS = "Prom. horas"
COL_PROM_DIAS = "Prom. dias"
COL_CAUSA_RAIZ = "Causa raiz"
COL_MOTIVO_CASO = "Motivo del caso"
COL_PRINCIPAL_TIPIFICACION = "Principal tipificacion"
COL_PRINCIPAL_SOPORTE = "Tipologia principal"
COL_PRINCIPAL_CAUSA_COMUN = "Principal causa comun"
COL_SLA_OBJETIVO = "SLA objetivo"
COL_SLA_OBJETIVO_H = "SLA objetivo h"
COL_MAX_HORAS = "Max. horas"
COL_MAX_DIAS = "Max. dias"
COL_CLIENTES_DISTINTOS = "Clientes"
COL_PRINCIPALES_CLIENTES = "Principales clientes"
COL_FIRMA_REINCIDENCIA = "Firma reincidencia"
COL_INCIDENTES_REINCIDENTES = "Incidentes reincidentes"
COL_REINCIDENTE = "Reincidente"
COL_REINCIDENTE_AGENDA = "Reincidente agenda"
COL_CASOS_REINCIDENTES_AGENDA = "Casos reincidentes agenda"
COL_REINCIDENCIA_AGENDAMIENTO = "Reincidencia agendamiento %"
MENU_CLIENTES_CLAVE = "Clientes clave"
MENU_KPI_CLIENTES_CLAVE = "KPI Clientes Clave"
MENU_DASHBOARD_CASOS_SOPORTE = "Dashboard Casos Soporte"
MENU_KPI_CASOS_CLIENTE_EXTERNO = "KPI Casos Cliente Externo"
MENU_KPI_INCIDENTES = "KPI Incidentes"
MENU_KPI_COMPARATIVO_ANUAL = "KPI 2025 vs 2026"
MENU_REINCIDENCIAS_PROBLEMAS = "Reincidencias y problemas sugeridos"
MENU_SEGUIMIENTO_INCIDENTES_VIEWER = "Seguimiento incidentes"
MENU_SEGUIMIENTO_INCIDENTES_ADMIN = "Seguimiento Incidentes"
MENU_SEGUIMIENTO_RPOST = "Seguimiento de RPost"
LABEL_CASOS_CLIENTE_EXTERNO = "Casos cliente externo"
TEXT_TIPOLOGIA_SOPORTE = "Tipologia caso"
SOPORTE_USO = "Soporte Uso"
ENVIO_AGENDA_MANUAL_USO = "Envio agenda / manual de uso"
SOLICITUDES_CASOS = "Solicitudes"
INCIDENTES_CASOS = "Incidentes"
CASE_TYPOLOGY_CACHE_VERSION = "casos-tipologia-agenda-directa-v2"
KEY_CLIENT_CASE_YEAR_BASE = 2025
KEY_CLIENT_CASE_YEAR_FOCUS = 2026
KEY_CLIENT_CASE_MONTHS_FOCUS = [3, 4, 6]
ANS_INCIDENT_YEAR_BASE = 2025
ANS_INCIDENT_YEAR_FOCUS = 2026
ANS_INCIDENT_MONTHS_FOCUS = [3, 4, 5]
ANS_INCIDENT_REAL_TYPES = [TIPIFICACION_INCIDENTE_CLIENTE_EXTERNO, TIPIFICACION_INCIDENTE_INTERNO]
MONTH_NAMES_ES = {
    1: "Enero",
    2: "Febrero",
    3: "Marzo",
    4: "Abril",
    5: "Mayo",
    6: "Junio",
    7: "Julio",
    8: "Agosto",
    9: "Septiembre",
    10: "Octubre",
    11: "Noviembre",
    12: "Diciembre",
}


def parse_mes_periodo(valor):
    if not valor or not isinstance(valor, str) or "-" not in valor:
        return None, None
    anio, mes = valor.split("-", 1)
    return int(anio), int(mes)


def etiqueta_mes_periodo(anio, mes):
    if mes in (None, "", TEXT_TODOS):
        return f"{anio} completo" if anio else TEXT_TODOS
    return f"{MONTH_NAMES_ES.get(int(mes), str(mes))} {int(anio)}"


def periodo_key_sql(anio, mes):
    if anio and mes:
        return f"{int(anio):04d}-{int(mes):02d}"
    return TEXT_TODOS


def selector_periodo_sql(tabla, key, label_anio="AÃ±o", label_mes="Mes"):
    meses = cargar_meses_disponibles_cache(tabla)
    return selector_periodo_desde_meses(meses, key, label_anio, label_mes)


def selector_periodo_multi_sql(tablas, key, label_anio="AÃ±o", label_mes="Mes"):
    meses = cargar_meses_disponibles_multi_cache(tuple(tablas))
    return selector_periodo_desde_meses(meses, key, label_anio, label_mes)


def selector_periodo_desde_meses(meses, key, label_anio="AÃ±o", label_mes="Mes"):
    if not meses:
        st.caption("No hay fechas validas para filtrar por periodo.")
        return None, None, TEXT_TODOS

    ultimo_mes = meses[-1]
    ultimo_anio, ultimo_mes_num = parse_mes_periodo(ultimo_mes)
    anios = sorted({parse_mes_periodo(mes)[0] for mes in meses if parse_mes_periodo(mes)[0]})

    col_anio, col_mes = st.columns([1, 1])
    with col_anio:
        anio = st.selectbox(
            label_anio,
            anios,
            index=anios.index(ultimo_anio) if ultimo_anio in anios else len(anios) - 1,
            key=f"{key}_anio",
        )

    meses_anio = sorted(
        {
            parse_mes_periodo(mes)[1]
            for mes in meses
            if parse_mes_periodo(mes)[0] == int(anio)
        }
    )
    opciones_mes = [TEXT_TODOS] + [MONTH_NAMES_ES.get(mes, f"{mes:02d}") for mes in meses_anio]
    mapa_mes = {MONTH_NAMES_ES.get(mes, f"{mes:02d}"): mes for mes in meses_anio}
    mes_default = ultimo_mes_num if int(anio) == int(ultimo_anio) else (meses_anio[-1] if meses_anio else None)
    etiqueta_default = MONTH_NAMES_ES.get(mes_default, f"{mes_default:02d}") if mes_default else TEXT_TODOS

    with col_mes:
        mes_label = st.selectbox(
            label_mes,
            opciones_mes,
            index=opciones_mes.index(etiqueta_default) if etiqueta_default in opciones_mes else 0,
            key=f"{key}_mes",
        )

    mes = mapa_mes.get(mes_label)
    return int(anio), mes, etiqueta_mes_periodo(anio, mes)


def periodo_sql_valido(anio, etiqueta="datos"):
    if anio is not None:
        return True
    st.info(f"No hay fechas validas para consultar {etiqueta}.")
    return False

CASE_TIPIFICATION_RENAMES = {
    "8 - Instalaciones": TIPIFICACION_REDIRECCIONAMIENTO_AGENDA,
    "8 - Agenda Instalaciones IVR": TIPIFICACION_REDIRECCIONAMIENTO_AGENDA,
    "9 - Agenda Instalaciones IVR": TIPIFICACION_REDIRECCIONAMIENTO_AGENDA,
    "9 - Redireccionamiento Agenda IVR": TIPIFICACION_REDIRECCIONAMIENTO_AGENDA,
    "9 - Agenda sin evidencia": TIPIFICACION_CLIENTE_NO_ASISTIO,
    "10 - Agenda sin evidencia": TIPIFICACION_CLIENTE_NO_ASISTIO,
}

CASE_FIELDS_SEGUIMIENTO_RPOST = [
    TEXT_DESCRIPCION_2,
    TEXT_CAUSA,
    TEXT_CODIGO_RESOLUCION,
    "notas_resolucion",
    TEXT_OBSERVACIONES_ADICIONALES,
    TEXT_OBSERVACIONES_TRABAJO,
    TEXT_PRODUCTO,
    TEXT_CUENTA,
]

INCIDENT_FIELDS_SEGUIMIENTO_RPOST = [
    TEXT_EMPRESA,
    TEXT_SOLICITANTE,
    TEXT_BREVE_DESCRIPCION,
    "categoria",
    TEXT_DESCRIPCION_2,
    TEXT_OBSERVACIONES_TRABAJO,
    TEXT_OBSERVACIONES_ADICIONALES,
    "actualizaciones",
    "impacto",
    "lista_notas_trabajo",
    TEXT_SERVICIO_NEGOCIO,
    TEXT_TIPO_FALLA,
    TEXT_CAUSA_RAIZ_AUTO,
    TEXT_TIPO_INCIDENTE_AUTO,
    TEXT_TIPIFICACION_AUTO,
]

PATRONES_NO_RECIBIO_ACUSE = [
    r"\bno\s+(?:se\s+)?(?:ha\s+|han\s+)?recib(?:io|ido|ieron|e|en|imos)\s+(?:el\s+|los\s+|la\s+|las\s+)?acuses?(?:\s+de\s+recibo)?\b",
    r"\bno\s+(?:se\s+)?(?:esta|estan|estamos)\s+recibiendo\s+(?:el\s+|los\s+|la\s+|las\s+)?acuses?(?:\s+de\s+recibo)?\b",
    r"\bno\s+(?:le\s+|les\s+)?(?:llego|llegaron|llega|llegan)\s+(?:el\s+|los\s+|la\s+|las\s+)?acuses?(?:\s+de\s+recibo)?\b",
    r"\b(?:cliente|usuario|solicitante)\s+no\s+(?:ha\s+)?recib(?:io|ido|e)\s+(?:el\s+|los\s+)?acuses?(?:\s+de\s+recibo)?\b",
]

CASE_RPOST_RELATION_RULES = [
    ("RPost", ["rpost"]),
    ("Certimail / Certicmal", ["certimail", "certi mail", "certicmal", "certimall"]),
]

CASE_SUPPORT_TYPOLOGY_GUIDE = [
    {
        TEXT_TIPOLOGIA_SOPORTE: ENVIO_AGENDA_MANUAL_USO,
        TEXT_DESCRIPCION: (
            "Casos donde el agente direcciona al usuario a agenda, cita oficial o canal de agendamiento. "
            "Incluye los casos donde ademas envia manual, guia, instructivo, PDF o adjunto de uso."
        ),
    },
    {
        TEXT_TIPOLOGIA_SOPORTE: SOPORTE_USO,
        TEXT_DESCRIPCION: (
            "Acompanamiento, configuracion, instalacion, acceso, recuperacion, certificados, firma digital, "
            "token fisico, token virtual, errores funcionales, activaciones, agenda y apoyo para usar el servicio."
        ),
    },
    {
        TEXT_TIPOLOGIA_SOPORTE: SOLICITUDES_CASOS,
        TEXT_DESCRIPCION: (
            "Requerimientos administrativos, operativos o comerciales: renovaciones, cambios de informacion, "
            "solicitudes de certificados, pagos, biometria, consultas administrativas y requerimientos especiales."
        ),
    },
    {
        TEXT_TIPOLOGIA_SOPORTE: INCIDENTES_CASOS,
        TEXT_DESCRIPCION: (
            "Afectaciones con indisponibilidad total o parcial, degradacion del servicio, afectaciones masivas, "
            "incidentes formales o incidentes asociados a proveedores."
        ),
    },
]

CASE_SUPPORT_TYPOLOGY_ORDER = [item[TEXT_TIPOLOGIA_SOPORTE] for item in CASE_SUPPORT_TYPOLOGY_GUIDE]

CASE_SUPPORT_TOKEN_VIRTUAL_TERMS = [
    "certitoken",
    "certi token",
    "token virtual",
    "token movil",
    "token digital",
    "app token",
    "codigo token",
    "codigo de token",
    "codigo otp",
    "otp",
]

CASE_SUPPORT_TOKEN_FISICO_TERMS = [
    "token",
    "token fisico",
    "epass",
    "e pass",
    "safenet",
    "safe net",
    "token usb",
    "usb token",
    "dispositivo fisico",
    "driver",
    "drivers",
    "no reconoce token",
    "token no reconocido",
    "no detecta token",
    "token no detectado",
]

CASE_TOKEN_FISICO_READING_TERMS = [
    "token fisico",
    "epass",
    "e pass",
    "safenet",
    "safe net",
    "token usb",
    "usb token",
    "dispositivo fisico",
    "driver",
    "drivers",
    "no reconoce token",
    "token no reconocido",
    "no detecta token",
    "token no detectado",
    "token gris",
    "feitian",
]

CASE_SIGNING_PROBLEM_TERMS = [
    "problema para firmar",
    "problemas para firmar",
    "dificultad para firmar",
    "dificultades para firmar",
    "error al firmar",
    "error firmando",
    "no firma",
    "no puede firmar",
    "no permite firmar",
    "falla al firmar",
    "fallas al firmar",
    "firmar documento",
    "validacion de firma",
    "validar firma",
    "firma invalida",
]

CASE_ACUSES_TERMS = [
    "acuse",
    "acuses",
    "acusese",
    "acuse de recibo",
    "acuses de recibo",
    "no recibio acuse",
    "no recibe acuse",
    "no llegan acuses",
    "no se reciben acuses",
    "trazabilidad de envio",
    "trazabilidad de envios",
]

CASE_SUPPORT_SECURITY_TERMS = [
    "phishing",
    "pishing",
    "correo sospechoso",
    "suplantacion",
    "fraude",
    "seguridad",
    "malicioso",
    "spam",
]

CASE_SUPPORT_CERT_FIRMA_ACUSES_TERMS = [
    "certificado",
    "certificados",
    "firma",
    "firmar",
    "certifirma",
    "validacion de firma",
    "validar firma",
    "firma invalida",
    "error al firmar",
    "no firma",
    "renovacion",
    "activar certificado",
    "activacion certificado",
    "descarga certificado",
    "descargar certificado",
    "certimail",
    "certi mail",
    "certicmal",
    "rpost",
    "acuse",
    "acuses",
    "acusese",
    "acuse de recibo",
    "acuses de recibo",
    "trazabilidad de envio",
    "trazabilidad de envios",
]

CASE_SUPPORT_PLATFORM_ACCESS_TERMS = [
    "portal",
    "usuario",
    "contrasena",
    "password",
    "clave",
    "permiso",
    "permisos",
    "bloqueo",
    "bloqueado",
    "autenticacion",
    "login",
    "acceso",
    "navegador",
    "configuracion",
    "java",
    "autofirma",
    "adobe",
    "lentitud",
    "indisponibilidad",
    "error",
    "falla",
]

CASE_MANUAL_USAGE_TERMS = [
    "manual",
    "link manual",
    "guia",
    "instructivo",
    "pdf",
    "archivo adjunto",
    "adjunto",
    "instalacion_firma_digital",
    "instalacion firma digital",
    "feitian",
    "token gris",
]

CASE_AGENDA_DIRECT_TERMS = [
    "agenda",
    "agendamiento",
    "agenda directa",
    "cita",
    "programa tu cita",
    "cita oficial",
    "canal de agendamiento",
    "zcal",
    "agendate",
    "elige el horario",
    "tecnico especializado",
    "paso a paso",
]

CASE_REQUEST_TERMS = [
    "renovacion",
    "renovar",
    "cambio de informacion",
    "cambio informacion",
    "actualizacion de datos",
    "actualizacion informacion",
    "modificacion de datos",
    "solicitud de certificado",
    "solicitud certificado",
    "certificado solicitado",
    "pago",
    "pagos",
    "pagar",
    "biometria",
    "consulta administrativa",
    "administrativo",
    "administrativa",
    "comercial",
    "requerimiento especial",
    "requerimientos especiales",
    "factura",
    "facturacion",
    "cotizacion",
    "compra",
]

CASE_INCIDENT_TERMS = [
    "incidente",
    "indisponibilidad",
    "no disponible",
    "caida",
    "degradacion",
    "afectacion masiva",
    "masivo",
    "interrupcion",
    "servicio caido",
    "fuera de servicio",
    "proveedor",
    "phishing",
    "pishing",
    "suplantacion",
    "fraude",
    "correo sospechoso",
    "seguridad",
]

def normalizar_tipificaciones_casos_df(df):
    if df.empty or TEXT_TIPIFICACION_2 not in df.columns:
        return df
    trabajo = df.copy()
    trabajo[TEXT_TIPIFICACION_2] = trabajo[TEXT_TIPIFICACION_2].replace(CASE_TIPIFICATION_RENAMES)
    return trabajo


def preparar_fechas_dashboard(df, columna=TEXT_CREADO):
    trabajo = df.copy()
    trabajo[TEXT_CREADO_DT_DASHBOARD] = pd.to_datetime(
        trabajo[columna].apply(normalizar_fecha),
        errors=TEXT_COERCE,
    )
    return trabajo


def meses_disponibles(df, columna_dt=TEXT_CREADO_DT_DASHBOARD):
    if df.empty or columna_dt not in df.columns:
        return []
    meses = df[columna_dt].dropna().dt.to_period("M").astype(str).sort_values().unique().tolist()
    return meses


def selector_mes_dashboard(
    df,
    key,
    columna_dt=TEXT_CREADO_DT_DASHBOARD,
    label="Mes del dashboard",
    incluir_todos=True,
):
    meses = meses_disponibles(df, columna_dt)
    if not meses:
        st.caption("No hay fechas validas para filtrar por mes.")
        return TEXT_TODOS
    opciones = ([TEXT_TODOS] if incluir_todos else []) + meses
    return st.selectbox(label, opciones, index=len(opciones) - 1, key=key)


def serie_servicio_filtro(df, columna_servicio):
    if df.empty or columna_servicio not in df.columns:
        return pd.Series(dtype=TEXT_OBJECT)
    return df[columna_servicio].fillna("").astype(str).str.strip().replace("", SIN_SERVICIO)


def opciones_filtro_servicio(df, columna_servicio):
    servicios = serie_servicio_filtro(df, columna_servicio)
    if servicios.empty:
        return []
    return sorted(servicios.dropna().unique().tolist())


def filtrar_por_servicio(df, columna_servicio, servicio):
    if servicio == TEXT_TODOS or df.empty or columna_servicio not in df.columns:
        return df
    servicios = serie_servicio_filtro(df, columna_servicio)
    return df[servicios == servicio].copy()


def filtrar_mes_dashboard(df, mes, columna_dt=TEXT_CREADO_DT_DASHBOARD):
    if mes == TEXT_TODOS or df.empty or columna_dt not in df.columns:
        return df
    return df[df[columna_dt].dt.to_period("M").astype(str) == mes].copy()


CASE_TIPIFICATION_GUIDE = [
    {
        TEXT_TIPIFICACION: "1 - phishing",
        TEXT_DESCRIPCION: "Correos sospechosos, suplantacion, enlaces fraudulentos o reportes de phishing.",
    },
    {
        TEXT_TIPIFICACION: "2 - Soporte Uso",
        TEXT_DESCRIPCION: "Dudas de uso, acompanamiento funcional, configuracion, orientacion y paso a paso.",
    },
    {
        TEXT_TIPIFICACION: "3 - Soporte Falla",
        TEXT_DESCRIPCION: "Errores, fallas, caidas, lentitud, indisponibilidad o afectaciones tecnicas.",
    },
    {
        TEXT_TIPIFICACION: "4 - solicitudes",
        TEXT_DESCRIPCION: "Solicitudes operativas o comerciales como certificados, biometria, pagos u otros tramites.",
    },
    {
        TEXT_TIPIFICACION: "5 - incidente",
        TEXT_DESCRIPCION: "Casos marcados como incidente o con afectacion operativa reportada.",
    },
    {
        TEXT_TIPIFICACION: "6 - Plataformas Ext",
        TEXT_DESCRIPCION: "Problemas relacionados con plataformas externas como Adobe, Autofirma o DocuSign.",
    },
    {
        TEXT_TIPIFICACION: "7 - No Aplica",
        TEXT_DESCRIPCION: "Casos sin informacion suficiente o que no encajan en las reglas definidas.",
    },
    {
        TEXT_TIPIFICACION: TIPIFICACION_REDIRECCIONAMIENTO_AGENDA,
        TEXT_DESCRIPCION: "Casos detectados como instalacion que deben redirigirse a agenda.",
    },
    {
        TEXT_TIPIFICACION: TIPIFICACION_CLIENTE_NO_ASISTIO,
        TEXT_DESCRIPCION: "Cliente no conectado, no ingreso o no se presento a la agenda.",
    },
]

AGENDA_CASE_TIPIFICATION = TIPIFICACION_REDIRECCIONAMIENTO_AGENDA
AGENDA_MOTIVO_TOKEN_FISICO = "Token fisico / ePass"
AGENDA_DIRECT_CHANNEL_HINTS = ["calendario"]
AGENDA_HELP_DESK_CHANNELS = {"web", "telefono", "correo electronico", "en persona", ""}
AGENDA_REASON_RULES = [
    (
        AGENDA_MOTIVO_TOKEN_FISICO,
        [
            "token fisico",
            "token",
            "epass",
            "safenet",
            "usb",
            "dispositivo",
            "no reconoce",
        ],
    ),
    (
        "Activacion o descarga de certificado",
        [
            "activacion",
            "activar",
            "descarga",
            "descargar",
            "certificado",
            ".cer",
            "cer ",
        ],
    ),
    (
        "Instalacion o reinstalacion",
        [
            "instalacion",
            "instalar",
            "instale",
            "reinstalacion",
            "reinstalar",
            "configuracion",
            "configurar",
        ],
    ),
    (
        "Firma o pruebas de uso",
        [
            "firma",
            "firmar",
            "certifirma",
            "prueba de firma",
            "pruebas de firma",
        ],
    ),
    (
        "Cita o redireccionamiento a agenda",
        [
            "agenda",
            "agendar",
            "agendamiento",
            "cita",
            "meetings.hubspot",
        ],
    ),
]

CASE_COMMON_CAUSE_RULES = [
    (
        "Token fisico / ePass",
        ["token fisico", "token", "epass", "safenet", "usb", "dispositivo"],
    ),
    (
        "Firma digital",
        ["firma", "firmar", "certifirma", "validar firma", "falla en la firma"],
    ),
    (
        "Instalacion o configuracion",
        ["instalacion", "reinstalacion", "instalar", "configuracion", "configurar", "parametrizacion"],
    ),
    (
        "Activacion o descarga de certificado",
        ["activacion", "activar", "descarga", "descargar", "certificado", ".cer"],
    ),
    (
        "Error o falla tecnica",
        ["error", "falla", "no funciona", "novedad", "inconveniente", "problema", "persiste"],
    ),
    (
        "Solicitud operativa",
        ["solicitud", "biometria", "pago", "pagar", "orden", "actualizacion", "cambio"],
    ),
    (
        "Phishing o seguridad",
        ["phishing", "correo sospechoso", "suplantacion", "fraude"],
    ),
    (
        "Plataforma externa",
        ["adobe", "acrobat", "autofirma", "docusign"],
    ),
]

CASE_CAUSE_DETAIL_RULES = [
    ("Fallas al firmar", ["falla en la firma", "error al firmar", "no firma", "no puede firmar", "firmar"]),
    ("Validacion de firma", ["validar firma", "validacion de firma", "validar la firma", "firma invalida"]),
    ("Pruebas de firma", ["prueba de firma", "pruebas de firma", "pruebas de funcionamiento"]),
    ("Token fisico/ePass", ["token", "epass", "safenet", "usb", "dispositivo"]),
    ("Instalacion o configuracion", ["instalacion", "reinstalacion", "configuracion", "configurar"]),
    ("Activacion o descarga", ["activacion", "activar", "descarga", "descargar"]),
    ("Certificado", ["certificado", ".cer"]),
    ("Acompanamiento de uso", ["acompanamiento", "asesoria", "orientacion", "paso a paso", "soporte"]),
    ("Solicitud operativa", ["solicitud", "biometria", "pago", "orden", "actualizacion"]),
]

NOC_TIPIFICATION = "Alertas y Consultas NOC"

INCIDENT_TIPIFICATION_GUIDE = {
    NOC_TIPIFICATION: "Alertas y consultas gestionadas por NOC. Tienen el mismo objetivo operativo de atencion.",
    TIPIFICACION_INCIDENTE_CLIENTE_EXTERNO: "Afectaciones reportadas por clientes externos o con impacto hacia clientes externos.",
    TIPIFICACION_INCIDENTE_INTERNO: "Afectaciones de infraestructura, red, plataforma o procesos internos.",
    TIPIFICACION_CASO_CLIENTE_EXTERNO: "El cliente lo reporta como incidente, pero operativamente se gestiona como caso externo.",
}

INCIDENT_TIPIFICATION_ORDER = [
    NOC_TIPIFICATION,
    TIPIFICACION_INCIDENTE_CLIENTE_EXTERNO,
    TIPIFICACION_INCIDENTE_INTERNO,
    TIPIFICACION_CASO_CLIENTE_EXTERNO,
]

CLIENTES_CLAVE = [
    CLIENTE_SICOV,
    CLIENTE_TELEFONICA,
    CLIENTE_TUYA,
    CLIENTE_SUFI_BANCOLOMBIA,
    CLIENTE_RCI_COLOMBIA,
    CLIENTE_PORVENIR,
    CLIENTE_MIBANCO,
    CLIENTE_BBVA,
    CLIENTE_BANCOOMEVA,
    CLIENTE_BANCOLOMBIA,
    CLIENTE_BANCO_POPULAR,
    CLIENTE_BANCO_FALABELLA,
    CLIENTE_BANCO_DE_OCCIDENTE,
    CLIENTE_BANCO_DAVIVIENDA,
    CLIENTE_BANCO_CAJA_SOCIAL,
    CLIENTE_AV_VILLAS,
    CLIENTE_FALLABELLA,
    CLIENTE_COLPENSIONES,
    CLIENTE_CLARO,
    "Coopcentral",
]

CLIENTES_CLAVE_ALIASES = {
    CLIENTE_SICOV: [CLIENTE_SICOV],
    CLIENTE_TELEFONICA: [CLIENTE_TELEFONICA, "TELEFÃ“NICA", "MOVISTAR"],
    CLIENTE_TUYA: [CLIENTE_TUYA, "TUYA S.A"],
    CLIENTE_SUFI_BANCOLOMBIA: [CLIENTE_SUFI_BANCOLOMBIA, "SUFI"],
    CLIENTE_RCI_COLOMBIA: [
        CLIENTE_RCI_COLOMBIA,
        "RCI COLOMBIA",
        "RCI",
    ],
    CLIENTE_PORVENIR: [CLIENTE_PORVENIR],
    CLIENTE_MIBANCO: [CLIENTE_MIBANCO, "MIBANCO"],
    CLIENTE_BBVA: [CLIENTE_BBVA],
    CLIENTE_BANCOOMEVA: [CLIENTE_BANCOOMEVA, "BANCOOMEVA S.A"],
    CLIENTE_BANCOLOMBIA: [CLIENTE_BANCOLOMBIA, "BANCOLOMBIA S.A"],
    CLIENTE_BANCO_POPULAR: [CLIENTE_BANCO_POPULAR],
    CLIENTE_BANCO_FALABELLA: [CLIENTE_BANCO_FALABELLA],
    CLIENTE_BANCO_DE_OCCIDENTE: [CLIENTE_BANCO_DE_OCCIDENTE],
    CLIENTE_BANCO_DAVIVIENDA: [CLIENTE_BANCO_DAVIVIENDA, "BANCO DAVIVIENDA", "DAVIVIENDA"],
    CLIENTE_BANCO_CAJA_SOCIAL: [CLIENTE_BANCO_CAJA_SOCIAL, "CAJA SOCIAL"],
    CLIENTE_AV_VILLAS: [CLIENTE_AV_VILLAS, "BANCO AV VILLAS"],
    CLIENTE_FALLABELLA: [CLIENTE_FALLABELLA, "FALABELLA"],
    CLIENTE_COLPENSIONES: [CLIENTE_COLPENSIONES, "COLPENSIONES", "COLPEN"],
    CLIENTE_CLARO: [CLIENTE_CLARO, "COMCEL"],
    "Coopcentral": ["COOPCENTRAL", "BANCO COOPCENTRAL"],
}


def aplicar_tema_visual():
    st.markdown(
        f"""
        <style>
        :root {{
            --bg: {UI_PALETTE["bg"]};
            --bg-soft: {UI_PALETTE["bg_soft"]};
            --surface: {UI_PALETTE["surface"]};
            --surface-alt: {UI_PALETTE["surface_alt"]};
            --border: {UI_PALETTE["border"]};
            --text: {UI_PALETTE["text"]};
            --muted: {UI_PALETTE["muted"]};
            --primary: {UI_PALETTE[TEXT_PRIMARY]};
            --primary-hover: {UI_PALETTE["primary_hover"]};
            --orange: {UI_PALETTE[TEXT_ORANGE]};
            --yellow: {UI_PALETTE[TEXT_YELLOW]};
            --yellow-soft: {UI_PALETTE["yellow_soft"]};
            --mustard: {UI_PALETTE["mustard"]};
            --lavender: {UI_PALETTE[TEXT_LAVENDER]};
            --purple: {UI_PALETTE[TEXT_PURPLE]};
            --red: {UI_PALETTE["red"]};
            --red-soft: {UI_PALETTE["red_soft"]};
            --green: {UI_PALETTE["green"]};
            --green-soft: {UI_PALETTE["green_soft"]};
            --blue: {UI_PALETTE["blue"]};
            --blue-soft: {UI_PALETTE["blue_soft"]};
        }}

        html, body, [data-testid="stAppViewContainer"], .stApp {{
            background: var(--bg) !important;
            color: var(--text) !important;
            color-scheme: light !important;
        }}

        [data-testid="stHeader"], [data-testid="stToolbar"], [data-testid="stDecoration"] {{
            background: transparent !important;
        }}

        [data-testid="stSidebar"] {{
            background: var(--surface) !important;
            border-right: 1px solid var(--border);
        }}

        .block-container {{
            padding-top: 1rem;
            padding-bottom: 2rem;
        }}

        h1, h2, h3, h4, h5, h6,
        p, label, span, div,
        [data-testid="stMarkdownContainer"],
        [data-testid="stCaptionContainer"] {{
            color: var(--text);
        }}

        [data-testid="stSidebar"] * {{
            color: var(--text) !important;
        }}

        [data-testid="stCaptionContainer"] {{
            font-size: 1.12rem !important;
            line-height: 1.5 !important;
        }}

        [data-testid="stCaptionContainer"] p {{
            font-size: 1.12rem !important;
            line-height: 1.5 !important;
        }}

        [data-testid="stFileUploader"],
        [data-testid="stDataFrame"],
        [data-testid="stMetric"],
        [data-testid="stAlert"],
        [data-baseweb="select"] > div,
        .stTextInput > div > div > input,
        .stTextArea textarea {{
            background: rgba(255, 250, 250, 0.96) !important;
            border: 1px solid var(--border) !important;
            border-radius: 8px !important;
            box-shadow: 0 8px 20px rgba(20, 20, 20, 0.05);
            color: var(--text) !important;
        }}

        .stButton > button,
        [data-testid="baseButton-secondary"] {{
            background: var(--primary) !important;
            color: white !important;
            border: none !important;
            border-radius: 8px !important;
            font-weight: 700 !important;
            box-shadow: 0 8px 18px rgba(243, 91, 4, 0.18);
        }}

        .stButton > button *,
        [data-testid="baseButton-secondary"] * {{
            color: white !important;
        }}

        .stButton > button:hover {{
            background: var(--primary-hover) !important;
            color: white !important;
        }}

        [data-testid="stTabs"] button[role="tab"] {{
            border-radius: 8px;
            border: 1px solid var(--border);
            background: rgba(255, 250, 250, 0.9);
            color: var(--muted) !important;
        }}

        [data-testid="stTabs"] button[aria-selected="true"] {{
            background: rgba(243, 91, 4, 0.10);
            border-color: var(--primary-hover);
            color: var(--primary) !important;
        }}

        [data-testid="stTabs"] [data-baseweb="tab-highlight"] {{
            background-color: var(--primary) !important;
        }}

        [role="radiogroup"] {{
            gap: 0.5rem;
            flex-wrap: wrap;
            margin-bottom: 1rem;
        }}

        [role="radiogroup"] label {{
            background: rgba(255, 250, 250, 0.9);
            border: 1px solid var(--border);
            border-radius: 8px;
            padding: 0.35rem 0.7rem;
        }}

        [role="radiogroup"] label:has(input:checked) {{
            background: rgba(243, 91, 4, 0.10);
            border-color: var(--primary-hover);
        }}

        [data-baseweb="tag"] {{
            background-color: rgba(243, 91, 4, 0.10) !important;
            border: 1px solid rgba(243, 91, 4, 0.24) !important;
            border-radius: 7px !important;
            color: var(--text) !important;
        }}

        [data-baseweb="tag"] span,
        [data-baseweb="tag"] svg {{
            color: var(--text) !important;
            fill: var(--text) !important;
        }}

        [data-baseweb="select"] svg {{
            color: var(--muted) !important;
            fill: var(--muted) !important;
        }}

        .stDivider {{
            border-color: rgba(20, 20, 20, 0.12) !important;
        }}

        .kpi-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 1rem;
            margin: 0.35rem 0 0.25rem;
        }}

        .kpi-card {{
            background: var(--surface);
            padding: 24px 20px;
            border-radius: 8px;
            text-align: center;
            color: var(--text);
            min-height: 152px;
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            border: 1px solid var(--border);
            box-shadow: 0 10px 24px rgba(20, 20, 20, 0.06);
            position: relative;
            overflow: hidden;
        }}

        .kpi-card::before {{
            content: "";
            position: absolute;
            inset: 0 auto auto 0;
            width: 100%;
            height: 4px;
            background: var(--primary);
        }}

        .kpi-title {{
            font-size: 17px;
            font-weight: 900 !important;
            line-height: 1.25;
            margin-bottom: 12px;
            color: var(--text);
            text-transform: uppercase;
            letter-spacing: 0;
            max-width: 100%;
            overflow-wrap: anywhere;
        }}

        .kpi-value {{
            font-size: 46px;
            font-weight: 900 !important;
            color: var(--primary);
            line-height: 1.05;
            font-variant-numeric: tabular-nums;
        }}

        .executive-note {{
            background: rgba(255, 250, 250, 0.96);
            border: 1px solid var(--border);
            border-radius: 8px;
            color: var(--text);
            padding: 18px 18px 16px;
            margin: 0.1rem 0 0.25rem;
            line-height: 1.55;
            font-size: 1.12rem;
            box-shadow: 0 6px 16px rgba(20, 20, 20, 0.04);
        }}

        .executive-note-title {{
            color: var(--primary);
            font-weight: 800;
            margin-bottom: 0.75rem;
            font-size: 1.16rem;
        }}

        .executive-note-line {{
            color: var(--muted);
            margin: 0.35rem 0;
            font-size: 1.08rem;
        }}

        .executive-note-line strong {{
            color: var(--text);
        }}

        .executive-note-detail {{
            color: var(--muted);
            margin-top: 0.55rem;
            font-size: 1.06rem;
        }}

        .executive-note-detail strong {{
            color: var(--text);
        }}

        .executive-note-conclusion {{
            border-top: 1px solid var(--border);
            color: var(--muted);
            margin-top: 0.9rem;
            padding-top: 0.8rem;
            font-size: 1.04rem;
        }}

        .executive-table-card {{
            background: rgba(255, 250, 250, 0.96);
            border: 1px solid var(--border);
            border-radius: 8px;
            box-shadow: 0 6px 16px rgba(20, 20, 20, 0.04);
            margin: 0.1rem 0 0.75rem;
            overflow: hidden;
        }}

        .executive-table-title {{
            color: var(--primary);
            font-size: 1.12rem;
            font-weight: 900;
            padding: 14px 16px 8px;
        }}

        .executive-table {{
            border-collapse: collapse;
            table-layout: fixed;
            width: 100%;
        }}

        .executive-table th,
        .executive-table td {{
            border-top: 1px solid var(--border);
            color: var(--text);
            font-size: 0.92rem;
            line-height: 1.35;
            padding: 10px 12px;
            text-align: left;
            vertical-align: top;
            white-space: normal;
            word-break: normal;
            overflow-wrap: anywhere;
        }}

        .executive-table th {{
            background: #fff7f2;
            color: var(--muted);
            font-size: 0.78rem;
            font-weight: 900;
            text-transform: uppercase;
        }}

        .executive-table .number-cell {{
            color: var(--primary);
            font-weight: 900;
            text-align: right;
            white-space: nowrap;
        }}

        .ans-panel {{
            background: var(--surface);
            border: 1px solid var(--border);
            border-radius: 8px;
            padding: 22px;
            margin: 0.8rem 0 1rem;
            box-shadow: 0 10px 24px rgba(20, 20, 20, 0.05);
        }}

        .ans-panel-title {{
            color: var(--text);
            font-size: 1.28rem;
            font-weight: 900;
            line-height: 1.25;
            margin-bottom: 0.3rem;
        }}

        .ans-panel-subtitle {{
            color: var(--muted);
            font-size: 1.02rem;
            margin-bottom: 1rem;
        }}

        .ans-table {{
            width: 100%;
            border-collapse: collapse;
            table-layout: fixed;
            font-size: 1.02rem;
        }}

        .ans-table th {{
            background: rgba(20, 20, 20, 0.04);
            color: var(--text);
            font-weight: 900;
            padding: 12px 10px;
            text-align: center;
            border-bottom: 1px solid var(--border);
        }}

        .ans-table td {{
            color: var(--text);
            padding: 14px 10px;
            text-align: center;
            border-bottom: 1px solid var(--border);
            font-variant-numeric: tabular-nums;
        }}

        .ans-table tr:last-child td {{
            border-bottom: none;
        }}

        .ans-period {{
            font-weight: 900;
            text-align: left !important;
        }}

        .ans-pill {{
            display: inline-flex;
            align-items: center;
            justify-content: center;
            min-width: 76px;
            padding: 0.28rem 0.55rem;
            border-radius: 999px;
            background: rgba(243, 91, 4, 0.10);
            color: var(--primary);
            font-weight: 900;
        }}

        .ans-card-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(190px, 1fr));
            gap: 0.85rem;
            margin: 0.7rem 0 0.3rem;
        }}

        .ans-card {{
            border: 1px solid var(--border);
            border-radius: 8px;
            padding: 16px 14px;
            background: rgba(255, 250, 250, 0.92);
        }}

        .ans-card-label {{
            color: var(--muted);
            font-size: 0.95rem;
            font-weight: 800;
            margin-bottom: 0.4rem;
        }}

        .ans-card-value {{
            color: var(--primary);
            font-size: 2rem;
            font-weight: 900;
            line-height: 1.05;
            font-variant-numeric: tabular-nums;
        }}

        .kpi-ranking {{
            background: var(--surface);
            border: 1px solid var(--border);
            border-radius: 8px;
            padding: 18px 18px 16px;
            box-shadow: 0 8px 18px rgba(20, 20, 20, 0.04);
        }}

        .kpi-ranking-title {{
            color: var(--primary);
            font-size: 1.25rem;
            font-weight: 900;
            line-height: 1.25;
            margin-bottom: 1rem;
        }}

        .kpi-ranking-list {{
            display: flex;
            flex-direction: column;
            gap: 0.72rem;
        }}

        .kpi-ranking-row {{
            display: grid;
            grid-template-columns: minmax(0, 1.65fr) minmax(118px, 1fr) 3.2rem;
            align-items: center;
            gap: 0.85rem;
        }}

        .kpi-ranking-label {{
            color: var(--text);
            font-size: 1.02rem;
            font-weight: 800;
            line-height: 1.25;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }}

        .kpi-ranking-track {{
            height: 18px;
            background: #f1f3f5;
            border-radius: 999px;
            overflow: hidden;
        }}

        .kpi-ranking-bar {{
            height: 100%;
            background: var(--mustard);
            border-radius: inherit;
        }}

        .kpi-ranking-value {{
            color: var(--text);
            font-size: 1.12rem;
            font-weight: 900;
            line-height: 1;
            text-align: right;
            font-variant-numeric: tabular-nums;
        }}

        .kpi-ranking-empty {{
            color: var(--muted);
            font-size: 1rem;
            font-weight: 700;
        }}

        .slide-frame {{
            width: min(100%, 1600px);
            aspect-ratio: 16 / 9;
            background: #ffffff;
            border: 1px solid var(--border);
            border-radius: 8px;
            box-shadow: 0 12px 28px rgba(20, 20, 20, 0.08);
            padding: 28px 32px;
            display: flex;
            flex-direction: column;
            gap: 16px;
            overflow: hidden;
        }}

        .slide-period {{
            color: var(--muted);
            font-size: 1.04rem;
            font-weight: 700;
            line-height: 1.15;
        }}

        .slide-title {{
            color: var(--text);
            font-size: 1.85rem;
            font-weight: 900;
            line-height: 1.12;
        }}

        .slide-kpi-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(165px, 1fr));
            gap: 12px;
        }}

        .slide-kpi-card {{
            background: #ffffff;
            border: 1px solid var(--border);
            border-top: 4px solid var(--primary);
            border-radius: 8px;
            min-height: 104px;
            padding: 16px 12px;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            text-align: center;
        }}

        .slide-kpi-title {{
            color: var(--text);
            font-size: 0.92rem;
            font-weight: 900;
            line-height: 1.18;
            text-transform: uppercase;
        }}

        .slide-kpi-value {{
            color: var(--primary);
            font-size: 2.35rem;
            font-weight: 900;
            line-height: 1;
            margin-top: 0.42rem;
            font-variant-numeric: tabular-nums;
        }}

        .slide-caption {{
            color: var(--muted);
            font-size: 1.02rem;
            font-weight: 700;
            line-height: 1.25;
        }}

        .slide-body {{
            display: grid;
            grid-template-columns: minmax(0, 2fr) minmax(300px, 0.9fr);
            gap: 14px;
            min-height: 0;
            flex: 1;
        }}

        .slide-panel {{
            background: #ffffff;
            border: 1px solid var(--border);
            border-radius: 8px;
            padding: 16px 18px;
            min-height: 0;
            overflow: hidden;
        }}

        .slide-panel-group {{
            display: grid;
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: 14px;
            min-height: 0;
        }}

        .slide-panel-title {{
            color: var(--primary);
            font-size: 1.18rem;
            font-weight: 800;
            line-height: 1.2;
            margin-bottom: 0.9rem;
        }}

        .slide-ranking-list {{
            display: flex;
            flex-direction: column;
            gap: 0.64rem;
        }}

        .slide-ranking-row {{
            display: grid;
            grid-template-columns: minmax(0, 1.6fr) minmax(120px, 1fr) 3.1rem;
            gap: 0.8rem;
            align-items: center;
        }}

        .slide-ranking-label {{
            color: var(--text);
            font-size: 0.96rem;
            font-weight: 650;
            line-height: 1.18;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }}

        .slide-ranking-track {{
            height: 17px;
            background: #f1f3f5;
            border-radius: 999px;
            overflow: hidden;
        }}

        .slide-ranking-bar {{
            height: 100%;
            background: var(--mustard);
            border-radius: inherit;
        }}

        .slide-ranking-value {{
            color: var(--text);
            font-size: 1.05rem;
            font-weight: 800;
            text-align: right;
            font-variant-numeric: tabular-nums;
        }}

        .slide-note {{
            display: flex;
            flex-direction: column;
            gap: 0.64rem;
            color: var(--text);
            font-size: 1rem;
            font-weight: 500;
            line-height: 1.38;
        }}

        .slide-note strong {{
            color: var(--text);
            font-weight: 700;
        }}

        .slide-note-muted {{
            color: var(--muted);
            font-weight: 500;
        }}

        @media (max-width: 900px) {{
            .block-container {{
                padding-left: 1rem;
                padding-right: 1rem;
            }}

            [data-testid="stHorizontalBlock"] {{
                flex-direction: column;
            }}

            [data-testid="stHorizontalBlock"] > div {{
                width: 100% !important;
                min-width: 100% !important;
            }}

            .kpi-grid {{
                grid-template-columns: repeat(2, minmax(0, 1fr));
                gap: 0.75rem;
            }}

            .kpi-card {{
                min-height: 138px;
                padding: 22px 18px;
            }}

            .kpi-title {{
                font-size: 16px;
            }}

            .kpi-value {{
                font-size: 42px;
            }}

            .kpi-ranking-row {{
                grid-template-columns: minmax(0, 1.35fr) minmax(90px, 1fr) 2.8rem;
            }}

            .slide-frame {{
                aspect-ratio: auto;
                min-height: 760px;
                padding: 22px;
            }}

            .slide-body {{
                grid-template-columns: 1fr;
            }}

            .slide-panel-group {{
                grid-template-columns: 1fr;
            }}

            [data-testid="stTabs"] div[role="tablist"] {{
                gap: 0.35rem;
                overflow-x: auto;
                padding-bottom: 0.25rem;
            }}

            [data-testid="stTabs"] button[role="tab"] {{
                min-width: max-content;
                padding: 0.45rem 0.7rem;
            }}
        }}

        @media (max-width: 520px) {{
            .block-container {{
                padding-top: 0.75rem;
                padding-left: 0.75rem;
                padding-right: 0.75rem;
            }}

            h2, h3 {{
                font-size: 1.25rem !important;
                line-height: 1.25 !important;
            }}

            .kpi-grid {{
                grid-template-columns: 1fr;
            }}

            .kpi-card {{
                min-height: 120px;
                padding: 18px 14px;
            }}

            .kpi-title {{
                font-size: 15px;
                margin-bottom: 8px;
            }}

            .kpi-value {{
                font-size: 36px;
            }}

            .kpi-ranking {{
                padding: 16px 14px;
            }}

            .kpi-ranking-row {{
                grid-template-columns: 1fr 2.6rem;
                gap: 0.35rem 0.6rem;
            }}

            .kpi-ranking-label {{
                grid-column: 1 / -1;
                white-space: normal;
            }}

            .kpi-ranking-track {{
                height: 16px;
            }}

            [data-baseweb="tag"] {{
                max-width: 100%;
            }}

            [data-testid="stDataFrame"] {{
                font-size: 12px !important;
            }}
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def aplicar_estilo_figura(fig, titulo=None):
    fig.update_layout(
        title=titulo,
        paper_bgcolor="rgba(255,255,255,0)",
        plot_bgcolor="#ffffff",
        font={"color": UI_PALETTE["text"], "size": 17, "family": "Arial, sans-serif"},
        title_font={"color": UI_PALETTE[TEXT_PRIMARY], "size": 23},
        margin={"l": 12, "r": 12, "t": 52, "b": 12},
        legend={"bgcolor": "#ffffff", "font": {"size": 16, "color": UI_PALETTE["text"]}},
    )
    fig.update_xaxes(
        showgrid=True,
        gridcolor="rgba(20, 20, 20, 0.10)",
        zeroline=False,
        tickfont={"size": 16, "color": UI_PALETTE["text"]},
        title_font={"size": 17, "color": UI_PALETTE["text"]},
        automargin=True,
    )
    fig.update_yaxes(
        showgrid=False,
        zeroline=False,
        tickfont={"size": 16, "color": UI_PALETTE["text"]},
        title_font={"size": 17, "color": UI_PALETTE["text"]},
        automargin=True,
    )
    fig.update_traces(textfont={"size": 16, "color": UI_PALETTE["text"]}, selector={"type": "bar"})
    return fig


def limpiar_ejes_kpi(fig):
    fig.update_xaxes(title_text="", showticklabels=False)
    fig.update_yaxes(title_text="")
    return fig


def estilos_login():
    st.markdown(
        """
        <style>
        header {visibility: hidden;}
        footer {visibility: hidden;}

        /* Fondo general */
        .stApp {
            background: var(--bg) !important;
        }

        .block-container {
            max-width: 1180px;
            padding-top: 4rem;
        }

        .login-spacer {
            height: 5vh;
        }

        /* Card principal */
        .login-card {
            background: rgba(255, 250, 250, 0.98);
            padding: 36px 32px;
            border-radius: 8px;
            box-shadow: 0 18px 42px rgba(20, 20, 20, 0.12);
            text-align: center;
            max-width: 420px;
            margin: auto;
            width: 100%;
            border: 1px solid var(--border);
        }

        /* TÃ­tulo */
        .login-title {
            font-size: 28px;
            font-weight: 800;
            color: var(--primary);
            margin-bottom: 8px;
            text-align: left;
        }

        /* SubtÃ­tulo */
        .login-subtitle {
            font-size: 14px;
            color: var(--muted);
            margin-bottom: 24px;
            text-align: left;
        }

        /* Inputs */
        input {
            border-radius: 8px !important;
            border: 1px solid var(--border) !important;
            padding: 10px !important;
            transition: all 0.2s ease;
        }

        input:focus {
            border: 1px solid var(--primary) !important;
            box-shadow: 0 0 0 3px rgba(243, 91, 4, 0.14);
            outline: none;
        }

        /* BotÃ³n */
        div.stButton > button {
            width: 100%;
            border-radius: 8px;
            background: var(--primary);
            color: white;
            font-weight: 700;
            border: none;
            padding: 0.7rem 1rem;
            transition: all 0.25s ease;
            box-shadow: 0 8px 18px rgba(243, 91, 4, 0.20);
        }

        div.stButton > button:hover {
            background: var(--primary-hover);
            box-shadow: 0 12px 24px rgba(241, 135, 1, 0.24);
        }

        /* Placeholder */
        ::placeholder {
            color: var(--muted);
            font-size: 13px;
        }

        /* Responsive */
        @media (max-width: 768px) {
            .login-card {
                padding: 24px !important;
                border-radius: 8px !important;
            }

            .login-title {
                font-size: 22px !important;
            }

            .login-subtitle {
                font-size: 13px !important;
            }
        }
        </style>
        """,
        unsafe_allow_html=True
    )


def validar_email(correo):
    return re.match(r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$", str(correo).strip())


def login():
    if "user" in st.session_state:
        return True

    estilos_login()
    st.markdown('<div class="login-spacer"></div>', unsafe_allow_html=True)

    _, col_login, _ = st.columns([1.15, 0.7, 1.15])
    with col_login:
        st.markdown('<div class="login-title">Control de casos e incidentes</div>', unsafe_allow_html=True)
        st.markdown('<div class="login-subtitle">Ingresa con tu correo y contrasena</div>', unsafe_allow_html=True)
        correo = st.text_input("Correo corporativo", key="correo_login")
        password = st.text_input("Contrasena", type=TEXT_PASSWORD, key="password_login")

        if st.button("Ingresar", key="btn_login"):
            with st.spinner("Validando acceso..."):
                if not validar_email(correo):
                    st.error("Escribe un correo valido")
                    return False

                usuario = autenticar_usuario(correo, password)
                if not usuario:
                    st.error("Correo o contrasena incorrectos, o usuario inactivo")
                    return False

                st.session_state.user = usuario["email"]
                st.session_state.role = usuario["role"]
                st.rerun()

    return False


def configurar_primer_admin():
    estilos_login()
    st.markdown('<div class="login-spacer"></div>', unsafe_allow_html=True)

    _, col_setup, _ = st.columns([1.1, 0.8, 1.1])
    with col_setup:
        st.markdown('<div class="login-title">Configurar acceso</div>', unsafe_allow_html=True)
        st.markdown(
            '<div class="login-subtitle">Crea el primer usuario administrador para iniciar la app.</div>',
            unsafe_allow_html=True,
        )
        correo = st.text_input("Correo administrador", key="setup_admin_email")
        password = st.text_input("Contrasena", type=TEXT_PASSWORD, key="setup_admin_password")
        confirmar = st.text_input("Confirmar contrasena", type=TEXT_PASSWORD, key="setup_admin_password_confirm")

        if st.button("Crear administrador", key="btn_setup_admin"):
            if not validar_email(correo):
                st.error("Escribe un correo valido.")
                return
            if len(password or "") < 8:
                st.error("La contrasena debe tener minimo 8 caracteres.")
                return
            if password != confirmar:
                st.error("Las contrasenas no coinciden.")
                return

            guardar_usuario(correo, password, role=TEXT_ADMIN, active=True)
            st.session_state.user = normalizar_email(correo)
            st.session_state.role = TEXT_ADMIN
            st.success("Administrador creado.")
            st.rerun()


def tarjeta(titulo, valor):
    titulo = html.escape(str(titulo))
    valor = html.escape(str(valor))
    return (
        '<div class="kpi-card">'
        f'<div class="kpi-title">{titulo}</div>'
        f'<div class="kpi-value">{valor}</div>'
        "</div>"
    )


def render_tarjetas(items):
    contenido = "".join(tarjeta(titulo, valor) for titulo, valor in items)
    st.markdown(f'<div class="kpi-grid">{contenido}</div>', unsafe_allow_html=True)


def ejecutar_con_carga(nombre, funcion):
    with st.spinner(f"Cargando {nombre.lower()}..."):
        funcion()


def valor_limpio(valor):
    if valor is None or pd.isna(valor):
        return ""
    return str(valor).strip()


def porcentaje(valor, total):
    return round((valor / total) * 100, 2) if total else 0


RANGO_RESOLUCION_ORDEN = [
    "<=4h",
    "4-8h",
    "8-24h",
    "1-2 dias",
    "2-4 dias",
    "4-8 dias",
    ">8 dias",
    SIN_DURACION,
]


def formato_horas_dias(valor):
    try:
        horas = float(valor)
    except (TypeError, ValueError):
        return "Sin SLA"
    if pd.isna(horas):
        return "Sin SLA"
    dias = horas / 24
    horas_texto = f"{horas:g}h"
    if horas < 24:
        return f"{horas_texto} / {dias:.2f}{TEXT_DIAS}"
    dias_texto = f"{int(dias)}{TEXT_DIAS}" if float(dias).is_integer() else f"{dias:.2f}{TEXT_DIAS}"
    return f"{horas_texto} / {dias_texto}"


def clasificar_rango_resolucion(horas):
    try:
        horas = float(horas)
    except (TypeError, ValueError):
        return SIN_DURACION
    if pd.isna(horas):
        return SIN_DURACION
    if horas <= 4:
        return "<=4h"
    if horas <= 8:
        return "4-8h"
    if horas <= 24:
        return "8-24h"
    if horas <= 48:
        return "1-2 dias"
    if horas <= 96:
        return "2-4 dias"
    if horas <= 192:
        return "4-8 dias"
    return ">8 dias"


def mascara_cerrados(df):
    if df.empty:
        return pd.Series(False, index=df.index)
    cerrado_por_estado = pd.Series(False, index=df.index)
    if TEXT_ESTADO in df.columns:
        estado = df[TEXT_ESTADO].apply(normalizar_texto)
        cerrado_por_estado = estado.str.contains(
            r"(?<![a-z0-9])(?:cerrado|closed|resuelto|resolved|solucionado|finalizado|completado)(?![a-z0-9])",
            regex=True,
            na=False,
        )
    if TEXT_CERRADO in df.columns:
        cerrado_por_fecha = df[TEXT_CERRADO].apply(valor_limpio).ne("")
        return cerrado_por_estado | cerrado_por_fecha
    return cerrado_por_estado


def mascara_prioridad_alta(df):
    if df.empty or TEXT_PRIORIDAD not in df.columns:
        return pd.Series(False, index=df.index)
    return df[TEXT_PRIORIDAD].fillna("").str.contains(
        r"^(?:1|2\s*-\s*Alta|alta|critica|critico)",
        case=False,
        regex=True,
    )


def resumen_carga_agentes(df, columna_agente):
    columnas = [
        TEXT_RESPONSABLE,
        TEXT_CERRADOS,
        TEXT_ABIERTOS,
        "Total asignado",
        "% periodo",
        "% carga abierta",
    ]
    if df.empty or columna_agente not in df.columns:
        return pd.DataFrame(columns=columnas)

    trabajo = df.copy()
    trabajo[TEXT_RESPONSABLE] = trabajo[columna_agente].apply(valor_limpio).replace("", "Sin agente")
    trabajo[TEXT_CERRADO_2] = mascara_cerrados(trabajo)
    trabajo[TEXT_ABIERTO] = ~trabajo[TEXT_CERRADO_2]

    total_periodo = len(trabajo)
    total_abiertos = int(trabajo[TEXT_ABIERTO].sum())
    resumen = (
        trabajo.groupby(TEXT_RESPONSABLE, dropna=False)
        .agg(
            Cerrados=(TEXT_CERRADO_2, "sum"),
            Abiertos=(TEXT_ABIERTO, "sum"),
            Total_asignado=(TEXT_NUMERO, TEXT_COUNT),
        )
        .reset_index()
    )
    resumen[TEXT_CERRADOS] = resumen["Cerrados"].astype(int)
    resumen[TEXT_ABIERTOS] = resumen["Abiertos"].astype(int)
    resumen["Total asignado"] = resumen["Total_asignado"].astype(int)
    resumen["% periodo"] = resumen["Total asignado"].apply(lambda valor: porcentaje(valor, total_periodo))
    resumen["% carga abierta"] = resumen[TEXT_ABIERTOS].apply(lambda valor: porcentaje(valor, total_abiertos))
    return resumen[columnas].sort_values(
        by=[TEXT_ABIERTOS, "Total asignado", TEXT_RESPONSABLE],
        ascending=[False, False, True],
    )


def render_carga_agentes(df, columna_agente, titulo, etiqueta_registro):
    st.subheader(titulo)
    resumen = resumen_carga_agentes(df, columna_agente)
    if resumen.empty:
        st.info(f"No hay responsables asignados para calcular carga de {etiqueta_registro.lower()}.")
        return

    total_abiertos = int(resumen[TEXT_ABIERTOS].sum())
    st.caption(
        f"El % de carga abierta compara los pendientes de cada agente contra el 100% de "
        f"{etiqueta_registro.lower()} abiertos del periodo seleccionado."
    )
    grafico = resumen.sort_values(by="% carga abierta" if total_abiertos else "Total asignado", ascending=True)
    eje_x = "% carga abierta" if total_abiertos else "Total asignado"
    fig = px.bar(
        grafico,
        x=eje_x,
        y=TEXT_RESPONSABLE,
        orientation="h",
        text=eje_x,
        color_discrete_sequence=[UI_PALETTE[TEXT_PRIMARY]],
    )
    fig.update_traces(textposition=TEXT_OUTSIDE)
    st.plotly_chart(aplicar_estilo_figura(fig, titulo), use_container_width=True)
    st.dataframe(resumen, use_container_width=True, hide_index=True)


def preparar_seguimiento_operativo_incidentes(df, horas_proximas=24):
    trabajo = agregar_campos_sla_incidentes(df)
    if trabajo.empty:
        return trabajo

    ahora = pd.Timestamp.now()
    trabajo[TEXT_CERRADO_2] = mascara_cerrados(trabajo)
    trabajo[TEXT_ABIERTO] = ~trabajo[TEXT_CERRADO_2]
    trabajo[TEXT_CREADO_DT_2] = pd.to_datetime(trabajo[TEXT_CREADO].apply(normalizar_fecha), errors=TEXT_COERCE)
    trabajo["_vencimiento_dt"] = pd.to_datetime(
        trabajo[TEXT_FECHA_VENCIMIENTO_SLA].apply(normalizar_fecha),
        errors=TEXT_COERCE,
    )
    trabajo[TEXT_HORAS_ABIERTO] = ((ahora - trabajo[TEXT_CREADO_DT_2]).dt.total_seconds() / 3600).round(2)
    trabajo["_horas_para_vencer_sistema"] = ((trabajo["_vencimiento_dt"] - ahora).dt.total_seconds() / 3600).round(2)
    trabajo["_horas_para_vencer_matriz"] = (trabajo[TEXT_SLA_OBJETIVO_HORAS] - trabajo[TEXT_HORAS_ABIERTO]).round(2)
    trabajo[TEXT_HORAS_PARA_VENCER] = trabajo["_horas_para_vencer_matriz"].where(
        trabajo[TEXT_SLA_OBJETIVO_HORAS].notna(),
        trabajo["_horas_para_vencer_sistema"],
    )
    trabajo[TEXT_VENCIDO] = trabajo[TEXT_ABIERTO] & trabajo[TEXT_HORAS_PARA_VENCER].notna() & (trabajo[TEXT_HORAS_PARA_VENCER] < 0)
    trabajo[TEXT_PROXIMO_VENCER] = (
        trabajo[TEXT_ABIERTO]
        & trabajo[TEXT_HORAS_PARA_VENCER].notna()
        & trabajo[TEXT_HORAS_PARA_VENCER].between(0, horas_proximas, inclusive="both")
    )
    trabajo[TEXT_PRIORIDAD_ALTA] = mascara_prioridad_alta(trabajo)
    trabajo[TEXT_ALERTA] = trabajo[TEXT_ES_ALERTA_AUTO].fillna("No").eq("Si")
    trabajo[TEXT_CLIENTE_EXTERNO] = trabajo[TEXT_TIPIFICACION_AUTO].fillna("").isin(
        [TIPIFICACION_INCIDENTE_CLIENTE_EXTERNO, TIPIFICACION_CASO_CLIENTE_EXTERNO]
    )
    trabajo["_requiere_seguimiento"] = (
        trabajo[TEXT_ABIERTO]
        & (
            trabajo[TEXT_VENCIDO]
            | trabajo[TEXT_PROXIMO_VENCER]
            | trabajo[TEXT_PRIORIDAD_ALTA]
            | trabajo[TEXT_ALERTA]
            | trabajo[TEXT_CLIENTE_EXTERNO]
        )
    )
    return trabajo


def render_seguimiento_operativo_incidentes(df):
    seguimiento = preparar_seguimiento_operativo_incidentes(df)
    if seguimiento.empty:
        return

    abiertos = seguimiento[seguimiento[TEXT_ABIERTO]].copy()
    vencidos = seguimiento[seguimiento[TEXT_VENCIDO]].copy()
    proximos = seguimiento[seguimiento[TEXT_PROXIMO_VENCER]].copy()
    prioridad_alta = seguimiento[seguimiento[TEXT_ABIERTO] & seguimiento[TEXT_PRIORIDAD_ALTA]].copy()
    alertas_abiertas = seguimiento[seguimiento[TEXT_ABIERTO] & seguimiento[TEXT_ALERTA]].copy()
    cliente_externo_abierto = seguimiento[seguimiento[TEXT_ABIERTO] & seguimiento[TEXT_CLIENTE_EXTERNO]].copy()

    st.subheader("Seguimiento operativo")
    st.caption(
        "Vista de control para incidentes abiertos que requieren accion: vencidos por SLA, proximos a vencer "
        "en 24 horas, prioridad alta, alertas abiertas o afectacion de cliente externo."
    )
    render_tarjetas(
        [
            (TEXT_ABIERTOS, len(abiertos)),
            (TEXT_VENCIDOS, len(vencidos)),
            ("Proximos 24h", len(proximos)),
            (PRIORIDAD_ALTA, len(prioridad_alta)),
            ("Alertas abiertas", len(alertas_abiertas)),
        ]
    )
    st.caption(
        "Los vencidos y proximos a vencer se calculan primero con la matriz SLA por prioridad; "
        "si no hay objetivo configurado, se usa la fecha de vencimiento del sistema."
    )

    columnas = [
        TEXT_NUMERO,
        TEXT_ESTADO,
        TEXT_PRIORIDAD,
        TEXT_PRIORIDAD_NORMALIZADA,
        TEXT_SLA_OBJETIVO_HORAS,
        TEXT_ESTADO_SLA,
        "grupo_asignacion",
        TEXT_ASIGNADO_A,
        TEXT_EMPRESA,
        TEXT_SERVICIO_NEGOCIO,
        TEXT_CREADO,
        TEXT_FECHA_VENCIMIENTO_SLA,
        TEXT_HORAS_ABIERTO,
        TEXT_HORAS_PARA_VENCER,
        TEXT_TIPIFICACION_AUTO,
        TEXT_ES_ALERTA_AUTO,
        TEXT_CAUSA_RAIZ_AUTO,
    ]
    etiquetas = {
        TEXT_HORAS_ABIERTO: "horas_abierto",
        TEXT_HORAS_PARA_VENCER: "horas_para_vencer",
    }

    tab_vencidos, tab_proximos, tab_prioridad, tab_alertas, tab_cliente = st.tabs(
        [TEXT_VENCIDOS, "Proximos 24h", PRIORIDAD_ALTA, "Alertas", "Cliente externo"]
    )

    tablas = [
        (tab_vencidos, vencidos.sort_values(by=TEXT_HORAS_PARA_VENCER)),
        (tab_proximos, proximos.sort_values(by=TEXT_HORAS_PARA_VENCER)),
        (tab_prioridad, prioridad_alta.sort_values(by=TEXT_HORAS_ABIERTO, ascending=False)),
        (tab_alertas, alertas_abiertas.sort_values(by=TEXT_HORAS_ABIERTO, ascending=False)),
        (tab_cliente, cliente_externo_abierto.sort_values(by=TEXT_HORAS_ABIERTO, ascending=False)),
    ]

    for tab, tabla in tablas:
        with tab:
            if tabla.empty:
                st.info("No hay incidentes para este criterio en el periodo seleccionado.")
            else:
                visible = tabla[[col for col in columnas if col in tabla.columns]].rename(columns=etiquetas)
                st.dataframe(visible, use_container_width=True, hide_index=True)


def preparar_seguimiento_casos(df, horas_proximas=12):
    trabajo = df.copy()
    if trabajo.empty:
        return trabajo

    ahora = pd.Timestamp.now()
    trabajo[TEXT_CERRADO_2] = mascara_cerrados(trabajo)
    trabajo[TEXT_ABIERTO] = ~trabajo[TEXT_CERRADO_2]
    trabajo[TEXT_CREADO_DT_2] = pd.to_datetime(trabajo[TEXT_CREADO].apply(normalizar_fecha), errors=TEXT_COERCE)
    trabajo[TEXT_HORAS_ABIERTO] = trabajo[TEXT_CREADO_DT_2].apply(lambda fecha: horas_habiles_entre(fecha, ahora))
    trabajo[TEXT_HORAS_PARA_VENCER] = (SLA_CASOS_HORAS - trabajo[TEXT_HORAS_ABIERTO]).round(2)
    trabajo[TEXT_VENCIDO] = trabajo[TEXT_ABIERTO] & trabajo[TEXT_CREADO_DT_2].notna() & (trabajo[TEXT_HORAS_PARA_VENCER] < 0)
    trabajo[TEXT_PROXIMO_VENCER] = (
        trabajo[TEXT_ABIERTO]
        & trabajo[TEXT_CREADO_DT_2].notna()
        & trabajo[TEXT_HORAS_PARA_VENCER].between(0, horas_proximas, inclusive="both")
    )
    return trabajo


def horas_habiles_entre(inicio, fin):
    if pd.isna(inicio) or pd.isna(fin):
        return None

    actual = pd.Timestamp(inicio)
    fin = pd.Timestamp(fin)
    if actual >= fin:
        return 0

    total = pd.Timedelta(0)
    while actual < fin:
        if actual.weekday() >= 5:
            actual = (actual + pd.Timedelta(days=1)).replace(hour=8, minute=0, second=0, microsecond=0)
            continue

        inicio_dia = actual.replace(hour=8, minute=0, second=0, microsecond=0)
        fin_dia = actual.replace(hour=17, minute=0, second=0, microsecond=0)

        if actual < inicio_dia:
            actual = inicio_dia
        if actual >= fin_dia:
            actual = (actual + pd.Timedelta(days=1)).replace(hour=8, minute=0, second=0, microsecond=0)
            continue

        limite = min(fin, fin_dia)
        total += limite - actual
        actual = limite

    return round(total.total_seconds() / 3600, 2)


def render_seguimiento_casos(df):
    seguimiento = preparar_seguimiento_casos(df)
    if seguimiento.empty:
        return

    abiertos = seguimiento[seguimiento[TEXT_ABIERTO]].copy()
    vencidos = seguimiento[seguimiento[TEXT_VENCIDO]].copy()
    proximos = seguimiento[seguimiento[TEXT_PROXIMO_VENCER]].copy()

    st.divider()
    st.subheader("Control de vencimiento")
    st.caption(
        f"Calculado solo sobre casos abiertos con el mismo criterio de horas habiles del SLA de {SLA_CASOS_HORAS} horas. "
        "El SLA superior resume casos cerrados; esta tabla muestra pendientes abiertos."
    )
    render_tarjetas(
        [
            (TEXT_ABIERTOS, len(abiertos)),
            (TEXT_VENCIDOS, len(vencidos)),
            ("Proximos 12h", len(proximos)),
        ]
    )

    columnas = [
        TEXT_NUMERO,
        TEXT_ESTADO,
        TEXT_PRIORIDAD,
        TEXT_CUENTA,
        "contacto",
        TEXT_ASIGNADO,
        TEXT_CREADO,
        TEXT_HORAS_ABIERTO,
        TEXT_HORAS_PARA_VENCER,
        TEXT_TIPIFICACION_2,
        TEXT_DESCRIPCION_2,
    ]
    etiquetas = {
        TEXT_HORAS_ABIERTO: "horas_habiles_abierto",
        TEXT_HORAS_PARA_VENCER: "horas_habiles_para_vencer",
    }

    tab_vencidos, tab_proximos = st.tabs([TEXT_VENCIDOS, "Proximos 12h"])
    for tab, tabla in [
        (tab_vencidos, vencidos.sort_values(by=TEXT_HORAS_PARA_VENCER)),
        (tab_proximos, proximos.sort_values(by=TEXT_HORAS_PARA_VENCER)),
    ]:
        with tab:
            if tabla.empty:
                st.info("No hay casos para este criterio en el periodo seleccionado.")
            else:
                visible = tabla[[col for col in columnas if col in tabla.columns]].rename(columns=etiquetas)
                st.dataframe(visible, use_container_width=True, hide_index=True)


def aliases_clientes_ordenados():
    aliases = []
    for cliente, opciones in CLIENTES_CLAVE_ALIASES.items():
        for alias in opciones:
            alias_normalizado = normalizar_texto(alias)
            if alias_normalizado:
                aliases.append((cliente, alias_normalizado))
    return sorted(aliases, key=lambda item: len(item[1]), reverse=True)


CLIENTES_CLAVE_ALIAS_ORDENADOS = aliases_clientes_ordenados()


def texto_contiene_alias(texto_normalizado, alias_normalizado):
    patron = rf"(?<!\w){re.escape(alias_normalizado)}(?!\w)"
    return re.search(patron, texto_normalizado) is not None


def detectar_cliente_clave(texto):
    texto_normalizado = normalizar_texto(texto)
    if not texto_normalizado:
        return ""
    for cliente, alias in CLIENTES_CLAVE_ALIAS_ORDENADOS:
        if texto_contiene_alias(texto_normalizado, alias):
            return cliente
    return ""


def detectar_cliente_en_fila(row, campos):
    for campo in campos:
        cliente = detectar_cliente_clave(valor_limpio(row.get(campo)))
        if cliente:
            return cliente, campo
    return "", ""


def preparar_casos_clientes_clave(df):
    if df.empty:
        trabajo = df.copy()
        trabajo[TEXT_CLIENTE_CLAVE] = pd.Series(dtype=TEXT_OBJECT)
        trabajo[TEXT_FUENTE_CLIENTE] = pd.Series(dtype=TEXT_OBJECT)
        trabajo[TEXT_CREADO_DT] = pd.Series(dtype=PANDAS_DATETIME_DTYPE)
        trabajo[TEXT_CERRADO_DT] = pd.Series(dtype=PANDAS_DATETIME_DTYPE)
        trabajo[TEXT_TIEMPO_RESPUESTA_H] = pd.Series(dtype=TEXT_FLOAT)
        return trabajo

    trabajo = normalizar_tipificaciones_casos_df(df)
    detecciones = trabajo.apply(
        lambda row: detectar_cliente_en_fila(row, [TEXT_CUENTA]),
        axis=1,
        result_type="expand",
    )
    detecciones.columns = [TEXT_CLIENTE_CLAVE, TEXT_FUENTE_CLIENTE]
    trabajo[[TEXT_CLIENTE_CLAVE, TEXT_FUENTE_CLIENTE]] = detecciones
    trabajo = trabajo[trabajo[TEXT_CLIENTE_CLAVE] != ""].copy()
    trabajo[TEXT_CREADO_DT] = pd.to_datetime(trabajo[TEXT_CREADO].apply(normalizar_fecha), errors=TEXT_COERCE)
    trabajo[TEXT_CERRADO_DT] = pd.to_datetime(trabajo[TEXT_CERRADO].apply(normalizar_fecha), errors=TEXT_COERCE)
    trabajo[TEXT_TIEMPO_RESPUESTA_H] = pd.to_numeric(trabajo[TEXT_TIEMPO_RESPUESTA], errors=TEXT_COERCE)
    return trabajo


def preparar_incidentes_clientes_clave(df):
    if df.empty:
        trabajo = agregar_campos_sla_incidentes(df)
        trabajo[TEXT_CLIENTE_CLAVE] = pd.Series(dtype=TEXT_OBJECT)
        trabajo[TEXT_FUENTE_CLIENTE] = pd.Series(dtype=TEXT_OBJECT)
        trabajo[TEXT_CREADO_DT] = pd.Series(dtype=PANDAS_DATETIME_DTYPE)
        trabajo[TEXT_CERRADO_DT] = pd.Series(dtype=PANDAS_DATETIME_DTYPE)
        trabajo[TEXT_DURACION_HORAS_NUM] = pd.Series(dtype=TEXT_FLOAT)
        return trabajo

    trabajo = agregar_campos_sla_incidentes(df)
    campos = [
        TEXT_EMPRESA,
        TEXT_SOLICITANTE,
        TEXT_BREVE_DESCRIPCION,
        TEXT_DESCRIPCION_2,
        TEXT_OBSERVACIONES_TRABAJO,
        TEXT_OBSERVACIONES_ADICIONALES,
    ]
    detecciones = trabajo.apply(
        lambda row: detectar_cliente_en_fila(row, campos),
        axis=1,
        result_type="expand",
    )
    detecciones.columns = [TEXT_CLIENTE_CLAVE, TEXT_FUENTE_CLIENTE]
    trabajo[[TEXT_CLIENTE_CLAVE, TEXT_FUENTE_CLIENTE]] = detecciones
    trabajo = trabajo[trabajo[TEXT_CLIENTE_CLAVE] != ""].copy()
    trabajo[TEXT_CREADO_DT] = pd.to_datetime(trabajo[TEXT_CREADO].apply(normalizar_fecha), errors=TEXT_COERCE)
    trabajo[TEXT_CERRADO_DT] = pd.to_datetime(trabajo[TEXT_CERRADO].apply(normalizar_fecha), errors=TEXT_COERCE)
    trabajo[TEXT_DURACION_HORAS_NUM] = pd.to_numeric(trabajo[TEXT_DURACION_HORAS], errors=TEXT_COERCE)
    return trabajo


@st.cache_data(ttl=CACHE_TTL_SEGUNDOS, show_spinner=False)
def cargar_casos_clientes_clave_cache():
    return preparar_casos_clientes_clave(cargar_casos_cache())


@st.cache_data(ttl=CACHE_TTL_SEGUNDOS, show_spinner=False)
def cargar_casos_clientes_clave_filtrados_cache(anio=None, mes=None):
    return preparar_casos_clientes_clave(cargar_casos_filtrados_cache(anio, mes))


@st.cache_data(ttl=CACHE_TTL_SEGUNDOS, show_spinner=False)
def cargar_incidentes_clientes_clave_cache():
    return preparar_incidentes_clientes_clave(cargar_incidentes_cache())


@st.cache_data(ttl=CACHE_TTL_SEGUNDOS, show_spinner=False)
def cargar_incidentes_clientes_clave_filtrados_cache(anio=None, mes=None):
    return preparar_incidentes_clientes_clave(cargar_incidentes_filtrados_cache(anio, mes))


def fecha_maxima_cliente(casos_cliente, incidentes_cliente):
    fechas = []
    if not casos_cliente.empty:
        fechas.append(casos_cliente[TEXT_CREADO_DT].max())
    if not incidentes_cliente.empty:
        fechas.append(incidentes_cliente[TEXT_CREADO_DT].max())
    fechas_validas = [fecha for fecha in fechas if pd.notna(fecha)]
    if not fechas_validas:
        return ""
    return max(fechas_validas).strftime("%Y-%m-%d %H:%M")


def valor_mas_frecuente(serie, default=SIN_DATO):
    if serie.empty:
        return default
    valores = serie.replace("", pd.NA).dropna()
    if valores.empty:
        return default
    return valores.value_counts().idxmax()


def nivel_atencion_cliente(
    abiertos,
    incidentes_abiertos,
    alertas,
    prioridad_alta,
    sla_casos,
    sla_incidentes,
    casos_sin_causa,
):
    score = 100
    score -= min(25, abiertos * 5)
    score -= min(20, incidentes_abiertos * 10)
    score -= min(20, alertas * 4)
    score -= min(12, prioridad_alta * 6)
    score -= min(10, casos_sin_causa * 2)

    if sla_casos is not None and sla_casos < 90:
        score -= min(15, round((90 - sla_casos) / 3, 2))
    if sla_incidentes is not None and sla_incidentes < 90:
        score -= min(15, round((90 - sla_incidentes) / 3, 2))

    score = max(0, round(score, 2))
    if score >= 85:
        return "Verde", "Estable", score
    if score >= 65:
        return TEXT_AMARILLO, "Seguimiento", score
    return "Rojo", "Prioritario", score


def fila_cliente_sin_actividad(cliente):
    return {
        TEXT_CLIENTE: cliente,
        TEXT_NIVEL: "Sin actividad",
        COL_ESTADO_ATENCION: "Sin actividad",
        TEXT_SCORE: 100,
        TEXT_CASOS: 0,
        TEXT_INCIDENTES: 0,
        COL_TOTAL_ATENCIONES: 0,
        TEXT_ABIERTOS: 0,
        COL_SLA_CASOS_PCT: None,
        COL_SLA_INCIDENTES_PCT: None,
        COL_ALERTAS_INCIDENTES: 0,
        PRIORIDAD_ALTA: 0,
        COL_CASOS_SIN_CAUSA: 0,
        COL_PRODUCTO_PRINCIPAL: SIN_DATO,
        COL_CAUSA_INCIDENTE_PRINCIPAL: SIN_DATO,
        COL_ULTIMA_ATENCION: "",
    }


def metricas_casos_cliente(casos_cliente):
    if casos_cliente.empty:
        return 0, None, 0

    casos_cerrados = casos_cliente[mascara_cerrados(casos_cliente)]
    tiempos_casos = casos_cerrados[TEXT_TIEMPO_RESPUESTA_H].dropna()
    sla_casos = (
        porcentaje(len(tiempos_casos[tiempos_casos < SLA_CASOS_HORAS]), len(tiempos_casos))
        if len(tiempos_casos)
        else None
    )
    casos_sin_causa = len(
        casos_cliente[
            casos_cliente[TEXT_CAUSA].replace("", pd.NA).fillna(SIN_DATO).str.lower().isin(["sin dato"])
        ]
    )
    return len(casos_cliente) - len(casos_cerrados), sla_casos, casos_sin_causa


def metricas_incidentes_cliente(incidentes_cliente):
    if incidentes_cliente.empty:
        return 0, None, 0, 0

    incidentes_cerrados_todos = incidentes_cliente[mascara_cerrados(incidentes_cliente)]
    incidentes_cerrados = incidentes_cerrados_todos[
        incidentes_cerrados_todos[TEXT_APLICA_SLA_INCIDENTE].fillna(False)
    ]
    sla_incidentes_base = incidentes_cerrados[
        incidentes_cerrados[TEXT_ESTADO_SLA].isin([ESTADO_SLA_CUMPLE, ESTADO_SLA_NO_CUMPLE])
    ]
    sla_incidentes = (
        porcentaje(
            len(sla_incidentes_base[sla_incidentes_base[TEXT_ESTADO_SLA] == ESTADO_SLA_CUMPLE]),
            len(sla_incidentes_base),
        )
        if len(sla_incidentes_base)
        else None
    )
    alertas = len(
        incidentes_cliente[
            (incidentes_cliente[TEXT_ES_ALERTA_AUTO].fillna("No") == "Si")
            | (incidentes_cliente[TEXT_TIPIFICACION_AUTO].fillna("") == NOC_TIPIFICATION)
        ]
    )
    prioridad_alta = len(
        incidentes_cliente[
            incidentes_cliente[TEXT_PRIORIDAD].fillna("").str.contains(
                r"^(?:1|2\s*-\s*Alta)", case=False, regex=True
            )
        ]
    )
    return len(incidentes_cliente) - len(incidentes_cerrados_todos), sla_incidentes, alertas, prioridad_alta


def fila_cliente_con_actividad(cliente, casos_cliente, incidentes_cliente):
    total_casos = len(casos_cliente)
    total_incidentes = len(incidentes_cliente)
    casos_abiertos, sla_casos, casos_sin_causa = metricas_casos_cliente(casos_cliente)
    incidentes_abiertos, sla_incidentes, alertas, prioridad_alta = metricas_incidentes_cliente(incidentes_cliente)
    abiertos = casos_abiertos + incidentes_abiertos
    nivel, estado_atencion, score = nivel_atencion_cliente(
        abiertos,
        incidentes_abiertos,
        alertas,
        prioridad_alta,
        sla_casos,
        sla_incidentes,
        casos_sin_causa,
    )

    return {
        TEXT_CLIENTE: cliente,
        TEXT_NIVEL: nivel,
        COL_ESTADO_ATENCION: estado_atencion,
        TEXT_SCORE: score,
        TEXT_CASOS: total_casos,
        TEXT_INCIDENTES: total_incidentes,
        COL_TOTAL_ATENCIONES: total_casos + total_incidentes,
        TEXT_ABIERTOS: abiertos,
        COL_SLA_CASOS_PCT: sla_casos,
        COL_SLA_INCIDENTES_PCT: sla_incidentes,
        COL_ALERTAS_INCIDENTES: alertas,
        PRIORIDAD_ALTA: prioridad_alta,
        COL_CASOS_SIN_CAUSA: casos_sin_causa,
        COL_PRODUCTO_PRINCIPAL: valor_mas_frecuente(casos_cliente.get(TEXT_PRODUCTO, pd.Series(dtype=TEXT_OBJECT))),
        COL_CAUSA_INCIDENTE_PRINCIPAL: valor_mas_frecuente(
            incidentes_cliente.get(TEXT_CAUSA_RAIZ_AUTO, pd.Series(dtype=TEXT_OBJECT))
        ),
        COL_ULTIMA_ATENCION: fecha_maxima_cliente(casos_cliente, incidentes_cliente),
    }


def resumen_clientes_clave(casos, incidentes):
    filas = []
    for cliente in CLIENTES_CLAVE:
        casos_cliente = casos[casos[TEXT_CLIENTE_CLAVE] == cliente] if not casos.empty else pd.DataFrame()
        incidentes_cliente = (
            incidentes[incidentes[TEXT_CLIENTE_CLAVE] == cliente] if not incidentes.empty else pd.DataFrame()
        )
        if casos_cliente.empty and incidentes_cliente.empty:
            filas.append(fila_cliente_sin_actividad(cliente))
            continue
        filas.append(fila_cliente_con_actividad(cliente, casos_cliente, incidentes_cliente))

    return pd.DataFrame(filas)

def tabla_atenciones_abiertas_clientes(casos, incidentes):
    tablas = []
    if not casos.empty:
        abiertos_casos = casos[~mascara_cerrados(casos)].copy()
        if not abiertos_casos.empty:
            abiertos_casos["Tipo"] = "Caso"
            abiertos_casos[TEXT_CLIENTE] = abiertos_casos[TEXT_CLIENTE_CLAVE]
            abiertos_casos[TEXT_NUMERO_2] = abiertos_casos[TEXT_NUMERO]
            abiertos_casos[TEXT_ESTADO_2] = abiertos_casos[TEXT_ESTADO]
            abiertos_casos[TEXT_PRIORIDAD_2] = abiertos_casos[TEXT_PRIORIDAD]
            abiertos_casos[TEXT_RESPONSABLE] = abiertos_casos[TEXT_ASIGNADO]
            abiertos_casos[TEXT_CREADO_2] = abiertos_casos[TEXT_CREADO]
            abiertos_casos[COL_VENCIMIENTO_SLA] = ""
            abiertos_casos[TEXT_CLASIFICACION] = abiertos_casos[TEXT_TIPIFICACION_2]
            abiertos_casos[TEXT_RESUMEN] = abiertos_casos[TEXT_DESCRIPCION_2]
            tablas.append(
                abiertos_casos[
                    [
                        "Tipo",
                        TEXT_CLIENTE,
                        TEXT_NUMERO_2,
                        TEXT_ESTADO_2,
                        TEXT_PRIORIDAD_2,
                        TEXT_RESPONSABLE,
                        TEXT_CREADO_2,
                        COL_VENCIMIENTO_SLA,
                        TEXT_CLASIFICACION,
                        TEXT_RESUMEN,
                    ]
                ]
            )

    if not incidentes.empty:
        abiertos_incidentes = incidentes[~mascara_cerrados(incidentes)].copy()
        if not abiertos_incidentes.empty:
            abiertos_incidentes["Tipo"] = "Incidente"
            abiertos_incidentes[TEXT_CLIENTE] = abiertos_incidentes[TEXT_CLIENTE_CLAVE]
            abiertos_incidentes[TEXT_NUMERO_2] = abiertos_incidentes[TEXT_NUMERO]
            abiertos_incidentes[TEXT_ESTADO_2] = abiertos_incidentes[TEXT_ESTADO]
            abiertos_incidentes[TEXT_PRIORIDAD_2] = abiertos_incidentes[TEXT_PRIORIDAD]
            abiertos_incidentes[TEXT_RESPONSABLE] = abiertos_incidentes[TEXT_ASIGNADO_A]
            abiertos_incidentes[TEXT_CREADO_2] = abiertos_incidentes[TEXT_CREADO]
            abiertos_incidentes[COL_VENCIMIENTO_SLA] = abiertos_incidentes[TEXT_FECHA_VENCIMIENTO_SLA]
            abiertos_incidentes[TEXT_CLASIFICACION] = abiertos_incidentes[TEXT_TIPIFICACION_AUTO]
            abiertos_incidentes[TEXT_RESUMEN] = abiertos_incidentes[TEXT_BREVE_DESCRIPCION].replace("", pd.NA).fillna(
                abiertos_incidentes[TEXT_DESCRIPCION_2]
            )
            tablas.append(
                abiertos_incidentes[
                    [
                        "Tipo",
                        TEXT_CLIENTE,
                        TEXT_NUMERO_2,
                        TEXT_ESTADO_2,
                        TEXT_PRIORIDAD_2,
                        TEXT_RESPONSABLE,
                        TEXT_CREADO_2,
                        COL_VENCIMIENTO_SLA,
                        TEXT_CLASIFICACION,
                        TEXT_RESUMEN,
                    ]
                ]
            )

    if not tablas:
        return pd.DataFrame(
            columns=[
                "Tipo",
                TEXT_CLIENTE,
                TEXT_NUMERO_2,
                TEXT_ESTADO_2,
                TEXT_PRIORIDAD_2,
                TEXT_RESPONSABLE,
                TEXT_CREADO_2,
                COL_VENCIMIENTO_SLA,
                TEXT_CLASIFICACION,
                TEXT_RESUMEN,
            ]
        )
    return pd.concat(tablas, ignore_index=True).sort_values(by=[TEXT_CLIENTE, "Tipo", TEXT_CREADO_2])


def tabla_resumen_tipificaciones_casos(df):
    conteo_tipificaciones = df[TEXT_TIPIFICACION_2].value_counts()
    resumen = pd.DataFrame(CASE_TIPIFICATION_GUIDE)
    resumen[TEXT_CANTIDAD] = resumen[TEXT_TIPIFICACION].map(conteo_tipificaciones).fillna(0).astype(int)

    tipificaciones_faltantes = [
        tipificacion
        for tipificacion in conteo_tipificaciones.index.tolist()
        if tipificacion not in resumen[TEXT_TIPIFICACION].tolist()
    ]
    if tipificaciones_faltantes:
        adicionales = pd.DataFrame(
            {
                TEXT_TIPIFICACION: tipificaciones_faltantes,
                TEXT_DESCRIPCION: ["Tipificacion detectada sin descripcion configurada."] * len(tipificaciones_faltantes),
                TEXT_CANTIDAD: [int(conteo_tipificaciones[tip]) for tip in tipificaciones_faltantes],
            }
        )
        resumen = pd.concat([resumen, adicionales], ignore_index=True)

    return resumen.sort_values(by=[TEXT_CANTIDAD, TEXT_TIPIFICACION], ascending=[False, True]).reset_index(drop=True)


def texto_caso_para_causa_comun(row):
    campos = [
        TEXT_DESCRIPCION_2,
        TEXT_CAUSA,
        TEXT_CODIGO_RESOLUCION,
        "notas_resolucion",
        TEXT_OBSERVACIONES_ADICIONALES,
        TEXT_OBSERVACIONES_TRABAJO,
        TEXT_PRODUCTO,
    ]
    return " ".join(normalizar_texto(row.get(campo)) for campo in campos).strip()


def texto_caso_para_tipologia_soporte(row):
    campos = [
        TEXT_DESCRIPCION_2,
        TEXT_CAUSA,
        TEXT_CODIGO_RESOLUCION,
        "notas_resolucion",
        TEXT_OBSERVACIONES_ADICIONALES,
        TEXT_OBSERVACIONES_TRABAJO,
        TEXT_PRODUCTO,
        TEXT_CUENTA,
        TEXT_TIPIFICACION_2,
        TEXT_CANAL,
    ]
    return " ".join(normalizar_texto(row.get(campo)) for campo in campos).strip()


def texto_contiene_alguno(texto, palabras):
    return any(palabra in texto for palabra in palabras)


def clasificar_tipologia_soporte_caso(row):
    texto = texto_caso_para_tipologia_soporte(row)
    tipificacion = normalizar_texto(row.get(TEXT_TIPIFICACION_2))

    if "5 - incidente" in tipificacion or "1 - phishing" in tipificacion:
        return INCIDENTES_CASOS
    if texto_contiene_alguno(texto, CASE_INCIDENT_TERMS):
        return INCIDENTES_CASOS
    if texto_contiene_alguno(texto, CASE_AGENDA_DIRECT_TERMS):
        return ENVIO_AGENDA_MANUAL_USO
    if "4 - solicitudes" in tipificacion:
        return SOLICITUDES_CASOS
    if texto_contiene_alguno(texto, CASE_REQUEST_TERMS):
        return SOLICITUDES_CASOS
    return SOPORTE_USO


def agregar_tipologia_soporte_casos(df):
    trabajo = df.copy()
    if trabajo.empty:
        trabajo[TEXT_TIPOLOGIA_SOPORTE] = pd.Series(dtype=TEXT_OBJECT)
    else:
        trabajo[TEXT_TIPOLOGIA_SOPORTE] = trabajo.apply(clasificar_tipologia_soporte_caso, axis=1)
    return trabajo


@st.cache_data(ttl=CACHE_TTL_SEGUNDOS, show_spinner=False)
def cargar_casos_soporte_cache(_version=CASE_TYPOLOGY_CACHE_VERSION):
    return agregar_tipologia_soporte_casos(normalizar_tipificaciones_casos_df(cargar_casos_cache()))


@st.cache_data(ttl=CACHE_TTL_SEGUNDOS, show_spinner=False)
def cargar_casos_soporte_filtrados_cache(
    anio=None,
    mes=None,
    cliente="",
    estado="",
    servicio="",
    tipificacion="",
    _version=CASE_TYPOLOGY_CACHE_VERSION,
):
    return agregar_tipologia_soporte_casos(
        normalizar_tipificaciones_casos_df(
            cargar_casos_filtrados_cache(anio, mes, cliente, estado, servicio, tipificacion)
        )
    )


def resumen_tipologias_soporte_casos(base):
    columnas = [TEXT_TIPOLOGIA_SOPORTE, TEXT_CANTIDAD, "% casos", TEXT_DESCRIPCION]
    if base.empty or TEXT_TIPOLOGIA_SOPORTE not in base.columns:
        return pd.DataFrame(columns=columnas)
    conteo = base[TEXT_TIPOLOGIA_SOPORTE].value_counts()
    total = int(conteo.sum())
    filas = []
    descripciones = {
        item[TEXT_TIPOLOGIA_SOPORTE]: item[TEXT_DESCRIPCION]
        for item in CASE_SUPPORT_TYPOLOGY_GUIDE
    }
    for tipologia in CASE_SUPPORT_TYPOLOGY_ORDER:
        cantidad = int(conteo.get(tipologia, 0))
        filas.append(
            {
                TEXT_TIPOLOGIA_SOPORTE: tipologia,
                TEXT_CANTIDAD: cantidad,
                "% casos": porcentaje(cantidad, total),
                TEXT_DESCRIPCION: descripciones.get(tipologia, ""),
            }
        )
    return pd.DataFrame(filas, columns=columnas)


def inferir_causa_comun_caso(row):
    causa_existente = valor_limpio(row.get(TEXT_CAUSA_COMUN))
    if causa_existente:
        return causa_existente

    texto = texto_caso_para_causa_comun(row)
    for causa, palabras in CASE_COMMON_CAUSE_RULES:
        if any(palabra in texto for palabra in palabras):
            return causa

    causa_base = valor_limpio(row.get(TEXT_CAUSA))
    if causa_base:
        return causa_base
    return "Sin causa comun"


def inferir_detalle_causa_comun(row):
    texto = texto_caso_para_causa_comun(row)
    for detalle, palabras in CASE_CAUSE_DETAIL_RULES:
        if any(palabra in texto for palabra in palabras):
            return detalle
    return "Sin detalle especifico"


def resumen_detalle_causa_principal(base, causa_principal, top_n=3):
    if base.empty or TEXT_CAUSA_COMUN not in base.columns:
        return "No hay detalle suficiente para explicar mejor esta causa."

    casos_causa = base[base[TEXT_CAUSA_COMUN] == causa_principal].copy()
    if casos_causa.empty:
        return "No hay detalle suficiente para explicar mejor esta causa."

    detalles = casos_causa.apply(inferir_detalle_causa_comun, axis=1).value_counts()
    detalles = detalles[detalles.index != "Sin detalle especifico"]
    if detalles.empty:
        return f"Esta causa agrupa {len(casos_causa)} casos, pero el archivo no trae detalle suficiente del motivo."

    principales = ", ".join(f"{detalle} ({cantidad})" for detalle, cantidad in detalles.head(top_n).items())
    return f"Dentro de esta causa entran principalmente: {principales}."


def serie_categorica_limpia(df, columna, etiqueta_vacia):
    if df.empty or columna not in df.columns:
        return pd.Series([etiqueta_vacia] * len(df), index=df.index, dtype=TEXT_OBJECT)
    return df[columna].replace("", pd.NA).fillna(etiqueta_vacia).astype(str)


def conteo_top_con_otras(serie, top_n=3, etiqueta_otras="Otras categorias"):
    conteo = serie.value_counts(dropna=False)
    top = conteo.head(top_n)
    resumen = top.rename_axis(TEXT_TIPOLOGIA).reset_index(name=TEXT_CANTIDAD)
    otras = int(conteo.iloc[top_n:].sum())
    if otras > 0:
        resumen = pd.concat(
            [
                resumen,
                pd.DataFrame({TEXT_TIPOLOGIA: [etiqueta_otras], TEXT_CANTIDAD: [otras]}),
            ],
            ignore_index=True,
        )
    return resumen


def conteo_top(serie, etiqueta_columna, top_n=5):
    return (
        serie.value_counts(dropna=False)
        .head(top_n)
        .rename_axis(etiqueta_columna)
        .reset_index(name=TEXT_CANTIDAD)
    )


def preparar_kpi_casos_cliente_externo(df):
    trabajo = df.copy()
    if trabajo.empty:
        return trabajo, {}

    trabajo[TEXT_CAUSA_COMUN] = trabajo.apply(inferir_causa_comun_caso, axis=1)
    trabajo = agregar_tipologia_soporte_casos(trabajo)
    trabajo["_tipificacion_kpi"] = serie_categorica_limpia(trabajo, TEXT_TIPIFICACION_2, "Sin tipificacion")
    trabajo[TEXT_CERRADO_2] = mascara_cerrados(trabajo)
    trabajo[TEXT_ABIERTO] = ~trabajo[TEXT_CERRADO_2]
    trabajo["_tiempo_cerrado_h"] = pd.to_numeric(trabajo.get(TEXT_TIEMPO_RESPUESTA), errors=TEXT_COERCE)
    trabajo[TEXT_CREADO_DT_2] = pd.to_datetime(trabajo[TEXT_CREADO].apply(normalizar_fecha), errors=TEXT_COERCE)
    ahora = pd.Timestamp.now()
    trabajo[TEXT_HORAS_ABIERTO] = trabajo[TEXT_CREADO_DT_2].apply(lambda fecha: horas_habiles_entre(fecha, ahora))
    trabajo["_tiempo_eval_sla_h"] = trabajo["_tiempo_cerrado_h"].where(
        trabajo[TEXT_CERRADO_2],
        trabajo[TEXT_HORAS_ABIERTO],
    )
    trabajo["Cumple SLA <=36h"] = trabajo["_tiempo_eval_sla_h"].apply(
        lambda valor: "Si" if pd.notna(valor) and valor <= SLA_CASOS_HORAS else "No"
    )

    total = len(trabajo)
    cerrados = int(trabajo[TEXT_CERRADO_2].sum())
    abiertos = total - cerrados
    tiempos_validos = pd.to_numeric(trabajo["_tiempo_eval_sla_h"], errors=TEXT_COERCE).dropna()
    promedio = round(tiempos_validos.mean(), 2) if not tiempos_validos.empty else 0
    cumple_sla = int((trabajo["_tiempo_eval_sla_h"] <= SLA_CASOS_HORAS).fillna(False).sum())
    no_cumple_sla = total - cumple_sla
    cumplimiento_sla = porcentaje(cumple_sla, total)

    metricas = {
        "total": total,
        "cerrados": cerrados,
        "abiertos": abiertos,
        "promedio": promedio,
        "cumplimiento_sla": cumplimiento_sla,
        "cumple_sla": cumple_sla,
        "no_cumple_sla": no_cumple_sla,
        COL_PRINCIPAL_TIPIFICACION: valor_mas_frecuente(trabajo["_tipificacion_kpi"]),
        COL_PRINCIPAL_SOPORTE: valor_mas_frecuente(trabajo[TEXT_TIPOLOGIA_SOPORTE]),
        COL_PRINCIPAL_CAUSA_COMUN: valor_mas_frecuente(trabajo[TEXT_CAUSA_COMUN]),
    }
    return trabajo, metricas


def texto_ranking_kpi(valor, limite=78):
    texto = valor_limpio(valor)
    if len(texto) <= limite:
        return texto
    return texto[: limite - 3].rstrip() + "..."


def numero_ranking_kpi(valor):
    try:
        numero = float(valor)
    except (TypeError, ValueError):
        return str(valor)
    if numero.is_integer():
        return str(int(numero))
    return str(round(numero, 2))


def render_ranking_kpi(df, etiqueta_columna, valor_columna, titulo, top_n=6):
    if df.empty:
        st.info(f"No hay datos para {titulo.lower()}.")
        return

    ranking = df.copy()
    ranking[valor_columna] = pd.to_numeric(ranking[valor_columna], errors=TEXT_COERCE).fillna(0)
    ranking = ranking[ranking[valor_columna] > 0].sort_values(by=valor_columna, ascending=False).head(top_n)
    if ranking.empty:
        st.info(f"No hay datos para {titulo.lower()}.")
        return

    maximo = ranking[valor_columna].max()
    filas = []
    for _, row in ranking.iterrows():
        valor = row[valor_columna]
        porcentaje_barra = (valor / maximo) * 100 if maximo else 0
        etiqueta_completa = valor_limpio(row[etiqueta_columna]) or SIN_DATO
        etiqueta = texto_ranking_kpi(etiqueta_completa)
        filas.append(
            '<div class="kpi-ranking-row">'
            f'<div class="kpi-ranking-label" title="{html.escape(etiqueta_completa)}">{html.escape(etiqueta)}</div>'
            '<div class="kpi-ranking-track">'
            f'<div class="kpi-ranking-bar" style="width: max(8px, {porcentaje_barra:.2f}%);"></div>'
            "</div>"
            f'<div class="kpi-ranking-value">{html.escape(numero_ranking_kpi(valor))}</div>'
            "</div>"
        )

    contenido = (
        '<div class="kpi-ranking">'
        f'<div class="kpi-ranking-title">{html.escape(str(titulo))}</div>'
        f'<div class="kpi-ranking-list">{"".join(filas)}</div>'
        "</div>"
    )
    st.markdown(contenido, unsafe_allow_html=True)


def slide_kpi_cards_html(items):
    tarjetas = []
    for titulo, valor in items:
        tarjetas.append(
            '<div class="slide-kpi-card">'
            f'<div class="slide-kpi-title">{html.escape(str(titulo))}</div>'
            f'<div class="slide-kpi-value">{html.escape(str(valor))}</div>'
            "</div>"
        )
    return f'<div class="slide-kpi-grid">{"".join(tarjetas)}</div>'


def slide_ranking_html(df, etiqueta_columna, valor_columna, titulo, top_n=6, limite=68):
    if df.empty:
        filas = '<div class="kpi-ranking-empty">Sin datos para mostrar.</div>'
    else:
        ranking = df.copy()
        ranking[valor_columna] = pd.to_numeric(ranking[valor_columna], errors=TEXT_COERCE).fillna(0)
        ranking = ranking[ranking[valor_columna] > 0].sort_values(by=valor_columna, ascending=False).head(top_n)
        maximo = ranking[valor_columna].max() if not ranking.empty else 0
        filas_lista = []
        for _, row in ranking.iterrows():
            valor = row[valor_columna]
            porcentaje_barra = (valor / maximo) * 100 if maximo else 0
            etiqueta_completa = valor_limpio(row[etiqueta_columna]) or SIN_DATO
            etiqueta = texto_ranking_kpi(etiqueta_completa, limite)
            filas_lista.append(
                '<div class="slide-ranking-row">'
                f'<div class="slide-ranking-label" title="{html.escape(etiqueta_completa)}">{html.escape(etiqueta)}</div>'
                '<div class="slide-ranking-track">'
                f'<div class="slide-ranking-bar" style="width: max(8px, {porcentaje_barra:.2f}%);"></div>'
                "</div>"
                f'<div class="slide-ranking-value">{html.escape(numero_ranking_kpi(valor))}</div>'
                "</div>"
            )
        filas = "".join(filas_lista) if filas_lista else '<div class="kpi-ranking-empty">Sin datos para mostrar.</div>'

    return (
        '<div class="slide-panel">'
        f'<div class="slide-panel-title">{html.escape(str(titulo))}</div>'
        f'<div class="slide-ranking-list">{filas}</div>'
        "</div>"
    )


def slide_note_html(titulo, lineas):
    contenido_lineas = "".join(f'<div>{linea}</div>' for linea in lineas if linea)
    return (
        '<div class="slide-panel">'
        f'<div class="slide-panel-title">{html.escape(str(titulo))}</div>'
        f'<div class="slide-note">{contenido_lineas}</div>'
        "</div>"
    )


def slide_frame_html(titulo, periodo, tarjetas, caption, izquierda_html, derecha_html, frame_class=""):
    periodo_html = f"{TEXT_PERIODO}{periodo}" if periodo else ""
    clases = "slide-frame"
    if frame_class:
        clases += f" {html.escape(str(frame_class))}"
    return (
        f'<div class="{clases}" id="kpi-slide-frame" data-kpi-slide-id="kpi-slide-frame">'
        f'<div class="slide-period">{html.escape(periodo_html)}</div>'
        f'<div class="slide-title">{html.escape(str(titulo))}</div>'
        f"{slide_kpi_cards_html(tarjetas)}"
        f'<div class="slide-caption">{html.escape(str(caption))}</div>'
        '<div class="slide-body">'
        f"{izquierda_html}"
        f"{derecha_html}"
        "</div>"
        "</div>"
    )


def render_slide_frame_kpi(titulo, periodo, tarjetas, caption, izquierda_html, derecha_html, frame_class=""):
    st.markdown(
        slide_frame_html(titulo, periodo, tarjetas, caption, izquierda_html, derecha_html, frame_class),
        unsafe_allow_html=True,
    )


def slide_component_css():
    return f"""
    :root {{
        --border: {UI_PALETTE["border"]};
        --text: {UI_PALETTE["text"]};
        --muted: {UI_PALETTE["muted"]};
        --primary: {UI_PALETTE[TEXT_PRIMARY]};
        --mustard: {UI_PALETTE["mustard"]};
    }}

    html, body {{
        margin: 0;
        padding: 0;
        background: #ffffff;
        color: var(--text);
        font-family: Arial, sans-serif;
    }}

    * {{
        box-sizing: border-box;
    }}

    .slide-export-shell {{
        width: 100%;
        background: #ffffff;
    }}

    .slide-frame {{
        width: 100%;
        max-width: 1600px;
        aspect-ratio: 16 / 9;
        background: #ffffff;
        border: 1px solid var(--border);
        border-radius: 8px;
        box-shadow: 0 12px 28px rgba(20, 20, 20, 0.08);
        padding: 28px 32px;
        display: flex;
        flex-direction: column;
        gap: 16px;
        overflow: hidden;
    }}

    .slide-period {{
        color: var(--muted);
        font-size: 1.04rem;
        font-weight: 700;
        line-height: 1.15;
    }}

    .slide-title {{
        color: var(--text);
        font-size: 1.85rem;
        font-weight: 900;
        line-height: 1.12;
    }}

    .slide-kpi-grid {{
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(165px, 1fr));
        gap: 12px;
    }}

    .slide-kpi-card {{
        background: #ffffff;
        border: 1px solid var(--border);
        border-top: 4px solid var(--primary);
        border-radius: 8px;
        min-height: 104px;
        padding: 16px 12px;
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        text-align: center;
    }}

    .slide-kpi-title {{
        color: var(--text);
        font-size: 0.92rem;
        font-weight: 900;
        line-height: 1.18;
        text-transform: uppercase;
    }}

    .slide-kpi-value {{
        color: var(--primary);
        font-size: 2.35rem;
        font-weight: 900;
        line-height: 1;
        margin-top: 0.42rem;
        font-variant-numeric: tabular-nums;
    }}

    .slide-caption {{
        color: var(--muted);
        font-size: 1.02rem;
        font-weight: 700;
        line-height: 1.25;
    }}

    .slide-body {{
        display: grid;
        grid-template-columns: minmax(0, 2fr) minmax(300px, 0.9fr);
        gap: 14px;
        min-height: 0;
        flex: 1;
    }}

    .slide-panel {{
        background: #ffffff;
        border: 1px solid var(--border);
        border-radius: 8px;
        padding: 16px 18px;
        min-height: 0;
        overflow: hidden;
    }}

    .slide-panel-group {{
        display: grid;
        grid-template-columns: repeat(2, minmax(0, 1fr));
        gap: 14px;
        min-height: 0;
    }}

    .slide-panel-title {{
        color: var(--primary);
        font-size: 1.18rem;
        font-weight: 800;
        line-height: 1.2;
        margin-bottom: 0.9rem;
    }}

    .slide-ranking-list {{
        display: flex;
        flex-direction: column;
        gap: 0.64rem;
    }}

    .slide-ranking-row {{
        display: grid;
        grid-template-columns: minmax(0, 1.6fr) minmax(120px, 1fr) 3.1rem;
        gap: 0.8rem;
        align-items: center;
    }}

    .slide-ranking-label {{
        color: var(--text);
        font-size: 0.96rem;
        font-weight: 650;
        line-height: 1.18;
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
    }}

    .slide-ranking-track {{
        height: 17px;
        background: #f1f3f5;
        border-radius: 999px;
        overflow: hidden;
    }}

    .slide-ranking-bar {{
        height: 100%;
        background: var(--mustard);
        border-radius: inherit;
    }}

    .slide-ranking-value {{
        color: var(--text);
        font-size: 1.05rem;
        font-weight: 800;
        text-align: right;
        font-variant-numeric: tabular-nums;
    }}

    .kpi-ranking-empty {{
        color: var(--muted);
        font-size: 1rem;
        font-weight: 700;
    }}

    .slide-note {{
        display: flex;
        flex-direction: column;
        gap: 0.64rem;
        color: var(--text);
        font-size: 1rem;
        font-weight: 500;
        line-height: 1.38;
    }}

    .slide-note strong {{
        color: var(--text);
        font-weight: 700;
    }}

    .slide-note-muted {{
        color: var(--muted);
        font-weight: 500;
    }}

    .slide-cause-table {{
        width: 100%;
        border-collapse: collapse;
        table-layout: fixed;
    }}

    .slide-cause-table th,
    .slide-cause-table td {{
        border-top: 1px solid var(--border);
        color: var(--text);
        font-size: 0.92rem;
        line-height: 1.25;
        padding: 10px 9px;
        text-align: left;
        vertical-align: middle;
        white-space: normal;
        overflow-wrap: anywhere;
    }}

    .slide-cause-table th {{
        background: #fff7f2;
        color: var(--muted);
        font-size: 0.75rem;
        font-weight: 900;
        text-transform: uppercase;
    }}

    .slide-cause-table .number-header {{
        text-align: right;
    }}

    .slide-cause-table .cause-cell {{
        font-weight: 700;
    }}

    .slide-cause-table .number-cell {{
        color: var(--primary);
        font-weight: 900;
        text-align: right;
        white-space: nowrap;
        font-variant-numeric: tabular-nums;
    }}

    .slide-download-bar {{
        max-width: 1600px;
        display: flex;
        align-items: center;
        gap: 12px;
        margin-top: 12px;
    }}

    #download-kpi-slide {{
        min-height: 42px;
        border: 0;
        border-radius: 8px;
        background: var(--primary);
        color: #ffffff;
        padding: 0 20px;
        font-size: 15px;
        font-weight: 800;
        cursor: pointer;
    }}

    #download-kpi-slide:disabled {{
        opacity: 0.7;
        cursor: wait;
    }}

    #download-kpi-status {{
        color: var(--muted);
        font-size: 12px;
        min-height: 16px;
    }}
    """


def render_slide_component_kpi(titulo, periodo, tarjetas, caption, izquierda_html, derecha_html):
    archivo = nombre_archivo_slide(titulo, periodo)
    slide_html = slide_frame_html(titulo, periodo, tarjetas, caption, izquierda_html, derecha_html)
    componente = f"""
    <style>{slide_component_css()}</style>
    <div class="slide-export-shell">
        {slide_html}
        <div class="slide-download-bar">
            <button id="download-kpi-slide" type="button">Descargar imagen PNG</button>
            <div id="download-kpi-status"></div>
        </div>
    </div>
    <script>
    const fileName = {json.dumps(archivo)};
    const button = document.getElementById("download-kpi-slide");
    const status = document.getElementById("download-kpi-status");

    function setStatus(message) {{
        status.textContent = message || "";
    }}

    function copyComputedStyles(source, target) {{
        const computed = window.getComputedStyle(source);
        for (const property of computed) {{
            target.style.setProperty(
                property,
                computed.getPropertyValue(property),
                computed.getPropertyPriority(property)
            );
        }}
        target.style.transform = "none";
        target.style.animation = "none";
        target.style.transition = "none";
    }}

    function inlineStyles(source, target) {{
        copyComputedStyles(source, target);
        const sourceChildren = Array.from(source.children || []);
        const targetChildren = Array.from(target.children || []);
        for (let index = 0; index < sourceChildren.length; index += 1) {{
            inlineStyles(sourceChildren[index], targetChildren[index]);
        }}
    }}

    async function captureSlide() {{
        const slide = document.getElementById("kpi-slide-frame");
        if (!slide) {{
            throw new Error("No se encontro el recuadro del slide.");
        }}

        const rect = slide.getBoundingClientRect();
        const clone = slide.cloneNode(true);
        inlineStyles(slide, clone);
        clone.style.margin = "0";
        clone.style.width = `${{rect.width}}px`;
        clone.style.height = `${{rect.height}}px`;
        clone.style.boxSizing = "border-box";

        const wrapper = document.createElement("div");
        wrapper.setAttribute("xmlns", "http://www.w3.org/1999/xhtml");
        wrapper.style.width = `${{rect.width}}px`;
        wrapper.style.height = `${{rect.height}}px`;
        wrapper.style.background = "#ffffff";
        wrapper.appendChild(clone);

        const serialized = new XMLSerializer().serializeToString(wrapper);
        const svg = `
            <svg xmlns="http://www.w3.org/2000/svg"
                 width="${{rect.width}}" height="${{rect.height}}"
                 viewBox="0 0 ${{rect.width}} ${{rect.height}}">
                <foreignObject width="100%" height="100%">
                    ${{serialized}}
                </foreignObject>
            </svg>
        `;

        const blob = new Blob([svg], {{ type: "image/svg+xml;charset=utf-8" }});
        const url = URL.createObjectURL(blob);
        try {{
            const image = new Image();
            await new Promise((resolve, reject) => {{
                image.onload = resolve;
                image.onerror = reject;
                image.src = url;
            }});

            const targetWidth = 1920;
            const scale = targetWidth / rect.width;
            const canvas = document.createElement("canvas");
            canvas.width = Math.round(rect.width * scale);
            canvas.height = Math.round(rect.height * scale);
            const context = canvas.getContext("2d");
            context.fillStyle = "#ffffff";
            context.fillRect(0, 0, canvas.width, canvas.height);
            context.scale(scale, scale);
            context.drawImage(image, 0, 0);

            const link = document.createElement("a");
            link.href = canvas.toDataURL("image/png");
            link.download = fileName;
            document.body.appendChild(link);
            link.click();
            link.remove();
        }} finally {{
            URL.revokeObjectURL(url);
        }}
    }}

    button.addEventListener("click", async () => {{
        button.disabled = true;
        setStatus("Generando imagen desde el recuadro visible...");
        try {{
            await captureSlide();
            setStatus("Imagen generada.");
        }} catch (error) {{
            console.error(error);
            setStatus("No se pudo generar la imagen. Intenta de nuevo con el navegador en zoom 100%.");
        }} finally {{
            button.disabled = false;
        }}
    }});
    </script>
    """
    components.html(componente, height=980, scrolling=True)


def color_rgb(hex_color):
    color = hex_color.lstrip("#")
    return tuple(int(color[indice : indice + 2], 16) for indice in (0, 2, 4))


def fuente_slide(tamano, bold=False):
    candidatos = (
        [
            "C:/Windows/Fonts/arialbd.ttf",
            "C:/Windows/Fonts/segoeuib.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        ]
        if bold
        else [
            "C:/Windows/Fonts/arial.ttf",
            "C:/Windows/Fonts/segoeui.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        ]
    )
    for ruta in candidatos:
        try:
            return ImageFont.truetype(ruta, tamano)
        except OSError:
            continue
    return ImageFont.load_default()


def texto_plano_html(valor):
    texto = re.sub(r"<[^>]+>", "", str(valor or ""))
    return html.unescape(texto).strip()


def ancho_texto(draw, texto, fuente):
    bbox = draw.textbbox((0, 0), texto, font=fuente)
    return bbox[2] - bbox[0]


def cortar_texto(draw, texto, fuente, max_width):
    if ancho_texto(draw, texto, fuente) <= max_width:
        return texto
    puntos = "..."
    disponible = max_width - ancho_texto(draw, puntos, fuente)
    if disponible <= 0:
        return puntos
    inicio, fin = 0, len(texto)
    while inicio < fin:
        medio = (inicio + fin + 1) // 2
        if ancho_texto(draw, texto[:medio], fuente) <= disponible:
            inicio = medio
        else:
            fin = medio - 1
    return texto[:inicio].rstrip() + puntos


def envolver_texto(draw, texto, fuente, max_width, max_lines=None):
    palabras = str(texto or "").split()
    lineas = []
    linea = ""
    for palabra in palabras:
        candidata = f"{linea} {palabra}".strip()
        if ancho_texto(draw, candidata, fuente) <= max_width:
            linea = candidata
            continue
        if linea:
            lineas.append(linea)
        if ancho_texto(draw, palabra, fuente) > max_width:
            lineas.append(cortar_texto(draw, palabra, fuente, max_width))
            linea = ""
        else:
            linea = palabra
        if max_lines and len(lineas) >= max_lines:
            break
    if linea and (not max_lines or len(lineas) < max_lines):
        lineas.append(linea)
    if max_lines and len(lineas) > max_lines:
        lineas = lineas[:max_lines]
    if max_lines and len(lineas) == max_lines:
        lineas[-1] = cortar_texto(draw, lineas[-1], fuente, max_width)
    return lineas


def dibujar_texto_envuelto(draw, texto, xy, fuente, fill, max_width, line_gap=6, max_lines=None):
    x, y = xy
    alto_linea = fuente.size + line_gap
    for linea in envolver_texto(draw, texto, fuente, max_width, max_lines):
        draw.text((x, y), linea, font=fuente, fill=fill)
        y += alto_linea
    return y


def ranking_para_slide(panel):
    df = panel["df"]
    valor_columna = panel["valor_columna"]
    if df.empty:
        return pd.DataFrame(columns=[panel["etiqueta_columna"], valor_columna])
    ranking = df.copy()
    ranking[valor_columna] = pd.to_numeric(ranking[valor_columna], errors=TEXT_COERCE).fillna(0)
    return (
        ranking[ranking[valor_columna] > 0]
        .sort_values(by=valor_columna, ascending=False)
        .head(panel.get("top_n", 6))
    )


def dibujar_tarjetas_slide(draw, tarjetas, x, y, width):
    gap = 16
    alto = 150
    cantidad = len(tarjetas)
    ancho = (width - gap * (cantidad - 1)) / cantidad
    borde = color_rgb(UI_PALETTE["border"])
    naranja = color_rgb(UI_PALETTE[TEXT_PRIMARY])
    texto = color_rgb(UI_PALETTE["text"])
    titulo_font = fuente_slide(29, bold=True)
    valor_font = fuente_slide(76, bold=True)
    for indice, (titulo, valor) in enumerate(tarjetas):
        titulo_visible = re.sub(r"^Cumplimiento\s+", "", str(titulo), flags=re.IGNORECASE)
        x0 = int(x + indice * (ancho + gap))
        x1 = int(x0 + ancho)
        draw.rounded_rectangle((x0, y, x1, y + alto), radius=12, fill="white", outline=borde, width=2)
        draw.rounded_rectangle((x0, y, x1, y + 9), radius=8, fill=naranja)
        titulo_lineas = envolver_texto(draw, titulo_visible, titulo_font, int(ancho - 28), max_lines=2)
        titulo_y = y + 31 if len(titulo_lineas) == 1 else y + 20
        for linea in titulo_lineas:
            linea_x = x0 + (ancho - ancho_texto(draw, linea, titulo_font)) / 2
            draw.text((linea_x, titulo_y), linea, font=titulo_font, fill=texto)
            titulo_y += 32
        valor_texto = str(valor)
        valor_x = x0 + (ancho - ancho_texto(draw, valor_texto, valor_font)) / 2
        draw.text((valor_x, y + 78), valor_texto, font=valor_font, fill=naranja)
    return y + alto


def dibujar_ranking_panel(draw, panel, x, y, width, height):
    borde = color_rgb(UI_PALETTE["border"])
    texto = color_rgb(UI_PALETTE["text"])
    muted = color_rgb(UI_PALETTE["muted"])
    naranja = color_rgb(UI_PALETTE[TEXT_PRIMARY])
    mostaza = color_rgb(UI_PALETTE["mustard"])
    pista = (241, 243, 245)
    title_font = fuente_slide(34, bold=True)
    label_font = fuente_slide(30, bold=False)
    value_font = fuente_slide(32, bold=True)
    draw.rounded_rectangle((x, y, x + width, y + height), radius=12, fill="white", outline=borde, width=2)
    draw.text((x + 22, y + 20), str(panel["titulo"]), font=title_font, fill=naranja)
    ranking = ranking_para_slide(panel)
    if ranking.empty:
        draw.text((x + 22, y + 82), "Sin datos para mostrar.", font=label_font, fill=muted)
        return
    maximo = ranking[panel["valor_columna"]].max()
    row_top = y + 96
    row_gap = 18
    row_height = min(74, max(50, int((height - 122) / max(len(ranking), 1))))
    label_w = int(width * 0.43)
    value_w = 76
    bar_x = x + 24 + label_w + 22
    bar_w = width - 48 - label_w - value_w - 36
    for fila_idx, (_, row) in enumerate(ranking.iterrows()):
        y_row = row_top + fila_idx * row_height
        if y_row + 34 > y + height - 12:
            break
        etiqueta = texto_ranking_kpi(valor_limpio(row[panel["etiqueta_columna"]]) or SIN_DATO, panel.get("limite", 62))
        etiqueta = cortar_texto(draw, etiqueta, label_font, label_w)
        valor = float(row[panel["valor_columna"]])
        porcentaje_barra = valor / maximo if maximo else 0
        draw.text((x + 24, y_row), etiqueta, font=label_font, fill=texto)
        track_y = y_row + 9
        draw.rounded_rectangle((bar_x, track_y, bar_x + bar_w, track_y + 24), radius=12, fill=pista)
        draw.rounded_rectangle((bar_x, track_y, bar_x + max(10, int(bar_w * porcentaje_barra)), track_y + 24), radius=12, fill=mostaza)
        valor_texto = numero_ranking_kpi(valor)
        draw.text((x + width - 24 - ancho_texto(draw, valor_texto, value_font), y_row - 2), valor_texto, font=value_font, fill=texto)
        row_top += row_gap


def dibujar_nota_slide(draw, titulo, lineas, x, y, width, height):
    borde = color_rgb(UI_PALETTE["border"])
    texto = color_rgb(UI_PALETTE["text"])
    naranja = color_rgb(UI_PALETTE[TEXT_PRIMARY])
    title_font = fuente_slide(34, bold=True)
    body_font = fuente_slide(29, bold=False)
    draw.rounded_rectangle((x, y, x + width, y + height), radius=12, fill="white", outline=borde, width=2)
    draw.text((x + 22, y + 20), str(titulo), font=title_font, fill=naranja)
    cursor_y = y + 82
    max_width = width - 44
    for linea in lineas:
        texto_linea = texto_plano_html(linea)
        if not texto_linea:
            continue
        cursor_y = dibujar_texto_envuelto(
            draw,
            texto_linea,
            (x + 22, cursor_y),
            body_font,
            texto,
            max_width,
            line_gap=9,
            max_lines=4,
        )
        cursor_y += 16
        if cursor_y > y + height - 32:
            break


def crear_png_slide_kpi(titulo, periodo, tarjetas, caption, ranking_panels, note_lines):
    width, height = 1600, 900
    img = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(img)
    borde = color_rgb(UI_PALETTE["border"])
    texto = color_rgb(UI_PALETTE["text"])
    muted = color_rgb(UI_PALETTE["muted"])
    pad = 30
    draw.rounded_rectangle((2, 2, width - 3, height - 3), radius=14, outline=borde, width=2)
    period_font = fuente_slide(30, bold=False)
    title_font = fuente_slide(54, bold=True)
    caption_font = fuente_slide(30, bold=False)
    periodo_texto = f"{TEXT_PERIODO}{periodo}" if periodo else ""
    draw.text((pad, 20), periodo_texto, font=period_font, fill=muted)
    draw.text((pad, 64), str(titulo), font=title_font, fill=texto)
    cards_bottom = dibujar_tarjetas_slide(draw, tarjetas, pad, 128, width - 2 * pad)
    draw.text((pad, cards_bottom + 16), str(caption), font=caption_font, fill=muted)
    body_y = cards_bottom + 62
    body_h = height - body_y - pad
    gap = 18
    note_w = 470
    left_w = width - 2 * pad - note_w - gap
    if len(ranking_panels) > 1:
        panel_gap = 14
        panel_w = int((left_w - panel_gap) / 2)
        for indice, panel in enumerate(ranking_panels[:2]):
            dibujar_ranking_panel(draw, panel, pad + indice * (panel_w + panel_gap), body_y, panel_w, body_h)
    else:
        dibujar_ranking_panel(draw, ranking_panels[0], pad, body_y, left_w, body_h)
    dibujar_nota_slide(draw, "Lectura", note_lines, pad + left_w + gap, body_y, note_w, body_h)
    buffer = BytesIO()
    img.save(buffer, format="PNG")
    return buffer.getvalue()


def nombre_archivo_slide(titulo, periodo):
    base = normalizar_texto(f"{titulo} {periodo or ''}")
    base = re.sub(r"[^a-z0-9]+", "_", base).strip("_") or "kpi_slide"
    return f"{base}.png"


def render_descarga_slide_png(titulo, periodo, tarjetas, caption, ranking_panels, note_lines):
    png = crear_png_slide_kpi(titulo, periodo, tarjetas, caption, ranking_panels, note_lines)
    st.download_button(
        "Descargar imagen PNG",
        data=png,
        file_name=nombre_archivo_slide(titulo, periodo),
        mime="image/png",
        use_container_width=True,
    )

def grafico_barras_kpi(df, x, y, titulo, color):
    render_ranking_kpi(df, y, x, titulo)


def grafico_porcentaje_tipologias_soporte(resumen):
    if resumen.empty or resumen[TEXT_CANTIDAD].sum() <= 0:
        st.info("No hay datos para graficar las tipologias de casos.")
        return

    datos = resumen.copy()
    datos["Etiqueta %"] = datos["% casos"].apply(lambda valor: f"{valor:g}%")
    fig = px.bar(
        datos,
        x="% casos",
        y=TEXT_TIPOLOGIA_SOPORTE,
        text="Etiqueta %",
        orientation="h",
        hover_data={TEXT_CANTIDAD: True, "% casos": ":.2f", TEXT_DESCRIPCION: True},
        color=TEXT_TIPOLOGIA_SOPORTE,
        color_discrete_sequence=[
            UI_PALETTE[TEXT_PRIMARY],
            UI_PALETTE[TEXT_YELLOW],
            UI_PALETTE[TEXT_LAVENDER],
            UI_PALETTE[TEXT_PURPLE],
        ],
    )
    fig.update_traces(textposition=TEXT_OUTSIDE, cliponaxis=False)
    fig.update_xaxes(range=[0, max(100, float(datos["% casos"].max()) + 8)], ticksuffix="%")
    fig.update_layout(showlegend=False, height=390)
    st.plotly_chart(aplicar_estilo_figura(fig, "Distribucion porcentual por tipologia"), use_container_width=True)


def formato_entero_es(valor):
    try:
        numero = int(round(float(valor)))
    except (TypeError, ValueError):
        numero = 0
    return f"{numero:,}".replace(",", ".")


def formato_porcentaje_es(valor, decimales=1):
    if valor is None or pd.isna(valor):
        valor = 0
    return f"{float(valor):.{decimales}f}%".replace(".", ",")


def html_streamlit(texto):
    return "\n".join(line.lstrip() for line in dedent(texto).strip().splitlines())


def producto_visible_caso(valor):
    return valor_limpio(valor) or "Sin producto"


def cliente_visible_caso(valor):
    return valor_limpio(valor) or SIN_CUENTA


def resumen_clientes_producto(grupo, top_n=3):
    if TEXT_CUENTA not in grupo.columns:
        return pd.Series(
            {
                COL_CLIENTES_DISTINTOS: 0,
                COL_PRINCIPALES_CLIENTES: SIN_DATO,
            }
        )

    clientes = grupo[TEXT_CUENTA].apply(cliente_visible_caso)
    clientes_validos = clientes[clientes != SIN_CUENTA]
    if clientes_validos.empty:
        return pd.Series(
            {
                COL_CLIENTES_DISTINTOS: 0,
                COL_PRINCIPALES_CLIENTES: SIN_DATO,
            }
        )

    conteo_clientes = clientes_validos.value_counts(dropna=False)
    principales = ", ".join(
        f"{cliente} ({formato_entero_es(cantidad)})"
        for cliente, cantidad in conteo_clientes.head(top_n).items()
    )
    return pd.Series(
        {
            COL_CLIENTES_DISTINTOS: int(clientes_validos.nunique()),
            COL_PRINCIPALES_CLIENTES: principales or SIN_DATO,
        }
    )


def resumen_productos_soporte(df, top_n=5):
    columnas = ["Producto", TEXT_CANTIDAD, "% participacion", COL_CLIENTES_DISTINTOS, COL_PRINCIPALES_CLIENTES]
    if df.empty or TEXT_PRODUCTO not in df.columns:
        return pd.DataFrame(columns=columnas)

    trabajo = df.copy()
    trabajo["Producto"] = trabajo[TEXT_PRODUCTO].apply(producto_visible_caso)
    conteo = trabajo["Producto"].value_counts(dropna=False)
    total = int(conteo.sum())
    if total <= 0:
        return pd.DataFrame(columns=columnas)

    filas = []
    productos_top = conteo.head(top_n).index.tolist()
    for producto in productos_top:
        grupo = trabajo[trabajo["Producto"] == producto]
        clientes = resumen_clientes_producto(grupo)
        cantidad = int(len(grupo))
        filas.append(
            {
                "Producto": producto,
                TEXT_CANTIDAD: cantidad,
                "% participacion": round((cantidad / total) * 100, 1),
                COL_CLIENTES_DISTINTOS: int(clientes[COL_CLIENTES_DISTINTOS]),
                COL_PRINCIPALES_CLIENTES: clientes[COL_PRINCIPALES_CLIENTES],
            }
        )

    otros_productos = conteo.iloc[top_n:].index.tolist()
    if otros_productos:
        grupo = trabajo[trabajo["Producto"].isin(otros_productos)]
        clientes = resumen_clientes_producto(grupo)
        otros = int(len(grupo))
        filas.append(
            {
                "Producto": "Otros productos",
                TEXT_CANTIDAD: otros,
                "% participacion": round((otros / total) * 100, 1),
                COL_CLIENTES_DISTINTOS: int(clientes[COL_CLIENTES_DISTINTOS]),
                COL_PRINCIPALES_CLIENTES: clientes[COL_PRINCIPALES_CLIENTES],
            }
        )
    return pd.DataFrame(filas, columns=columnas)


def render_tabla_productos_soporte(resumen):
    estilos = f"""
    <style>
    .product-share-table {{
        width: 100%;
        border: 1px solid {UI_PALETTE["border"]};
        border-radius: 8px;
        overflow: hidden;
        background: #ffffff;
    }}
    .product-share-table table {{
        width: 100%;
        border-collapse: collapse;
        table-layout: fixed;
        font-family: Arial, sans-serif;
    }}
    .product-share-table th {{
        background: #0b2b5c;
        color: #ffffff;
        font-size: 14px;
        font-weight: 800;
        padding: 12px 10px;
        text-align: center;
    }}
    .product-share-table td {{
        border-top: 1px solid {UI_PALETTE["border"]};
        color: {UI_PALETTE["text"]};
        font-size: 15px;
        padding: 13px 10px;
        vertical-align: middle;
    }}
    .product-share-name {{
        display: flex;
        align-items: center;
        gap: 10px;
        word-break: break-word;
    }}
    .product-share-dot {{
        width: 14px;
        min-width: 14px;
        height: 14px;
        border-radius: 999px;
        display: inline-block;
    }}
    .product-share-number {{
        text-align: center;
        font-weight: 700;
        white-space: nowrap;
    }}
    .product-share-clients {{
        color: {UI_PALETTE["text"]};
        font-size: 13px;
        font-weight: 650;
        line-height: 1.25;
        word-break: break-word;
    }}
    .product-share-total td {{
        background: #edf2f7;
        color: #0b1f3a;
        font-weight: 900;
        font-size: 16px;
    }}
    </style>
    """
    filas = []
    total = int(resumen[TEXT_CANTIDAD].sum())
    for indice, row in resumen.iterrows():
        color = PRODUCT_PIE_COLORS[indice % len(PRODUCT_PIE_COLORS)]
        filas.append(
            "<tr>"
            '<td><div class="product-share-name">'
            f'<span class="product-share-dot" style="background:{color};"></span>'
            f"{html.escape(str(row['Producto']))}"
            "</div></td>"
            f'<td class="product-share-number">{formato_entero_es(row[TEXT_CANTIDAD])}</td>'
            f'<td class="product-share-number">{formato_porcentaje_es(row["% participacion"])}</td>'
            f'<td class="product-share-number">{formato_entero_es(row[COL_CLIENTES_DISTINTOS])}</td>'
            f'<td class="product-share-clients">{html.escape(str(row[COL_PRINCIPALES_CLIENTES]))}</td>'
            "</tr>"
        )
    filas.append(
        '<tr class="product-share-total">'
        "<td class=\"product-share-number\">TOTAL</td>"
        f'<td class="product-share-number">{formato_entero_es(total)}</td>'
        '<td class="product-share-number">100,0%</td>'
        '<td class="product-share-number">-</td>'
        '<td class="product-share-clients">-</td>'
        "</tr>"
    )
    tabla = (
        estilos
        + '<div class="product-share-table"><table>'
        "<colgroup>"
        "<col style=\"width:28%\"><col style=\"width:14%\"><col style=\"width:16%\">"
        "<col style=\"width:13%\"><col style=\"width:29%\">"
        "</colgroup>"
        "<thead><tr><th>PRODUCTO</th><th>CANTIDAD</th><th>% PARTICIPACION</th>"
        "<th>CLIENTES</th><th>PRINCIPALES CLIENTES</th></tr></thead>"
        f"<tbody>{''.join(filas)}</tbody>"
        "</table></div>"
    )
    st.markdown(tabla, unsafe_allow_html=True)


def render_distribucion_productos_soporte(df, periodo_label):
    resumen = resumen_productos_soporte(df)
    if resumen.empty:
        st.info("No hay productos para calcular la distribucion de tickets de soporte.")
        return

    total = int(resumen[TEXT_CANTIDAD].sum())
    titulo_periodo = str(periodo_label).upper()
    st.markdown(
        f"""
        <div style="text-align:center; margin: 0 0 6px 0;">
            <div style="font-size:30px; font-weight:900; color:#0b1f3a; line-height:1.1;">
                TICKETS DE SOPORTE - {html.escape(titulo_periodo)}
            </div>
            <div style="font-size:18px; font-weight:700; color:#4b5563; margin-top:4px;">
                DISTRIBUCION POR PRODUCTO Y CLIENTE
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    datos = resumen.copy()
    datos["Etiqueta %"] = datos["% participacion"].apply(formato_porcentaje_es)
    principal_idx = int(datos["% participacion"].to_numpy().argmax())
    datos["Etiqueta torta"] = [
        "" if indice == principal_idx else etiqueta
        for indice, etiqueta in enumerate(datos["Etiqueta %"].tolist())
    ]
    col_grafico, col_tabla = st.columns([1.05, 1])

    with col_grafico:
        fig = px.pie(
            datos,
            values=TEXT_CANTIDAD,
            names="Producto",
            color_discrete_sequence=PRODUCT_PIE_COLORS,
            custom_data=["Etiqueta %", COL_CLIENTES_DISTINTOS, COL_PRINCIPALES_CLIENTES],
        )
        fig.update_traces(
            sort=False,
            direction="clockwise",
            text=datos["Etiqueta torta"],
            texttemplate="%{text}",
            textposition=TEXT_OUTSIDE,
            textfont={"size": 22, "color": "#111827", "family": "Arial, sans-serif"},
            pull=[0 if indice == principal_idx else 0.035 for indice in range(len(datos))],
            automargin=True,
            marker={"line": {"color": "#ffffff", "width": 2}},
            hovertemplate=(
                "<b>%{label}</b><br>Cantidad: %{value}<br>Participacion: %{customdata[0]}"
                "<br>Clientes: %{customdata[1]}<br>Principales clientes: %{customdata[2]}<extra></extra>"
            ),
        )
        principal = float(datos["% participacion"].max())
        if principal >= 45:
            fig.add_annotation(
                text=formato_porcentaje_es(principal),
                x=0.5,
                y=0.48,
                showarrow=False,
                font={"size": 44, "color": "#ffffff", "family": "Arial, sans-serif"},
            )
        fig.update_layout(
            showlegend=False,
            height=520,
            margin={"l": 82, "r": 82, "t": 26, "b": 36},
            paper_bgcolor="rgba(255,255,255,0)",
            font={"color": UI_PALETTE["text"], "size": 18, "family": "Arial, sans-serif"},
            uniformtext={"minsize": 18, "mode": "show"},
        )
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
        st.markdown(
            f"""
            <div style="text-align:center; margin-top:-10px; color:#0b1f3a;">
                <div style="font-weight:900; font-size:18px;">TOTAL: {formato_entero_es(total)}</div>
                <div style="font-size:15px;">Tickets de soporte</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with col_tabla:
        render_tabla_productos_soporte(datos)


def conic_gradient_productos(resumen):
    total = int(resumen[TEXT_CANTIDAD].sum())
    if total <= 0:
        return "#f1f3f5 0% 100%"
    partes = []
    inicio = 0.0
    for indice, row in resumen.iterrows():
        porcentaje_segmento = (float(row[TEXT_CANTIDAD]) / total) * 100
        fin = inicio + porcentaje_segmento
        color = PRODUCT_PIE_COLORS[indice % len(PRODUCT_PIE_COLORS)]
        partes.append(f"{color} {inicio:.4f}% {fin:.4f}%")
        inicio = fin
    return ", ".join(partes)


def slide_product_distribution_html(df, periodo_label):
    resumen = resumen_productos_soporte(df)
    if resumen.empty:
        return (
            '<div class="slide-panel">'
            '<div class="slide-panel-title">Distribucion por producto</div>'
            '<div class="kpi-ranking-empty">Sin productos para mostrar.</div>'
            "</div>"
        )

    total = int(resumen[TEXT_CANTIDAD].sum())
    principal = resumen.iloc[0]
    gradient = conic_gradient_productos(resumen)
    filas = []
    etiquetas = []
    for indice, row in resumen.iterrows():
        color = PRODUCT_PIE_COLORS[indice % len(PRODUCT_PIE_COLORS)]
        porcentaje_texto = formato_porcentaje_es(row["% participacion"])
        etiquetas.append(
            '<div class="slide-product-label">'
            f'<span style="background:{color};"></span>'
            f'<strong>{html.escape(porcentaje_texto)}</strong>'
            f'<em>{html.escape(str(row["Producto"]))}</em>'
            "</div>"
        )
        filas.append(
            "<tr>"
            '<td><div class="slide-product-name">'
            f'<span style="background:{color};"></span>'
            f"{html.escape(str(row['Producto']))}"
            "</div></td>"
            f"<td>{formato_entero_es(row[TEXT_CANTIDAD])}</td>"
            f"<td>{porcentaje_texto}</td>"
            f"<td>{formato_entero_es(row[COL_CLIENTES_DISTINTOS])}</td>"
            f"<td>{html.escape(str(row[COL_PRINCIPALES_CLIENTES]))}</td>"
            "</tr>"
        )
    filas.append(
        '<tr class="slide-product-total">'
        "<td>TOTAL</td>"
        f"<td>{formato_entero_es(total)}</td>"
        "<td>100,0%</td>"
        "<td>-</td>"
        "<td>-</td>"
        "</tr>"
    )
    focos_html = slide_focos_operativos_kpi_html(df)
    return f"""
    <style>
    .kpi-ce-slide {{
        padding: 20px 24px;
        gap: 10px;
    }}
    .kpi-ce-slide .slide-period {{
        font-size: 0.84rem;
        line-height: 1.05;
    }}
    .kpi-ce-slide .slide-title {{
        font-size: 1.42rem;
        line-height: 1.05;
    }}
    .kpi-ce-slide .slide-kpi-grid {{
        gap: 8px;
    }}
    .kpi-ce-slide .slide-kpi-card {{
        min-height: 68px;
        padding: 8px 9px;
        border-top-width: 3px;
    }}
    .kpi-ce-slide .slide-kpi-title {{
        font-size: 0.64rem;
        line-height: 1.06;
    }}
    .kpi-ce-slide .slide-kpi-value {{
        font-size: 1.45rem;
        margin-top: 0.22rem;
    }}
    .kpi-ce-slide .slide-caption {{
        font-size: 0.78rem;
        line-height: 1.08;
    }}
    .kpi-ce-slide .slide-body {{
        grid-template-columns: minmax(0, 1.8fr) minmax(230px, 0.82fr);
        gap: 10px;
    }}
    .kpi-ce-slide .slide-panel {{
        padding: 10px 12px;
    }}
    .kpi-ce-slide .slide-panel-title {{
        font-size: 0.92rem;
        margin-bottom: 0.48rem;
    }}
    .kpi-ce-slide .slide-note {{
        gap: 0.38rem;
        font-size: 0.76rem;
        line-height: 1.18;
    }}
    .slide-product-panel {{
        padding: 14px 16px;
    }}
    .kpi-ce-slide .slide-product-panel {{
        padding: 9px 10px;
    }}
    .slide-product-title {{
        color: #0b1f3a;
        font-size: 1.26rem;
        font-weight: 900;
        line-height: 1.1;
        text-align: center;
        text-transform: uppercase;
    }}
    .kpi-ce-slide .slide-product-title {{
        font-size: 0.96rem;
    }}
    .slide-product-subtitle {{
        color: #4b5563;
        font-size: 0.86rem;
        font-weight: 900;
        line-height: 1.1;
        margin-top: 4px;
        text-align: center;
        text-transform: uppercase;
    }}
    .kpi-ce-slide .slide-product-subtitle {{
        font-size: 0.6rem;
        margin-top: 2px;
    }}
    .slide-product-content {{
        display: grid;
        grid-template-columns: minmax(230px, 0.78fr) minmax(420px, 1.22fr);
        gap: 18px;
        align-items: center;
        margin-top: 12px;
        min-height: 0;
    }}
    .kpi-ce-slide .slide-product-content {{
        grid-template-columns: minmax(145px, 0.55fr) minmax(270px, 1.45fr);
        gap: 10px;
        margin-top: 6px;
        align-items: start;
    }}
    .slide-product-pie-wrap {{
        display: flex;
        flex-direction: column;
        align-items: center;
        min-width: 0;
    }}
    .slide-product-pie {{
        width: min(100%, 292px);
        aspect-ratio: 1 / 1;
        border-radius: 50%;
        background: conic-gradient({gradient});
        border: 3px solid #ffffff;
        box-shadow: 0 8px 18px rgba(20, 20, 20, 0.13);
        display: grid;
        place-items: center;
        position: relative;
    }}
    .kpi-ce-slide .slide-product-pie {{
        width: min(100%, 168px);
    }}
    .slide-product-pie::after {{
        content: "";
        position: absolute;
        inset: 0;
        border-radius: inherit;
        box-shadow: inset 0 0 0 2px rgba(255, 255, 255, 0.78);
        pointer-events: none;
    }}
    .slide-product-center {{
        color: #ffffff;
        font-size: 2.15rem;
        font-weight: 900;
        line-height: 1;
        text-shadow: 0 2px 6px rgba(0, 0, 0, 0.35);
        z-index: 1;
    }}
    .kpi-ce-slide .slide-product-center {{
        font-size: 1.24rem;
    }}
    .slide-product-labels {{
        display: grid;
        grid-template-columns: repeat(2, minmax(0, 1fr));
        gap: 6px 8px;
        margin-top: 12px;
        width: 100%;
    }}
    .kpi-ce-slide .slide-product-labels {{
        display: none;
    }}
    .slide-product-label {{
        display: grid;
        grid-template-columns: 10px auto;
        grid-template-areas:
            "dot pct"
            "dot name";
        column-gap: 6px;
        align-items: center;
        min-width: 0;
    }}
    .slide-product-label span {{
        grid-area: dot;
        width: 10px;
        height: 10px;
        border-radius: 999px;
    }}
    .slide-product-label strong {{
        grid-area: pct;
        color: #0b1f3a;
        font-size: 1rem;
        font-weight: 900;
        line-height: 1;
    }}
    .slide-product-label em {{
        grid-area: name;
        color: #4b5563;
        font-size: 0.68rem;
        font-style: normal;
        font-weight: 800;
        line-height: 1.08;
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
    }}
    .slide-product-total-caption {{
        color: #0b1f3a;
        font-size: 1.05rem;
        font-weight: 900;
        margin-top: 8px;
        text-align: center;
    }}
    .kpi-ce-slide .slide-product-total-caption {{
        font-size: 0.72rem;
        margin-top: 5px;
    }}
    .slide-product-total-caption span {{
        color: #4b5563;
        display: block;
        font-size: 0.78rem;
        font-weight: 700;
        margin-top: 2px;
    }}
    .kpi-ce-slide .slide-product-total-caption span {{
        font-size: 0.54rem;
    }}
    .slide-product-table {{
        border: 1px solid #d9dee6;
        border-radius: 8px;
        overflow: hidden;
        min-width: 0;
    }}
    .slide-product-table table {{
        width: 100%;
        border-collapse: collapse;
        table-layout: fixed;
    }}
    .slide-product-table th {{
        background: #0b2b5c;
        color: #ffffff;
        font-size: 0.76rem;
        font-weight: 900;
        padding: 9px 8px;
        text-align: center;
    }}
    .kpi-ce-slide .slide-product-table th {{
        font-size: 0.5rem;
        padding: 5px 4px;
    }}
    .slide-product-table td {{
        border-top: 1px solid #d9dee6;
        color: #0b1f3a;
        font-size: 0.72rem;
        font-weight: 750;
        padding: 7px 7px;
        text-align: center;
        line-height: 1.15;
        overflow-wrap: anywhere;
    }}
    .kpi-ce-slide .slide-product-table td {{
        font-size: 0.5rem;
        line-height: 1.08;
        padding: 4px 4px;
    }}
    .slide-product-table th:first-child,
    .slide-product-table td:first-child {{
        text-align: left;
    }}
    .slide-product-table th:nth-child(5),
    .slide-product-table td:nth-child(5) {{
        text-align: left;
    }}
    .slide-product-name {{
        display: flex;
        align-items: center;
        gap: 8px;
        min-width: 0;
    }}
    .slide-product-name span {{
        width: 11px;
        min-width: 11px;
        height: 11px;
        border-radius: 999px;
    }}
    .kpi-ce-slide .slide-product-name {{
        gap: 5px;
    }}
    .kpi-ce-slide .slide-product-name span {{
        width: 8px;
        min-width: 8px;
        height: 8px;
    }}
    .slide-product-name {{
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
    }}
    .slide-product-total td {{
        background: #edf2f7;
        color: #0b1f3a;
        font-size: 0.91rem;
        font-weight: 900;
    }}
    .kpi-ce-slide .slide-product-total td {{
        font-size: 0.58rem;
    }}
    .slide-focus-strip {{
        border: 1px solid #d7c5f8;
        border-radius: 8px;
        margin-top: 12px;
        overflow: hidden;
    }}
    .kpi-ce-slide .slide-focus-strip {{
        margin-top: 7px;
    }}
    .slide-focus-strip-title {{
        color: {UI_PALETTE[TEXT_PURPLE]};
        font-size: 0.78rem;
        font-weight: 900;
        line-height: 1;
        padding: 8px 10px 0;
        text-align: center;
        text-transform: uppercase;
    }}
    .kpi-ce-slide .slide-focus-strip-title {{
        font-size: 0.52rem;
        padding-top: 5px;
    }}
    .slide-focus-grid {{
        display: grid;
        grid-template-columns: repeat(4, minmax(0, 1fr));
        gap: 7px;
        padding: 8px 10px 10px;
    }}
    .kpi-ce-slide .slide-focus-grid {{
        gap: 5px;
        padding: 5px 7px 7px;
    }}
    .slide-focus-card {{
        border: 1px solid #d7c5f8;
        border-radius: 8px;
        display: grid;
        grid-template-columns: 34px minmax(0, 1fr);
        gap: 5px;
        align-items: center;
        min-height: 58px;
        padding: 6px;
    }}
    .kpi-ce-slide .slide-focus-card {{
        grid-template-columns: 24px minmax(0, 1fr);
        gap: 4px;
        min-height: 42px;
        padding: 4px;
    }}
    .slide-focus-card-important {{
        border-color: {UI_PALETTE[TEXT_PRIMARY]};
    }}
    .slide-focus-total-card {{
        background: #f2eaf8;
        border-color: #f2eaf8;
        display: block;
        text-align: center;
    }}
    .slide-focus-icon {{
        width: 31px;
        height: 31px;
        border: 1.5px solid #0b3a78;
        border-radius: 999px;
        color: #0b3a78;
        display: grid;
        place-items: center;
        font-size: 0.58rem;
        font-weight: 900;
        line-height: 1;
    }}
    .kpi-ce-slide .slide-focus-icon {{
        width: 23px;
        height: 23px;
        font-size: 0.42rem;
    }}
    .slide-focus-card-important .slide-focus-icon {{
        border-color: {UI_PALETTE[TEXT_PRIMARY]};
        color: {UI_PALETTE[TEXT_PRIMARY]};
    }}
    .slide-focus-label {{
        color: #0b3a78;
        font-size: 0.52rem;
        font-weight: 900;
        line-height: 1.05;
        overflow-wrap: anywhere;
    }}
    .kpi-ce-slide .slide-focus-label {{
        font-size: 0.42rem;
    }}
    .slide-focus-card-important .slide-focus-label {{
        color: {UI_PALETTE[TEXT_PRIMARY]};
    }}
    .slide-focus-value {{
        color: #0b3a78;
        font-size: 1.03rem;
        font-weight: 900;
        line-height: 1;
        margin-top: 2px;
        font-variant-numeric: tabular-nums;
    }}
    .kpi-ce-slide .slide-focus-value {{
        font-size: 0.74rem;
        margin-top: 1px;
    }}
    .slide-focus-total-card .slide-focus-value {{
        color: {UI_PALETTE[TEXT_PURPLE]};
    }}
    .slide-focus-percent {{
        color: {UI_PALETTE[TEXT_PURPLE]};
        font-size: 0.56rem;
        font-weight: 900;
        line-height: 1;
        margin-top: 2px;
    }}
    .kpi-ce-slide .slide-focus-percent {{
        font-size: 0.4rem;
        margin-top: 1px;
    }}
    </style>
    <div class="slide-panel slide-product-panel">
        <div class="slide-product-title">Tickets de soporte - {html.escape(str(periodo_label).upper())}</div>
        <div class="slide-product-subtitle">Distribucion por producto y cliente</div>
        <div class="slide-product-content">
            <div class="slide-product-pie-wrap">
                <div class="slide-product-pie">
                    <div class="slide-product-center">{formato_porcentaje_es(principal["% participacion"])}</div>
                </div>
                <div class="slide-product-labels">{"".join(etiquetas)}</div>
                <div class="slide-product-total-caption">
                    TOTAL: {formato_entero_es(total)}
                    <span>Tickets de soporte</span>
                </div>
            </div>
            <div class="slide-product-table">
                <table>
                    <colgroup>
                        <col style="width:25%">
                        <col style="width:13%">
                        <col style="width:11%">
                        <col style="width:12%">
                        <col style="width:39%">
                    </colgroup>
                    <thead>
                        <tr><th>PRODUCTO</th><th>CANT.</th><th>%</th><th>CLIENTES</th><th>PRINCIPALES CLIENTES</th></tr>
                    </thead>
                    <tbody>{"".join(filas)}</tbody>
                </table>
            </div>
        </div>
        {focos_html}
    </div>
    """


def resumen_otras_tipificaciones(base, top_n=3):
    conteo = base["_tipificacion_kpi"].value_counts(dropna=False)
    otras = conteo.iloc[top_n:]
    total_otras = int(otras.sum())
    if total_otras <= 0:
        return "No hay categorias adicionales fuera del top principal."

    principales_otras = ", ".join(f"{tip} ({cantidad})" for tip, cantidad in otras.head(3).items())
    return f"Otras categorias agrupan {total_otras} casos fuera del top 3: {principales_otras}."


def resumen_tipologia_soporte_principal(base):
    if base.empty or TEXT_TIPOLOGIA_SOPORTE not in base.columns:
        return "No hay tipologia de casos suficiente para resumir."
    conteo = base[TEXT_TIPOLOGIA_SOPORTE].value_counts()
    if conteo.empty:
        return "No hay tipologia de casos suficiente para resumir."
    principal = conteo.index[0]
    cantidad = int(conteo.iloc[0])
    return f"{principal} concentra {cantidad} casos ({porcentaje(cantidad, len(base))}%)."


def resumen_focos_uso_casos(base):
    columnas = ["Foco", TEXT_CANTIDAD, "% casos"]
    if base.empty:
        return pd.DataFrame(columns=columnas)

    texto = base.apply(texto_caso_para_tipologia_soporte, axis=1)
    total = len(base)
    focos = [
        ("Token fisico", CASE_TOKEN_FISICO_READING_TERMS),
        ("Token virtual", CASE_SUPPORT_TOKEN_VIRTUAL_TERMS),
        ("Envio agenda", CASE_AGENDA_DIRECT_TERMS),
        ("Problemas para firmar", CASE_SIGNING_PROBLEM_TERMS),
        ("Acuses", CASE_ACUSES_TERMS),
    ]
    filas = []
    for foco, palabras in focos:
        if foco == "Envio agenda":
            mascara = mascara_envio_agenda_casos(base, texto)
        else:
            mascara = texto.apply(lambda valor: texto_contiene_alguno(valor, palabras))
        cantidad = int(mascara.sum())
        filas.append(
            {
                "Foco": foco,
                TEXT_CANTIDAD: cantidad,
                "% casos": porcentaje(cantidad, total),
            }
        )
    return pd.DataFrame(filas, columns=columnas)


def mascara_envio_agenda_casos(base, texto=None):
    mascara = pd.Series(False, index=base.index)
    if base.empty:
        return mascara

    if TEXT_TIPOLOGIA_SOPORTE in base.columns:
        mascara = mascara | (base[TEXT_TIPOLOGIA_SOPORTE].fillna("") == ENVIO_AGENDA_MANUAL_USO)
    if TEXT_TIPIFICACION_2 in base.columns:
        tipificacion = base[TEXT_TIPIFICACION_2].apply(normalizar_texto)
        mascara = mascara | tipificacion.str.contains("agenda", na=False)
    if texto is not None:
        mascara = mascara | texto.apply(lambda valor: texto_contiene_alguno(valor, CASE_AGENDA_DIRECT_TERMS))
    return mascara.fillna(False)


def resumen_focos_destacados_kpi_casos(base):
    columnas = ["Foco", TEXT_CANTIDAD, "% casos", "Icono", "Destacado"]
    if base.empty:
        return pd.DataFrame(columns=columnas)

    texto = base.apply(texto_caso_para_tipologia_soporte, axis=1)
    total = len(base)
    focos = [
        ("Token fisico", "USB", texto.apply(lambda valor: texto_contiene_alguno(valor, CASE_TOKEN_FISICO_READING_TERMS)), False),
        ("Token virtual", "PC", texto.apply(lambda valor: texto_contiene_alguno(valor, CASE_SUPPORT_TOKEN_VIRTUAL_TERMS)), False),
        ("Envio agenda", "AG", mascara_envio_agenda_casos(base, texto), True),
    ]
    filas = []
    for foco, icono, mascara, destacado in focos:
        cantidad = int(mascara.sum())
        filas.append(
            {
                "Foco": foco,
                TEXT_CANTIDAD: cantidad,
                "% casos": porcentaje(cantidad, total),
                "Icono": icono,
                "Destacado": destacado,
            }
        )
    return pd.DataFrame(filas, columns=columnas)


def focos_operativos_kpi_html(base):
    resumen = resumen_focos_destacados_kpi_casos(base)
    if resumen.empty:
        return ""

    total = len(base)
    tarjetas = []
    for _, row in resumen.iterrows():
        clase_destacada = " kpi-focus-card-important" if row["Destacado"] else ""
        tarjetas.append(
            f'<div class="kpi-focus-card{clase_destacada}">'
            f'<div class="kpi-focus-icon">{html.escape(str(row["Icono"]))}</div>'
            '<div class="kpi-focus-copy">'
            f'<div class="kpi-focus-label">{html.escape(str(row["Foco"]).upper())}</div>'
            f'<div class="kpi-focus-value">{formato_entero_es(row[TEXT_CANTIDAD])}</div>'
            f'<div class="kpi-focus-percent">({formato_porcentaje_es(row["% casos"])})</div>'
            "</div>"
            "</div>"
        )

    return f"""
    <style>
    .kpi-focus-panel {{
        border: 1px solid #d7c5f8;
        border-radius: 8px;
        padding: 14px 16px 0;
        margin-top: 18px;
        background: #ffffff;
    }}
    .kpi-focus-title {{
        color: {UI_PALETTE[TEXT_PURPLE]};
        font-size: 15px;
        font-weight: 900;
        line-height: 1.2;
        margin-bottom: 12px;
        text-align: center;
        text-transform: uppercase;
    }}
    .kpi-focus-grid {{
        display: grid;
        grid-template-columns: repeat(3, minmax(0, 1fr));
        gap: 12px;
    }}
    .kpi-focus-card {{
        border: 1px solid #d7c5f8;
        border-radius: 8px;
        padding: 12px;
        display: grid;
        grid-template-columns: 58px minmax(0, 1fr);
        align-items: center;
        gap: 10px;
        min-height: 92px;
    }}
    .kpi-focus-card-important {{
        border-color: {UI_PALETTE[TEXT_PRIMARY]};
        box-shadow: inset 0 0 0 1px rgba(243, 91, 4, 0.12);
    }}
    .kpi-focus-icon {{
        width: 52px;
        height: 52px;
        border: 2px solid #0b3a78;
        border-radius: 999px;
        color: #0b3a78;
        display: grid;
        place-items: center;
        font-size: 15px;
        font-weight: 900;
        line-height: 1;
    }}
    .kpi-focus-card-important .kpi-focus-icon {{
        border-color: {UI_PALETTE[TEXT_PRIMARY]};
        color: {UI_PALETTE[TEXT_PRIMARY]};
    }}
    .kpi-focus-copy {{
        min-width: 0;
        text-align: center;
    }}
    .kpi-focus-label {{
        color: #0b3a78;
        font-size: 13px;
        font-weight: 900;
        line-height: 1.1;
    }}
    .kpi-focus-card-important .kpi-focus-label {{
        color: {UI_PALETTE[TEXT_PRIMARY]};
    }}
    .kpi-focus-value {{
        color: #0b3a78;
        font-size: 28px;
        font-weight: 900;
        line-height: 1;
        margin-top: 6px;
        font-variant-numeric: tabular-nums;
    }}
    .kpi-focus-percent {{
        color: {UI_PALETTE[TEXT_PURPLE]};
        font-size: 13px;
        font-weight: 900;
        margin-top: 3px;
    }}
    .kpi-focus-total {{
        background: #f2eaf8;
        border-radius: 0 0 8px 8px;
        color: {UI_PALETTE[TEXT_PURPLE]};
        display: grid;
        grid-template-columns: minmax(0, 1fr) auto auto;
        gap: 10px;
        align-items: baseline;
        margin: 14px -16px 0;
        padding: 13px 18px;
    }}
    .kpi-focus-total span {{
        font-size: 15px;
        font-weight: 900;
        text-transform: uppercase;
    }}
    .kpi-focus-total strong {{
        font-size: 34px;
        font-weight: 900;
        line-height: 1;
    }}
    .kpi-focus-total em {{
        font-size: 14px;
        font-style: normal;
        font-weight: 900;
    }}
    @media (max-width: 900px) {{
        .kpi-focus-grid {{
            grid-template-columns: 1fr;
        }}
        .kpi-focus-total {{
            grid-template-columns: 1fr;
            text-align: center;
        }}
    }}
    </style>
    <div class="kpi-focus-panel">
        <div class="kpi-focus-title">Focos operativos sobre el total de tickets</div>
        <div class="kpi-focus-grid">{"".join(tarjetas)}</div>
        <div class="kpi-focus-total">
            <span>Total tickets de soporte</span>
            <strong>{formato_entero_es(total)}</strong>
            <em>(100%)</em>
        </div>
    </div>
    """


def render_focos_operativos_kpi_casos(base):
    contenido = focos_operativos_kpi_html(base)
    if contenido:
        st.markdown(contenido, unsafe_allow_html=True)


def slide_focos_operativos_kpi_html(base):
    resumen = resumen_focos_destacados_kpi_casos(base)
    if resumen.empty:
        return ""

    total = len(base)
    tarjetas = []
    for _, row in resumen.iterrows():
        clase_destacada = " slide-focus-card-important" if row["Destacado"] else ""
        tarjetas.append(
            f'<div class="slide-focus-card{clase_destacada}">'
            f'<div class="slide-focus-icon">{html.escape(str(row["Icono"]))}</div>'
            '<div>'
            f'<div class="slide-focus-label">{html.escape(str(row["Foco"]).upper())}</div>'
            f'<div class="slide-focus-value">{formato_entero_es(row[TEXT_CANTIDAD])}</div>'
            f'<div class="slide-focus-percent">({formato_porcentaje_es(row["% casos"])})</div>'
            "</div>"
            "</div>"
        )
    tarjetas.append(
        '<div class="slide-focus-card slide-focus-total-card">'
        '<div class="slide-focus-label">TOTAL TICKETS</div>'
        f'<div class="slide-focus-value">{formato_entero_es(total)}</div>'
        '<div class="slide-focus-percent">(100%)</div>'
        "</div>"
    )
    return (
        '<div class="slide-focus-strip">'
        '<div class="slide-focus-strip-title">Focos operativos</div>'
        f'<div class="slide-focus-grid">{"".join(tarjetas)}</div>'
        "</div>"
    )


def lectura_focos_uso_casos(base):
    resumen = resumen_focos_uso_casos(base)
    if resumen.empty:
        return "No hay datos suficientes para calcular focos de uso."
    partes = [
        f"{fila['Foco']}: {int(fila[TEXT_CANTIDAD])} ({fila['% casos']}%)"
        for _, fila in resumen.iterrows()
    ]
    return "Focos sobre el total de casos : " + " | ".join(partes) + "."


def lineas_lectura_kpi_casos(metricas, base):
    causa_principal = metricas[COL_PRINCIPAL_CAUSA_COMUN]
    detalle_causa = resumen_detalle_causa_principal(base, causa_principal)
    return [
        (
            "Tipologia principal: "
            f"<strong>{html.escape(str(metricas[COL_PRINCIPAL_SOPORTE]))}</strong>"
        ),
        f'<div class="slide-note-muted">{html.escape(resumen_tipologia_soporte_principal(base))}</div>',
        f'<div class="slide-note-muted">{html.escape(lectura_focos_uso_casos(base))}</div>',
        f"Causa comun: <strong>{html.escape(str(causa_principal))}</strong>",
        f'<div class="slide-note-muted">{html.escape(detalle_causa)}</div>',
        (
            "Tipificacion original dominante: "
            f"<strong>{html.escape(str(metricas[COL_PRINCIPAL_TIPIFICACION]))}</strong>"
        ),
    ]


def render_lectura_kpi(metricas, base):
    lineas = lineas_lectura_kpi_casos(metricas, base)
    contenido = f"""
    <div class="executive-note">
        <div class="executive-note-title">Lectura</div>
        <div class="executive-note-line">{lineas[0]}</div>
        <div class="executive-note-detail">{lineas[1]}</div>
        <div class="executive-note-detail">{lineas[2]}</div>
        <div class="executive-note-line">{lineas[3]}</div>
        <div class="executive-note-detail">{lineas[4]}</div>
        <div class="executive-note-conclusion">{lineas[5]}</div>
    </div>
    """
    st.markdown(contenido, unsafe_allow_html=True)


def render_slide_kpi_casos_cliente_externo(base, metricas, mes_dashboard):
    tarjetas = [
        ("Total casos", metricas["total"]),
        ("Cerrados", metricas["cerrados"]),
        ("Abiertos", metricas["abiertos"]),
        (f"Cumplimiento SLA <={SLA_CASOS_HORAS} h", f"{metricas['cumplimiento_sla']}%"),
    ]
    caption = (
        f"Tiempo promedio: {metricas['promedio']} h | "
        f"Cumplen SLA: {metricas['cumple_sla']} | No cumplen: {metricas['no_cumple_sla']}"
    )
    lineas = lineas_lectura_kpi_casos(metricas, base)
    izquierda = slide_product_distribution_html(base, mes_dashboard or TEXT_TODOS)
    derecha = slide_note_html("Lectura", lineas)
    render_slide_frame_kpi(
        "KPI Casos Cliente Externo",
        mes_dashboard,
        tarjetas,
        caption,
        izquierda,
        derecha,
        frame_class="kpi-ce-slide",
    )


def render_kpi_casos_cliente_externo(df, mes_dashboard=None):
    base, metricas = preparar_kpi_casos_cliente_externo(df)
    if base.empty:
        return

    modo_diapositiva = st.toggle("Formato diapositiva 16:9", key="slide_kpi_casos_cliente_externo")
    if modo_diapositiva:
        render_slide_kpi_casos_cliente_externo(base, metricas, mes_dashboard)
        return

    if mes_dashboard:
        st.caption(f"{TEXT_PERIODO}{mes_dashboard}")
    st.subheader("KPI Casos Cliente Externo")

    render_tarjetas(
        [
            ("Total casos", metricas["total"]),
            ("Cerrados", metricas["cerrados"]),
            ("Abiertos", metricas["abiertos"]),
            (f"Cumplimiento SLA <={SLA_CASOS_HORAS} h", f"{metricas['cumplimiento_sla']}%"),
        ]
    )
    st.caption(
        f"Tiempo promedio: {metricas['promedio']} h | "
        f"Cumplen SLA: {metricas['cumple_sla']} | No cumplen: {metricas['no_cumple_sla']}"
    )

    st.divider()
    render_distribucion_productos_soporte(base, mes_dashboard or TEXT_TODOS)
    render_focos_operativos_kpi_casos(base)

    st.divider()
    col_grafico, col_lectura = st.columns([2.15, 1])
    with col_grafico:
        tipificaciones = resumen_tipologias_soporte_casos(base)
        grafico_porcentaje_tipologias_soporte(tipificaciones)
    with col_lectura:
        render_lectura_kpi(metricas, base)

    with st.expander("Descripcion de las tipologias de casos"):
        st.dataframe(
            pd.DataFrame(CASE_SUPPORT_TYPOLOGY_GUIDE),
            use_container_width=True,
            hide_index=True,
        )

    with st.expander("Detalle completo de casos usados en el calculo"):
        columnas = [
            TEXT_NUMERO,
            TEXT_ESTADO,
            TEXT_CUENTA,
            TEXT_DESCRIPCION_2,
            TEXT_TIPOLOGIA_SOPORTE,
            TEXT_TIPIFICACION_2,
            TEXT_CAUSA_COMUN,
            TEXT_PRODUCTO,
            TEXT_TIEMPO_RESPUESTA,
            "_tiempo_eval_sla_h",
            "Cumple SLA <=36h",
            TEXT_CANAL,
            TEXT_ASIGNADO,
            TEXT_CREADO,
            TEXT_CERRADO,
        ]
        visible = base[[col for col in columnas if col in base.columns]].rename(
            columns={
                "_tiempo_eval_sla_h": "Tiempo evaluado SLA h",
                TEXT_PRODUCTO: "Servicio",
            }
        )
        dataframe_liviano(visible)


def valor_foco_kpi_casos(resumen_focos, foco, columna=TEXT_CANTIDAD):
    filas = resumen_focos[resumen_focos["Foco"] == foco]
    if filas.empty:
        return 0
    return filas.iloc[0].get(columna, 0)


def fila_kpi_casos_rango(df, etiqueta, fecha_inicio, fecha_fin):
    datos = filtrar_rango_dashboard(df, fecha_inicio, fecha_fin)
    base, metricas = preparar_kpi_casos_cliente_externo(datos)
    if base.empty:
        return {
            "Periodo": etiqueta,
            "Rango": etiqueta_rango_fechas(fecha_inicio, fecha_fin),
            TEXT_TOTAL: 0,
            TEXT_CERRADOS: 0,
            TEXT_ABIERTOS: 0,
            "SLA %": 0,
            "Cumple SLA": 0,
            "No cumple SLA": 0,
            COL_PROM_HORAS: 0,
            "Token fisico": 0,
            "Token virtual": 0,
            "Envio agenda": 0,
            "Clientes": 0,
        }

    focos = resumen_focos_destacados_kpi_casos(base)
    clientes = base[TEXT_CUENTA].apply(cliente_visible_caso) if TEXT_CUENTA in base.columns else pd.Series(dtype=TEXT_OBJECT)
    clientes_validos = clientes[clientes != SIN_CUENTA]
    return {
        "Periodo": etiqueta,
        "Rango": etiqueta_rango_fechas(fecha_inicio, fecha_fin),
        TEXT_TOTAL: metricas["total"],
        TEXT_CERRADOS: metricas["cerrados"],
        TEXT_ABIERTOS: metricas["abiertos"],
        "SLA %": metricas["cumplimiento_sla"],
        "Cumple SLA": metricas["cumple_sla"],
        "No cumple SLA": metricas["no_cumple_sla"],
        COL_PROM_HORAS: metricas["promedio"],
        "Token fisico": int(valor_foco_kpi_casos(focos, "Token fisico")),
        "Token virtual": int(valor_foco_kpi_casos(focos, "Token virtual")),
        "Envio agenda": int(valor_foco_kpi_casos(focos, "Envio agenda")),
        "Clientes": int(clientes_validos.nunique()) if not clientes_validos.empty else 0,
    }


def tabla_kpi_casos_comparativo_rangos(df, rangos):
    filas = [
        fila_kpi_casos_rango(df, etiqueta, fecha_inicio, fecha_fin)
        for etiqueta, fecha_inicio, fecha_fin in rangos
    ]
    return pd.DataFrame(filas)


def valor_periodo_kpi_casos(tabla, periodo, columna):
    filas = tabla[tabla["Periodo"] == periodo]
    if filas.empty:
        return 0
    return filas.iloc[0].get(columna, 0)


def etiqueta_mes_periodo_key(periodo):
    anio, mes = parse_mes_periodo(periodo)
    return etiqueta_mes_periodo(anio, mes)


def rango_mes_periodo_key(periodo):
    anio, mes = parse_mes_periodo(periodo)
    inicio = pd.Timestamp(year=anio, month=mes, day=1)
    fin = inicio + pd.offsets.MonthEnd(0)
    return inicio.date(), fin.date()


def selector_meses_kpi_casos_comparativo():
    meses = cargar_meses_disponibles_cache("cases")
    if not meses:
        st.info("No hay meses disponibles para comparar casos.")
        return None

    indice_base = max(len(meses) - 2, 0)
    indice_comparado = len(meses) - 1
    col_base, col_comparado = st.columns(2)
    with col_base:
        mes_base = st.selectbox(
            "Primer mes",
            meses,
            index=indice_base,
            format_func=etiqueta_mes_periodo_key,
            key="kpi_casos_comparativo_mes_base",
        )
    with col_comparado:
        mes_comparado = st.selectbox(
            "Segundo mes",
            meses,
            index=indice_comparado,
            format_func=etiqueta_mes_periodo_key,
            key="kpi_casos_comparativo_mes_comparado",
        )

    if mes_base == mes_comparado:
        st.warning("Selecciona dos meses diferentes.")
        return None
    return mes_base, mes_comparado


def tabla_variacion_kpi_casos(tabla):
    if tabla.empty:
        return pd.DataFrame()
    base = tabla[tabla["Periodo"] == "Base"]
    comparado = tabla[tabla["Periodo"] == "Comparado"]
    if base.empty or comparado.empty:
        return pd.DataFrame()
    base = base.iloc[0]
    comparado = comparado.iloc[0]
    filas = []
    for metrica in [TEXT_TOTAL, TEXT_CERRADOS, TEXT_ABIERTOS, "SLA %", COL_PROM_HORAS, "Token fisico", "Token virtual", "Envio agenda", "Clientes"]:
        filas.append(
            {
                "Metrica": metrica,
                "Base": base.get(metrica, 0),
                "Comparado": comparado.get(metrica, 0),
                "Diferencia": round(float(comparado.get(metrica, 0)) - float(base.get(metrica, 0)), 2),
                "Variacion %": variacion_porcentual(comparado.get(metrica, 0), base.get(metrica, 0)),
            }
        )
    return pd.DataFrame(filas)


def render_tarjetas_kpi_casos_comparativo(tabla, etiqueta_base, etiqueta_comparado):
    total_base = valor_periodo_kpi_casos(tabla, "Base", TEXT_TOTAL)
    total_comparado = valor_periodo_kpi_casos(tabla, "Comparado", TEXT_TOTAL)
    sla_base = valor_periodo_kpi_casos(tabla, "Base", "SLA %")
    sla_comparado = valor_periodo_kpi_casos(tabla, "Comparado", "SLA %")
    render_tarjetas(
        [
            (f"Casos {etiqueta_base}", total_base),
            (f"Casos {etiqueta_comparado}", total_comparado),
            ("Diferencia casos", f"{int(total_comparado - total_base):+d}"),
            (f"SLA {etiqueta_base}", f"{sla_base}%"),
            (f"SLA {etiqueta_comparado}", f"{sla_comparado}%"),
            ("Diferencia SLA p.p.", f"{round(float(sla_comparado) - float(sla_base), 2):+g}"),
        ]
    )


def render_bloque_mes_kpi_casos_comparativo(df, periodo_key, titulo):
    anio, mes = parse_mes_periodo(periodo_key)
    etiqueta = etiqueta_mes_periodo(anio, mes)
    datos = filtrar_anio_mes_dashboard(df, anio, mes)
    st.markdown(f"### {html.escape(titulo)} - {html.escape(etiqueta)}")
    if datos.empty:
        st.info(f"No hay casos cargados para {etiqueta}.")
        return
    render_distribucion_productos_soporte(datos, etiqueta)
    render_focos_operativos_kpi_casos(datos)


def normalizar_productos_para_comparativo(df):
    trabajo = df.copy()
    if TEXT_PRODUCTO not in trabajo.columns:
        trabajo[TEXT_PRODUCTO] = ""
    trabajo["_producto_comparativo"] = trabajo[TEXT_PRODUCTO].apply(producto_visible_caso)
    return trabajo


def resumen_periodo_productos_comparativo(trabajo, productos_top, etiquetas):
    total = int(len(trabajo))
    resumen = {}
    for etiqueta in etiquetas:
        if etiqueta == "Otros productos":
            grupo = trabajo[~trabajo["_producto_comparativo"].isin(productos_top)]
        else:
            grupo = trabajo[trabajo["_producto_comparativo"] == etiqueta]
        cantidad = int(len(grupo))
        clientes = resumen_clientes_producto(grupo) if cantidad else pd.Series(
            {COL_CLIENTES_DISTINTOS: 0, COL_PRINCIPALES_CLIENTES: SIN_DATO}
        )
        resumen[etiqueta] = {
            TEXT_CANTIDAD: cantidad,
            "% participacion": round((cantidad / total) * 100, 1) if total else 0,
            COL_CLIENTES_DISTINTOS: int(clientes[COL_CLIENTES_DISTINTOS]),
            COL_PRINCIPALES_CLIENTES: clientes[COL_PRINCIPALES_CLIENTES],
        }
    return resumen


def tabla_productos_comparativo_soporte(datos_base, datos_comparado, top_n=5):
    columnas = [
        "Producto",
        "Color",
        "Cantidad base",
        "% base",
        "Clientes base",
        "Principales clientes base",
        "Cantidad comparado",
        "% comparado",
        "Clientes comparado",
        "Principales clientes comparado",
        "Diferencia",
        "Diferencia p.p.",
    ]
    base = normalizar_productos_para_comparativo(datos_base)
    comparado = normalizar_productos_para_comparativo(datos_comparado)
    productos = pd.concat(
        [base["_producto_comparativo"], comparado["_producto_comparativo"]],
        ignore_index=True,
    )
    productos = productos[productos.astype(str).str.len() > 0]
    if productos.empty:
        return pd.DataFrame(columns=columnas)

    productos_top = productos.value_counts(dropna=False).head(top_n).index.tolist()
    tiene_otros = bool(
        (~base["_producto_comparativo"].isin(productos_top)).any()
        or (~comparado["_producto_comparativo"].isin(productos_top)).any()
    )
    etiquetas = productos_top + (["Otros productos"] if tiene_otros else [])
    resumen_base = resumen_periodo_productos_comparativo(base, productos_top, etiquetas)
    resumen_comparado = resumen_periodo_productos_comparativo(comparado, productos_top, etiquetas)

    filas = []
    for indice, producto in enumerate(etiquetas):
        base_producto = resumen_base[producto]
        comparado_producto = resumen_comparado[producto]
        filas.append(
            {
                "Producto": producto,
                "Color": PRODUCT_PIE_COLORS[indice % len(PRODUCT_PIE_COLORS)],
                "Cantidad base": base_producto[TEXT_CANTIDAD],
                "% base": base_producto["% participacion"],
                "Clientes base": base_producto[COL_CLIENTES_DISTINTOS],
                "Principales clientes base": base_producto[COL_PRINCIPALES_CLIENTES],
                "Cantidad comparado": comparado_producto[TEXT_CANTIDAD],
                "% comparado": comparado_producto["% participacion"],
                "Clientes comparado": comparado_producto[COL_CLIENTES_DISTINTOS],
                "Principales clientes comparado": comparado_producto[COL_PRINCIPALES_CLIENTES],
                "Diferencia": comparado_producto[TEXT_CANTIDAD] - base_producto[TEXT_CANTIDAD],
                "Diferencia p.p.": round(
                    comparado_producto["% participacion"] - base_producto["% participacion"],
                    1,
                ),
            }
        )
    return pd.DataFrame(filas, columns=columnas)


def conic_gradient_comparativo_productos(comparativo, columna_cantidad):
    total = int(comparativo[columna_cantidad].sum()) if not comparativo.empty else 0
    if total <= 0:
        return "#f1f3f5 0% 100%"
    partes = []
    inicio = 0.0
    for _, row in comparativo.iterrows():
        cantidad = int(row[columna_cantidad])
        if cantidad <= 0:
            continue
        fin = inicio + ((cantidad / total) * 100)
        partes.append(f"{row['Color']} {inicio:.4f}% {fin:.4f}%")
        inicio = fin
    return ", ".join(partes) if partes else "#f1f3f5 0% 100%"


def torta_producto_comparativo_html(comparativo, titulo, columna_cantidad, columna_porcentaje):
    total = int(comparativo[columna_cantidad].sum()) if not comparativo.empty else 0
    principal = float(comparativo[columna_porcentaje].max()) if total and not comparativo.empty else 0
    gradient = conic_gradient_comparativo_productos(comparativo, columna_cantidad)
    etiquetas = []
    for _, row in comparativo.iterrows():
        if int(row[columna_cantidad]) <= 0:
            continue
        etiquetas.append(
            '<div class="kpi-product-compare-label">'
            f'<span style="background:{row["Color"]};"></span>'
            f'<strong>{formato_porcentaje_es(row[columna_porcentaje])}</strong>'
            f'<em>{html.escape(str(row["Producto"]))}</em>'
            "</div>"
        )
    return dedent(f"""
    <div class="kpi-product-compare-pie-card">
        <div class="kpi-product-compare-period">{html.escape(str(titulo))}</div>
        <div class="kpi-product-compare-pie" style="background: conic-gradient({gradient});">
            <div class="kpi-product-compare-center">{formato_porcentaje_es(principal)}</div>
        </div>
        <div class="kpi-product-compare-labels">{"".join(etiquetas)}</div>
        <div class="kpi-product-compare-total">
            TOTAL: {formato_entero_es(total)}
            <span>Tickets de soporte</span>
        </div>
    </div>
    """).strip()


def tabla_productos_comparativo_html(comparativo, etiqueta_base, etiqueta_comparado):
    filas = []
    total_base = int(comparativo["Cantidad base"].sum()) if not comparativo.empty else 0
    total_comparado = int(comparativo["Cantidad comparado"].sum()) if not comparativo.empty else 0
    for _, row in comparativo.iterrows():
        diferencia = int(row["Diferencia"])
        clase_diferencia = "positive" if diferencia > 0 else ("negative" if diferencia < 0 else "neutral")
        clientes = f"{formato_entero_es(row['Clientes base'])} -> {formato_entero_es(row['Clientes comparado'])}"
        principales = row["Principales clientes comparado"]
        if principales == SIN_DATO:
            principales = row["Principales clientes base"]
        filas.append(
            "<tr>"
            '<td><div class="kpi-product-compare-name">'
            f'<span style="background:{row["Color"]};"></span>'
            f"{html.escape(str(row['Producto']))}"
            "</div></td>"
            '<td class="kpi-product-compare-number">'
            f"{formato_entero_es(row['Cantidad base'])}<small>{formato_porcentaje_es(row['% base'])}</small>"
            "</td>"
            '<td class="kpi-product-compare-number">'
            f"{formato_entero_es(row['Cantidad comparado'])}<small>{formato_porcentaje_es(row['% comparado'])}</small>"
            "</td>"
            f'<td class="kpi-product-compare-diff {clase_diferencia}">'
            f"{diferencia:+d}<small>{row['Diferencia p.p.']:+.1f} p.p.</small>"
            "</td>"
            f'<td class="kpi-product-compare-number">{html.escape(clientes)}</td>'
            f'<td class="kpi-product-compare-clients">{html.escape(str(principales))}</td>'
            "</tr>"
        )
    filas.append(
        '<tr class="kpi-product-compare-total-row">'
        "<td>TOTAL</td>"
        f"<td>{formato_entero_es(total_base)}<small>100,0%</small></td>"
        f"<td>{formato_entero_es(total_comparado)}<small>100,0%</small></td>"
        f"<td>{total_comparado - total_base:+d}</td>"
        "<td>-</td><td>-</td>"
        "</tr>"
    )
    return dedent(f"""
    <div class="kpi-product-compare-table">
        <table>
            <colgroup>
                <col style="width:24%">
                <col style="width:13%">
                <col style="width:13%">
                <col style="width:12%">
                <col style="width:12%">
                <col style="width:26%">
            </colgroup>
            <thead>
                <tr>
                    <th>PRODUCTO</th>
                    <th>{html.escape(str(etiqueta_base).upper())}</th>
                    <th>{html.escape(str(etiqueta_comparado).upper())}</th>
                    <th>DIF.</th>
                    <th>CLIENTES</th>
                    <th>PRINCIPALES CLIENTES</th>
                </tr>
            </thead>
            <tbody>{"".join(filas)}</tbody>
        </table>
    </div>
    """).strip()


def nombre_mes_corto(etiqueta_periodo):
    partes = str(etiqueta_periodo).split()
    return partes[0] if partes else str(etiqueta_periodo)


def focos_comparativo_html(datos_base, datos_comparado, etiqueta_base, etiqueta_comparado):
    focos_base = resumen_focos_destacados_kpi_casos(datos_base)
    focos_comparado = resumen_focos_destacados_kpi_casos(datos_comparado)
    if focos_base.empty and focos_comparado.empty:
        return ""

    mes_base_corto = nombre_mes_corto(etiqueta_base)
    mes_comparado_corto = nombre_mes_corto(etiqueta_comparado)
    base_por_foco = {row["Foco"]: row for _, row in focos_base.iterrows()}
    comparado_por_foco = {row["Foco"]: row for _, row in focos_comparado.iterrows()}
    orden = ["Token fisico", "Token virtual", "Envio agenda"]
    tarjetas = []
    for foco in orden:
        base = base_por_foco.get(foco, {})
        comparado = comparado_por_foco.get(foco, {})
        icono = base.get("Icono") or comparado.get("Icono") or ""
        destacado = bool(base.get("Destacado", False) or comparado.get("Destacado", False))
        cantidad_base = int(base.get(TEXT_CANTIDAD, 0))
        cantidad_comparado = int(comparado.get(TEXT_CANTIDAD, 0))
        porcentaje_base = float(base.get("% casos", 0))
        porcentaje_comparado = float(comparado.get("% casos", 0))
        clase = " important" if destacado else ""
        tarjetas.append(
            f'<div class="kpi-product-compare-focus-card{clase}">'
            f'<div class="kpi-product-compare-focus-icon">{html.escape(str(icono))}</div>'
            '<div class="kpi-product-compare-focus-body">'
            f'<div class="kpi-product-compare-focus-label">{html.escape(str(foco).upper())}</div>'
            '<div class="kpi-product-compare-focus-values">'
            f'<span><em>{html.escape(mes_base_corto)}</em><strong>{formato_entero_es(cantidad_base)}</strong>'
            f'<small>{formato_porcentaje_es(porcentaje_base)}</small></span>'
            f'<span><em>{html.escape(mes_comparado_corto)}</em><strong>{formato_entero_es(cantidad_comparado)}</strong>'
            f'<small>{formato_porcentaje_es(porcentaje_comparado)}</small></span>'
            "</div>"
            f'<div class="kpi-product-compare-focus-diff">{cantidad_comparado - cantidad_base:+d}</div>'
            "</div></div>"
        )

    total_base = len(datos_base)
    total_comparado = len(datos_comparado)
    tarjetas.append(
        '<div class="kpi-product-compare-focus-card total">'
        '<div class="kpi-product-compare-focus-label">TOTAL TICKETS</div>'
        '<div class="kpi-product-compare-focus-values">'
        f'<span><em>{html.escape(mes_base_corto)}</em><strong>{formato_entero_es(total_base)}</strong><small>100,0%</small></span>'
        f'<span><em>{html.escape(mes_comparado_corto)}</em><strong>{formato_entero_es(total_comparado)}</strong><small>100,0%</small></span>'
        "</div>"
        f'<div class="kpi-product-compare-focus-diff">{total_comparado - total_base:+d}</div>'
        "</div>"
    )
    return (
        '<div class="kpi-product-compare-focus-strip">'
        '<div class="kpi-product-compare-focus-title">Focos operativos sobre el total de tickets</div>'
        f'<div class="kpi-product-compare-focus-grid">{"".join(tarjetas)}</div>'
        "</div>"
    )


def render_comparativo_visual_meses_kpi_casos(df, mes_base, mes_comparado):
    anio_base, mes_base_num = parse_mes_periodo(mes_base)
    anio_comparado, mes_comparado_num = parse_mes_periodo(mes_comparado)
    etiqueta_base = etiqueta_mes_periodo(anio_base, mes_base_num)
    etiqueta_comparado = etiqueta_mes_periodo(anio_comparado, mes_comparado_num)
    datos_base = filtrar_anio_mes_dashboard(df, anio_base, mes_base_num)
    datos_comparado = filtrar_anio_mes_dashboard(df, anio_comparado, mes_comparado_num)
    comparativo = tabla_productos_comparativo_soporte(datos_base, datos_comparado)
    if comparativo.empty:
        st.info("No hay productos para calcular la distribucion de tickets de soporte por mes.")
        return

    torta_base = torta_producto_comparativo_html(comparativo, etiqueta_base, "Cantidad base", "% base")
    torta_comparado = torta_producto_comparativo_html(
        comparativo,
        etiqueta_comparado,
        "Cantidad comparado",
        "% comparado",
    )
    tabla_html = tabla_productos_comparativo_html(comparativo, etiqueta_base, etiqueta_comparado)
    focos_html = focos_comparativo_html(datos_base, datos_comparado, etiqueta_base, etiqueta_comparado)
    st.markdown(
        html_streamlit(f"""
        <style>
        .kpi-product-compare-panel {{
            background: #ffffff;
            border: 2px solid {UI_PALETTE["border"]};
            border-radius: 8px;
            box-shadow: 0 14px 34px rgba(20, 20, 20, 0.08);
            padding: 26px 30px 28px;
            width: 100%;
            overflow: hidden;
        }}
        .kpi-product-compare-title {{
            color: #0b1f3a;
            font-size: 2.45rem;
            font-weight: 900;
            line-height: 1.05;
            text-align: center;
            text-transform: uppercase;
        }}
        .kpi-product-compare-subtitle {{
            color: #4b5563;
            font-size: 1.35rem;
            font-weight: 900;
            margin-top: 8px;
            text-align: center;
            text-transform: uppercase;
        }}
        .kpi-product-compare-pies {{
            display: grid;
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: 34px;
            margin-top: 28px;
            align-items: start;
        }}
        .kpi-product-compare-pie-card {{
            display: flex;
            flex-direction: column;
            align-items: center;
            min-width: 0;
        }}
        .kpi-product-compare-period {{
            color: #0b1f3a;
            font-size: 1.35rem;
            font-weight: 900;
            margin-bottom: 14px;
            text-align: center;
        }}
        .kpi-product-compare-pie {{
            width: min(100%, 350px);
            aspect-ratio: 1 / 1;
            border: 4px solid #ffffff;
            border-radius: 50%;
            box-shadow: 0 14px 30px rgba(20, 20, 20, 0.18);
            display: grid;
            place-items: center;
            position: relative;
        }}
        .kpi-product-compare-pie::after {{
            content: "";
            position: absolute;
            inset: 0;
            border-radius: inherit;
            box-shadow: inset 0 0 0 2px rgba(255, 255, 255, 0.78);
            pointer-events: none;
        }}
        .kpi-product-compare-center {{
            color: #ffffff;
            font-size: 3.25rem;
            font-weight: 900;
            line-height: 1;
            text-shadow: 0 2px 6px rgba(0, 0, 0, 0.35);
            z-index: 1;
        }}
        .kpi-product-compare-total {{
            color: #0b1f3a;
            font-size: 1.35rem;
            font-weight: 900;
            margin-top: 12px;
            text-align: center;
        }}
        .kpi-product-compare-labels {{
            display: grid;
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: 10px 16px;
            margin-top: 16px;
            width: min(100%, 430px);
        }}
        .kpi-product-compare-label {{
            display: grid;
            grid-template-columns: 15px minmax(0, 1fr);
            grid-template-areas:
                "dot pct"
                "dot name";
            column-gap: 9px;
            min-width: 0;
        }}
        .kpi-product-compare-label span {{
            grid-area: dot;
            width: 15px;
            height: 15px;
            border-radius: 999px;
            margin-top: 3px;
        }}
        .kpi-product-compare-label strong {{
            grid-area: pct;
            color: #0b1f3a;
            font-size: 1.05rem;
            font-weight: 900;
            line-height: 1;
        }}
        .kpi-product-compare-label em {{
            grid-area: name;
            color: #4b5563;
            font-size: 0.88rem;
            font-style: normal;
            font-weight: 800;
            line-height: 1.08;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }}
        .kpi-product-compare-total span {{
            color: #4b5563;
            display: block;
            font-size: 1rem;
            font-weight: 700;
            margin-top: 4px;
        }}
        .kpi-product-compare-table {{
            border: 1px solid #d9dee6;
            border-radius: 8px;
            margin-top: 26px;
            overflow: hidden;
        }}
        .kpi-product-compare-table table {{
            width: 100%;
            border-collapse: collapse;
            table-layout: fixed;
        }}
        .kpi-product-compare-table th {{
            background: #0b2b5c;
            color: #ffffff;
            font-size: 1.02rem;
            font-weight: 900;
            padding: 13px 10px;
            text-align: center;
        }}
        .kpi-product-compare-table td {{
            border-top: 1px solid #d9dee6;
            color: #0b1f3a;
            font-size: 1.02rem;
            font-weight: 750;
            line-height: 1.2;
            padding: 12px 10px;
            text-align: center;
            vertical-align: middle;
            overflow-wrap: anywhere;
        }}
        .kpi-product-compare-table th:first-child,
        .kpi-product-compare-table td:first-child,
        .kpi-product-compare-table th:last-child,
        .kpi-product-compare-table td:last-child {{
            text-align: left;
        }}
        .kpi-product-compare-name {{
            display: flex;
            align-items: center;
            gap: 10px;
            min-width: 0;
        }}
        .kpi-product-compare-name span {{
            width: 14px;
            min-width: 14px;
            height: 14px;
            border-radius: 999px;
        }}
        .kpi-product-compare-number,
        .kpi-product-compare-diff {{
            font-variant-numeric: tabular-nums;
        }}
        .kpi-product-compare-number small,
        .kpi-product-compare-diff small,
        .kpi-product-compare-total-row small {{
            color: #4b5563;
            display: block;
            font-size: 0.8rem;
            font-weight: 900;
            margin-top: 4px;
        }}
        .kpi-product-compare-diff.positive {{
            color: {UI_PALETTE[TEXT_PRIMARY]};
        }}
        .kpi-product-compare-diff.negative {{
            color: #0b3a78;
        }}
        .kpi-product-compare-clients {{
            font-size: 0.86rem;
            font-weight: 800;
            line-height: 1.18;
        }}
        .kpi-product-compare-total-row td {{
            background: #edf2f7;
            color: #0b1f3a;
            font-size: 1.1rem;
            font-weight: 900;
        }}
        .kpi-product-compare-focus-strip {{
            border: 1px solid #d7c5f8;
            border-radius: 8px;
            margin-top: 18px;
            overflow: hidden;
        }}
        .kpi-product-compare-focus-title {{
            color: {UI_PALETTE[TEXT_PURPLE]};
            font-size: 1.02rem;
            font-weight: 900;
            line-height: 1;
            padding: 12px 12px 0;
            text-align: center;
            text-transform: uppercase;
        }}
        .kpi-product-compare-focus-grid {{
            display: grid;
            grid-template-columns: repeat(4, minmax(0, 1fr));
            gap: 10px;
            padding: 12px 14px 14px;
        }}
        .kpi-product-compare-focus-card {{
            border: 1px solid #d7c5f8;
            border-radius: 8px;
            display: grid;
            grid-template-columns: 46px minmax(0, 1fr);
            gap: 9px;
            align-items: center;
            min-height: 110px;
            padding: 10px;
        }}
        .kpi-product-compare-focus-card.important {{
            border-color: {UI_PALETTE[TEXT_PRIMARY]};
        }}
        .kpi-product-compare-focus-card.total {{
            background: #f2eaf8;
            border-color: #f2eaf8;
            display: block;
            text-align: center;
        }}
        .kpi-product-compare-focus-icon {{
            width: 42px;
            height: 42px;
            border: 1.5px solid #0b3a78;
            border-radius: 999px;
            color: #0b3a78;
            display: grid;
            place-items: center;
            font-size: 0.72rem;
            font-weight: 900;
            line-height: 1;
        }}
        .kpi-product-compare-focus-card.important .kpi-product-compare-focus-icon {{
            border-color: {UI_PALETTE[TEXT_PRIMARY]};
            color: {UI_PALETTE[TEXT_PRIMARY]};
        }}
        .kpi-product-compare-focus-label {{
            color: #0b3a78;
            font-size: 0.74rem;
            font-weight: 900;
            line-height: 1.05;
        }}
        .kpi-product-compare-focus-values {{
            display: grid;
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: 5px;
            margin-top: 4px;
        }}
        .kpi-product-compare-focus-values span {{
            min-width: 0;
        }}
        .kpi-product-compare-focus-values em {{
            color: #4b5563;
            display: block;
            font-size: 0.62rem;
            font-style: normal;
            font-weight: 900;
            line-height: 1;
            text-transform: uppercase;
        }}
        .kpi-product-compare-focus-values strong {{
            color: #0b3a78;
            display: block;
            font-size: 1.35rem;
            font-weight: 900;
            line-height: 1;
            margin-top: 2px;
        }}
        .kpi-product-compare-focus-values small {{
            color: {UI_PALETTE[TEXT_PURPLE]};
            display: block;
            font-size: 0.62rem;
            font-weight: 900;
            line-height: 1;
            margin-top: 2px;
        }}
        .kpi-product-compare-focus-diff {{
            color: {UI_PALETTE[TEXT_PRIMARY]};
            font-size: 0.95rem;
            font-weight: 900;
            margin-top: 6px;
        }}
        @media (max-width: 900px) {{
            .kpi-product-compare-pies,
            .kpi-product-compare-focus-grid {{
                grid-template-columns: 1fr;
            }}
            .kpi-product-compare-table {{
                overflow-x: auto;
            }}
            .kpi-product-compare-table table {{
                min-width: 1120px;
            }}
        }}
        </style>
        <div class="kpi-product-compare-panel">
            <div class="kpi-product-compare-title">
                Tickets de soporte - {html.escape(str(etiqueta_base))} / {html.escape(str(etiqueta_comparado))}
            </div>
            <div class="kpi-product-compare-subtitle">
                Distribucion por producto y cliente
            </div>
            <div class="kpi-product-compare-pies">
                {torta_base}
                {torta_comparado}
            </div>
            {tabla_html}
            {focos_html}
        </div>
        """),
        unsafe_allow_html=True,
    )


def render_kpi_casos_cliente_externo_comparativo():
    st.subheader("KPI Casos Cliente Externo por mes")
    with st.spinner("Cargando casos por mes..."):
        df = cargar_casos_soporte_cache()
    if df.empty:
        st.info("No hay casos cargados para los meses seleccionados.")
        return

    df = preparar_fechas_dashboard(df)
    meses = selector_meses_kpi_casos_comparativo()
    if not meses:
        return

    mes_base, mes_comparado = meses
    etiqueta_base = etiqueta_mes_periodo_key(mes_base)
    etiqueta_comparado = etiqueta_mes_periodo_key(mes_comparado)
    base_inicio, base_fin = rango_mes_periodo_key(mes_base)
    comparado_inicio, comparado_fin = rango_mes_periodo_key(mes_comparado)
    rangos = [
        ("Base", base_inicio, base_fin),
        ("Comparado", comparado_inicio, comparado_fin),
    ]

    tabla = tabla_kpi_casos_comparativo_rangos(df, rangos)
    render_tarjetas_kpi_casos_comparativo(tabla, etiqueta_base, etiqueta_comparado)
    st.caption(f"{etiqueta_base} | {etiqueta_comparado}")

    st.divider()
    render_comparativo_visual_meses_kpi_casos(df, mes_base, mes_comparado)

    st.divider()
    st.subheader("Resumen por mes")
    variacion = tabla_variacion_kpi_casos(tabla)
    if not variacion.empty:
        variacion["Variacion %"] = variacion["Variacion %"].apply(
            lambda valor: f"Sin {etiqueta_base}" if pd.isna(valor) else formato_porcentaje_presentacion(valor)
        )
        variacion = variacion.rename(columns={"Base": etiqueta_base, "Comparado": etiqueta_comparado})
    st.dataframe(variacion, use_container_width=True, hide_index=True)

    with st.expander("Ver metricas por mes"):
        tabla_visible = tabla.copy()
        tabla_visible["Periodo"] = tabla_visible["Periodo"].replace(
            {"Base": etiqueta_base, "Comparado": etiqueta_comparado}
        )
        st.dataframe(tabla_visible, use_container_width=True, hide_index=True)


def segmento_incidente(valor):
    if valor == TIPIFICACION_INCIDENTE_CLIENTE_EXTERNO:
        return "Cliente externo"
    if valor == TIPIFICACION_INCIDENTE_INTERNO:
        return "Cliente interno"
    return "Otro"


VALORES_REINCIDENCIA_NO_INFORMATIVOS = {
    "",
    "na",
    "n/a",
    "no aplica",
    "otro",
    "otros",
    "pendiente",
    "sin dato",
    "sin inferencia",
    "sin informacion",
    "sin patron concluyente en descripcion y anotaciones",
}


def componente_reincidencia(valor):
    texto = normalizar_texto(valor)
    if texto in VALORES_REINCIDENCIA_NO_INFORMATIVOS:
        return ""
    return texto


def valor_visible_reincidencia(valor):
    texto = valor_limpio(valor)
    return texto if componente_reincidencia(texto) else ""


def firma_reincidencia_incidente(row):
    segmento = componente_reincidencia(row.get(TEXT_SEGMENTO))
    servicio = componente_reincidencia(row.get(TEXT_SERVICIO_NEGOCIO))
    causa = componente_reincidencia(row.get(TEXT_CAUSA_RAIZ_AUTO))
    tipo_falla = componente_reincidencia(row.get(TEXT_TIPO_FALLA))
    base = [f"segmento:{segmento}"] if segmento else []

    if servicio and causa:
        return "|".join(base + [f"servicio:{servicio}", f"causa:{causa}"])
    if servicio and tipo_falla:
        return "|".join(base + [f"servicio:{servicio}", f"tipo_falla:{tipo_falla}"])
    if causa and tipo_falla:
        return "|".join(base + [f"causa:{causa}", f"tipo_falla:{tipo_falla}"])
    return ""


def etiqueta_firma_reincidencia_incidente(row):
    servicio = valor_visible_reincidencia(row.get(TEXT_SERVICIO_NEGOCIO))
    causa = valor_visible_reincidencia(row.get(TEXT_CAUSA_RAIZ_AUTO))
    tipo_falla = valor_visible_reincidencia(row.get(TEXT_TIPO_FALLA))

    if servicio and causa:
        return f"Servicio: {servicio} | Causa: {causa}"
    if servicio and tipo_falla:
        return f"Servicio: {servicio} | Tipo falla: {tipo_falla}"
    if causa and tipo_falla:
        return f"Causa: {causa} | Tipo falla: {tipo_falla}"
    return "Sin firma tecnica suficiente"


def serie_fecha_normalizada(df, columna):
    if columna not in df.columns:
        return pd.Series(pd.NaT, index=df.index, dtype=PANDAS_DATETIME_DTYPE)
    return pd.to_datetime(df[columna].apply(normalizar_fecha), errors=TEXT_COERCE)


def marcar_reincidencias_por_firma(trabajo):
    marcas = pd.Series(False, index=trabajo.index)
    if trabajo.empty or TEXT_FIRMA_REINCIDENCIA not in trabajo.columns:
        return marcas

    firmas_validas = trabajo[TEXT_FIRMA_REINCIDENCIA].fillna("").ne("")
    for _, grupo in trabajo[firmas_validas].groupby(TEXT_FIRMA_REINCIDENCIA, sort=False):
        if len(grupo) < 2:
            continue

        grupo_ordenado = grupo.sort_values(
            by=[TEXT_CREADO_DT_2, TEXT_CERRADO_DT, TEXT_NUMERO],
            na_position="last",
        )
        tiene_fechas = (
            grupo_ordenado[TEXT_CREADO_DT_2].notna().any()
            and grupo_ordenado[TEXT_CERRADO_DT].notna().any()
        )
        if not tiene_fechas:
            marcas.loc[grupo_ordenado.index[1:]] = True
            continue

        cierres_previos = []
        for indice, fila in grupo_ordenado.iterrows():
            creado = fila.get(TEXT_CREADO_DT_2)
            if pd.notna(creado) and any(cierre <= creado for cierre in cierres_previos):
                marcas.at[indice] = True

            cierre = fila.get(TEXT_CERRADO_DT)
            if pd.notna(cierre):
                cierres_previos.append(cierre)

    return marcas


def agregar_reincidencia_incidentes(trabajo):
    trabajo = trabajo.copy()
    if trabajo.empty:
        trabajo[TEXT_CREADO_DT_2] = pd.Series(dtype=PANDAS_DATETIME_DTYPE)
        trabajo[TEXT_CERRADO_DT] = pd.Series(dtype=PANDAS_DATETIME_DTYPE)
        trabajo[TEXT_FIRMA_REINCIDENCIA] = pd.Series(dtype=TEXT_OBJECT)
        trabajo[TEXT_FIRMA_REINCIDENCIA_LABEL] = pd.Series(dtype=TEXT_OBJECT)
        trabajo[TEXT_REINCIDENTE] = pd.Series(dtype="bool")
        return trabajo

    trabajo[TEXT_CREADO_DT_2] = serie_fecha_normalizada(trabajo, TEXT_CREADO)
    trabajo[TEXT_CERRADO_DT] = serie_fecha_normalizada(trabajo, TEXT_CERRADO)
    trabajo[TEXT_FIRMA_REINCIDENCIA] = trabajo.apply(firma_reincidencia_incidente, axis=1)
    trabajo[TEXT_FIRMA_REINCIDENCIA_LABEL] = trabajo.apply(etiqueta_firma_reincidencia_incidente, axis=1)
    trabajo[TEXT_REINCIDENTE] = marcar_reincidencias_por_firma(trabajo)
    trabajo[COL_REINCIDENTE] = trabajo[TEXT_REINCIDENTE].map({True: "Si", False: "No"})
    trabajo[COL_FIRMA_REINCIDENCIA] = trabajo[TEXT_FIRMA_REINCIDENCIA_LABEL]
    return trabajo


def preparar_kpi_incidentes(df):
    trabajo = agregar_campos_sla_incidentes(df)
    if trabajo.empty:
        return trabajo, {}

    trabajo = trabajo[
        trabajo[TEXT_TIPIFICACION_AUTO].fillna("").isin(
            [TIPIFICACION_INCIDENTE_CLIENTE_EXTERNO, TIPIFICACION_INCIDENTE_INTERNO]
        )
    ].copy()
    if trabajo.empty:
        return trabajo, {}

    trabajo[TEXT_CERRADO_2] = mascara_cerrados(trabajo)
    trabajo[TEXT_ABIERTO] = ~trabajo[TEXT_CERRADO_2]
    trabajo[TEXT_SEGMENTO] = trabajo[TEXT_TIPIFICACION_AUTO].apply(segmento_incidente)
    trabajo[TEXT_DURACION_HORAS_NUM] = pd.to_numeric(trabajo[TEXT_DURACION_SLA_HORAS], errors=TEXT_COERCE)
    trabajo = agregar_reincidencia_incidentes(trabajo)

    cerrados_sla = trabajo[
        trabajo[TEXT_CERRADO_2]
        & trabajo[TEXT_ESTADO_SLA].isin([ESTADO_SLA_CUMPLE, ESTADO_SLA_NO_CUMPLE])
    ].copy()
    cumplen = int((cerrados_sla[TEXT_ESTADO_SLA] == ESTADO_SLA_CUMPLE).sum())
    no_cumplen = int((cerrados_sla[TEXT_ESTADO_SLA] == ESTADO_SLA_NO_CUMPLE).sum())
    reincidentes = int(trabajo[TEXT_REINCIDENTE].sum())
    externos_mask = trabajo[TEXT_SEGMENTO] == "Cliente externo"
    internos_mask = trabajo[TEXT_SEGMENTO] == "Cliente interno"

    metricas = {
        "total": len(trabajo),
        "externos": int(externos_mask.sum()),
        "internos": int(internos_mask.sum()),
        "abiertos": int(trabajo[TEXT_ABIERTO].sum()),
        "cerrados": int(trabajo[TEXT_CERRADO_2].sum()),
        "abiertos_externos": int((externos_mask & trabajo[TEXT_ABIERTO]).sum()),
        "abiertos_internos": int((internos_mask & trabajo[TEXT_ABIERTO]).sum()),
        "cerrados_externos": int((externos_mask & trabajo[TEXT_CERRADO_2]).sum()),
        "cerrados_internos": int((internos_mask & trabajo[TEXT_CERRADO_2]).sum()),
        "reincidentes": reincidentes,
        "tasa_reincidencia": porcentaje(reincidentes, len(trabajo)),
        "cumplimiento_sla": porcentaje(cumplen, cumplen + no_cumplen),
        "cumple_sla": cumplen,
        "no_cumple_sla": no_cumplen,
        "promedio": round(trabajo[TEXT_DURACION_HORAS_NUM].dropna().mean(), 2)
        if trabajo[TEXT_DURACION_HORAS_NUM].notna().any()
        else 0,
    }
    return trabajo, metricas


def resumen_causas_kpi_incidentes(base):
    if base.empty:
        return pd.DataFrame(
            columns=[
                TEXT_SEGMENTO,
                COL_CAUSA_RAIZ,
                COL_FAMILIA_INCIDENTE,
                TEXT_CANTIDAD,
                "% segmento",
                COL_LECTURA_EJECUTIVA,
                COL_EVIDENCIA_INCIDENTE,
                COL_ACCION_SUGERIDA,
                "Detalle tecnico observado",
            ]
        )

    resumenes = []
    for segmento, grupo in base.groupby(TEXT_SEGMENTO):
        causas = resumen_causas_incidentes(grupo, "% segmento")
        if not causas.empty:
            causas[COL_LECTURA_EJECUTIVA] = causas.apply(lectura_causa_kpi_incidente, axis=1)
        causas.insert(0, TEXT_SEGMENTO, segmento)
        resumenes.append(causas)
    return pd.concat(resumenes, ignore_index=True) if resumenes else pd.DataFrame()


def resumen_reincidencia_kpi_incidentes(base):
    columnas = [
        COL_FIRMA_REINCIDENCIA,
        TEXT_CANTIDAD,
        COL_INCIDENTES_REINCIDENTES,
        "Tasa reincidencia %",
        "Primer incidente",
        "Ultimo incidente",
    ]
    if base.empty or TEXT_FIRMA_REINCIDENCIA not in base.columns:
        return pd.DataFrame(columns=columnas)

    trabajo = base[base[TEXT_FIRMA_REINCIDENCIA].fillna("").ne("")].copy()
    if trabajo.empty:
        return pd.DataFrame(columns=columnas)

    resumen = (
        trabajo.groupby([TEXT_FIRMA_REINCIDENCIA, TEXT_FIRMA_REINCIDENCIA_LABEL], as_index=False)
        .agg(
            **{
                TEXT_CANTIDAD: (TEXT_NUMERO, "count"),
                COL_INCIDENTES_REINCIDENTES: (TEXT_REINCIDENTE, "sum"),
                "Primer incidente": (TEXT_CREADO_DT_2, "min"),
                "Ultimo incidente": (TEXT_CREADO_DT_2, "max"),
            }
        )
        .rename(columns={TEXT_FIRMA_REINCIDENCIA_LABEL: COL_FIRMA_REINCIDENCIA})
    )
    resumen[COL_INCIDENTES_REINCIDENTES] = resumen[COL_INCIDENTES_REINCIDENTES].astype(int)
    resumen = resumen[resumen[COL_INCIDENTES_REINCIDENTES] > 0].copy()
    if resumen.empty:
        return pd.DataFrame(columns=columnas)

    resumen["Tasa reincidencia %"] = resumen.apply(
        lambda row: porcentaje(row[COL_INCIDENTES_REINCIDENTES], row[TEXT_CANTIDAD]),
        axis=1,
    )
    return resumen[columnas].sort_values(
        by=[COL_INCIDENTES_REINCIDENTES, TEXT_CANTIDAD],
        ascending=False,
    )


def lectura_causa_kpi_incidente(row):
    lectura = valor_limpio(row.get(COL_LECTURA_EJECUTIVA))
    detalle = valor_limpio(row.get("Detalle tecnico observado"))
    lecturas_genericas = {
        "Hallazgos tecnicos con bajo volumen o descripcion no estandarizada.",
        "No hay informacion suficiente para explicar la causa.",
    }
    if lectura not in lecturas_genericas:
        return lectura
    if es_detalle_incidente_util(detalle):
        return f"Casos asociados a {detalle.lower()}."
    return "La causa tecnica requiere normalizacion en el cierre para una lectura ejecutiva precisa."


def causa_principal_segmento(causas, segmento):
    filas = causas[causas[TEXT_SEGMENTO] == segmento].sort_values(by=TEXT_CANTIDAD, ascending=False)
    if filas.empty:
        return None
    return filas.iloc[0]


def texto_lectura_causa_segmento(causas, segmento):
    fila = causa_principal_segmento(causas, segmento)
    if fila is None:
        return f"No hay causas raiz para {segmento.lower()} en el periodo."
    return (
        f"{segmento}: la causa raiz principal es {fila[COL_CAUSA_RAIZ]} "
        f"({int(fila[TEXT_CANTIDAD])} casos, {fila['% segmento']}%). "
        f"{fila[COL_LECTURA_EJECUTIVA]} Evidencia: {fila[COL_EVIDENCIA_INCIDENTE]}."
    )


def lineas_lectura_kpi_incidentes(causas):
    lectura_externo = texto_lectura_causa_segmento(causas, "Cliente externo")
    lectura_interno = texto_lectura_causa_segmento(causas, "Cliente interno")
    return [
        (
            "<strong>Cliente externo:</strong> "
            f'{html.escape(lectura_externo.replace("Cliente externo: ", ""))}'
        ),
        (
            "<strong>Cliente interno:</strong> "
            f'{html.escape(lectura_interno.replace("Cliente interno: ", ""))}'
        ),
    ]


def render_lectura_kpi_incidentes(causas):
    lineas = lineas_lectura_kpi_incidentes(causas)
    contenido = f"""
    <div class="executive-note">
        <div class="executive-note-title">Lectura</div>
        <div class="executive-note-detail">{lineas[0]}</div>
        <div class="executive-note-detail">{lineas[1]}</div>
    </div>
    """
    st.markdown(contenido, unsafe_allow_html=True)


def tabla_temas_incidentes_html(causas, segmento):
    datos = causas[causas[TEXT_SEGMENTO] == segmento].sort_values(
        by=[TEXT_CANTIDAD, COL_CAUSA_RAIZ],
        ascending=[False, True],
    )
    if datos.empty:
        return (
            '<div class="executive-note">'
            f'<div class="executive-note-title">{html.escape(segmento)}</div>'
            '<div class="executive-note-detail">No hay causas raiz para este segmento en el periodo.</div>'
            "</div>"
        )

    filas = []
    for _, row in datos.iterrows():
        filas.append(
            "<tr>"
            f"<td>{html.escape(str(row[COL_CAUSA_RAIZ]))}</td>"
            f"<td>{html.escape(str(row[COL_FAMILIA_INCIDENTE]))}</td>"
            f'<td class="number-cell">{int(row[TEXT_CANTIDAD])}</td>'
            f'<td class="number-cell">{html.escape(str(row["% segmento"]))}%</td>'
            f"<td>{html.escape(str(row[COL_LECTURA_EJECUTIVA]))}</td>"
            f"<td>{html.escape(str(row[COL_EVIDENCIA_INCIDENTE]))}</td>"
            "</tr>"
        )

    return f"""
    <div class="executive-table-card">
        <div class="executive-table-title">{html.escape(segmento)}</div>
        <table class="executive-table">
            <thead>
                <tr>
                    <th>Causa raiz</th>
                    <th>Familia</th>
                    <th>Cant.</th>
                    <th>%</th>
                    <th>Lectura</th>
                    <th>Evidencia</th>
                </tr>
            </thead>
            <tbody>
                {''.join(filas)}
            </tbody>
        </table>
    </div>
    """


def render_tablas_temas_kpi_incidentes(causas):
    if causas.empty:
        st.info("No hay causas raiz de incidentes para mostrar en el periodo seleccionado.")
        return

    col_externo, col_interno = st.columns(2)
    segmentos = [
        (col_externo, "Cliente externo"),
        (col_interno, "Cliente interno"),
    ]

    for columna, segmento in segmentos:
        with columna:
            st.markdown(tabla_temas_incidentes_html(causas, segmento), unsafe_allow_html=True)


def render_grafico_causas_kpi_incidentes(causas):
    render_tablas_temas_kpi_incidentes(causas)


def render_reincidencia_kpi_incidentes(base, metricas):
    st.caption(
        "Calculo: incidentes reincidentes / total incidentes x 100. "
        "La reincidencia se marca cuando se repite una firma tecnica despues de un cierre anterior."
    )
    resumen = resumen_reincidencia_kpi_incidentes(base)
    if resumen.empty:
        st.info("No se identificaron incidentes reincidentes con firma tecnica suficiente en el periodo.")
        return

    st.dataframe(resumen, use_container_width=True, hide_index=True)
    st.caption(
        f"Incidentes reincidentes: {metricas['reincidentes']} de {metricas['total']} "
        f"({metricas['tasa_reincidencia']}%)."
    )


def ranking_causas_segmento_kpi(causas, segmento):
    if causas.empty:
        return pd.DataFrame(columns=[COL_CAUSA_RAIZ, TEXT_CANTIDAD])
    return (
        causas[causas[TEXT_SEGMENTO] == segmento]
        .groupby(COL_CAUSA_RAIZ, as_index=False)
        .agg(Cantidad=(TEXT_CANTIDAD, "sum"))
        .sort_values(by=TEXT_CANTIDAD, ascending=False)
    )


def slide_tabla_causa_raiz_incidentes_html(causas):
    if causas.empty:
        filas = '<tr><td colspan="5">Sin causas raiz para mostrar.</td></tr>'
    else:
        tabla = causas.pivot_table(
            index=COL_CAUSA_RAIZ,
            columns=TEXT_SEGMENTO,
            values=TEXT_CANTIDAD,
            aggfunc="sum",
            fill_value=0,
        ).reset_index()
        for segmento in ["Cliente externo", "Cliente interno"]:
            if segmento not in tabla.columns:
                tabla[segmento] = 0
        tabla[TEXT_TOTAL] = tabla["Cliente externo"] + tabla["Cliente interno"]
        total = int(tabla[TEXT_TOTAL].sum())
        tabla["%"] = tabla[TEXT_TOTAL].apply(lambda valor: porcentaje(valor, total))
        tabla = tabla[tabla[TEXT_TOTAL] > 0].sort_values(
            by=[TEXT_TOTAL, COL_CAUSA_RAIZ],
            ascending=[False, True],
        )

        filas_lista = []
        for _, row in tabla.iterrows():
            filas_lista.append(
                "<tr>"
                f'<td class="cause-cell">{html.escape(str(row[COL_CAUSA_RAIZ]))}</td>'
                f'<td class="number-cell">{int(row["Cliente externo"])}</td>'
                f'<td class="number-cell">{int(row["Cliente interno"])}</td>'
                f'<td class="number-cell">{int(row[TEXT_TOTAL])}</td>'
                f'<td class="number-cell">{html.escape(str(row["%"]))}%</td>'
                "</tr>"
            )
        filas = "".join(filas_lista) if filas_lista else '<tr><td colspan="5">Sin causas raiz para mostrar.</td></tr>'

    return f"""
    <div class="slide-panel">
        <div class="slide-panel-title">Causa raiz de incidentes</div>
        <table class="slide-cause-table">
            <colgroup>
                <col style="width: 43%;">
                <col style="width: 14%;">
                <col style="width: 14%;">
                <col style="width: 14%;">
                <col style="width: 15%;">
            </colgroup>
            <thead>
                <tr>
                    <th>Causa raiz</th>
                    <th class="number-header">Externo</th>
                    <th class="number-header">Interno</th>
                    <th class="number-header">Total</th>
                    <th class="number-header">%</th>
                </tr>
            </thead>
            <tbody>{filas}</tbody>
        </table>
    </div>
    """


def render_slide_kpi_incidentes(metricas, causas, mes_dashboard):
    tarjetas = [
        ("Total incidentes", metricas["externos"] + metricas["internos"]),
        ("Abiertos", metricas["abiertos_externos"] + metricas["abiertos_internos"]),
        ("SLA incidentes", f"{metricas['cumplimiento_sla']}%"),
        ("Reincidencia", f"{metricas['tasa_reincidencia']}%"),
    ]
    caption = (
        f"Externos: {metricas['externos']} | Internos: {metricas['internos']} | "
        f"Cerrados cliente externo: {metricas['cerrados_externos']} | "
        f"Cerrados cliente interno: {metricas['cerrados_internos']} | "
        f"Promedio: {metricas['promedio']} h | "
        f"Cumplen SLA: {metricas['cumple_sla']} | No cumplen: {metricas['no_cumple_sla']}"
    )
    lineas = lineas_lectura_kpi_incidentes(causas)
    izquierda = slide_tabla_causa_raiz_incidentes_html(causas)
    derecha = slide_note_html("Lectura", lineas)
    render_slide_component_kpi(MENU_KPI_INCIDENTES, mes_dashboard, tarjetas, caption, izquierda, derecha)


def render_kpi_incidentes(df, mes_dashboard=None):
    base, metricas = preparar_kpi_incidentes(df)
    if base.empty:
        st.info("No hay incidentes cliente interno o externo para el periodo seleccionado.")
        return

    causas = resumen_causas_kpi_incidentes(base)

    modo_diapositiva = st.toggle("Formato diapositiva 16:9", key="slide_kpi_incidentes")
    if modo_diapositiva:
        render_slide_kpi_incidentes(metricas, causas, mes_dashboard)
        return

    if mes_dashboard:
        st.caption(f"{TEXT_PERIODO}{mes_dashboard}")
    st.subheader(MENU_KPI_INCIDENTES)
    st.caption(
        "Base KPI: solo incidentes reales cliente externo e incidentes reales cliente interno. "
        "No incluye Alertas y Consultas NOC ni Casos cliente externo."
    )
    render_tarjetas(
        [
            ("Incidentes cliente externo", metricas["externos"]),
            ("Incidentes cliente interno", metricas["internos"]),
            ("Abiertos cliente externo", metricas["abiertos_externos"]),
            ("Abiertos cliente interno", metricas["abiertos_internos"]),
            ("Reincidencia", f"{metricas['tasa_reincidencia']}%"),
            ("SLA incidentes", f"{metricas['cumplimiento_sla']}%"),
        ]
    )
    st.caption(
        f"Cerrados cliente externo: {metricas['cerrados_externos']} | "
        f"Cerrados cliente interno: {metricas['cerrados_internos']} | "
        f"Reincidentes: {metricas['reincidentes']} | "
        f"Promedio: {metricas['promedio']} h | "
        f"Cumplen SLA: {metricas['cumple_sla']} | No cumplen: {metricas['no_cumple_sla']}"
    )

    render_tablas_temas_kpi_incidentes(causas)
    render_lectura_kpi_incidentes(causas)

    tab_resumen, tab_externo, tab_interno, tab_reincidencia, tab_detalle = st.tabs(
        [TEXT_RESUMEN, "Cliente externo", "Cliente interno", "Reincidencia", "Detalle"]
    )
    with tab_resumen:
        dataframe_liviano(causas, height=360)
    with tab_externo:
        dataframe_liviano(
            causas[causas[TEXT_SEGMENTO] == "Cliente externo"],
            height=360,
        )
    with tab_interno:
        dataframe_liviano(
            causas[causas[TEXT_SEGMENTO] == "Cliente interno"],
            height=360,
        )
    with tab_reincidencia:
        render_reincidencia_kpi_incidentes(base, metricas)
    with tab_detalle:
        temas_detalle = base.apply(clasificacion_tema_incidente, axis=1)
        columnas_temas = [
            COL_CAUSA_RAIZ,
            COL_FAMILIA_INCIDENTE,
            COL_LECTURA_EJECUTIVA,
            COL_ACCION_SUGERIDA,
            COL_EVIDENCIA_INCIDENTE,
        ]
        base = base.copy()
        base[columnas_temas] = pd.DataFrame(temas_detalle.tolist(), index=base.index)
        columnas = [
            TEXT_NUMERO,
            TEXT_SEGMENTO,
            TEXT_ESTADO,
            TEXT_PRIORIDAD,
            TEXT_SERVICIO_NEGOCIO,
            TEXT_CAUSA_RAIZ_AUTO,
            COL_CAUSA_RAIZ,
            COL_FAMILIA_INCIDENTE,
            COL_EVIDENCIA_INCIDENTE,
            TEXT_TIPO_INCIDENTE_AUTO,
            TEXT_DURACION_SLA_HORAS,
            TEXT_ESTADO_SLA,
            COL_REINCIDENTE,
            COL_FIRMA_REINCIDENCIA,
            TEXT_EMPRESA,
            TEXT_SOLICITANTE,
            TEXT_CREADO,
            TEXT_CERRADO,
            TEXT_BREVE_DESCRIPCION,
        ]
        dataframe_liviano(
            base[[col for col in columnas if col in base.columns]],
            height=420,
        )


def es_tipificacion_agendamiento(valor):
    return valor_limpio(valor) == AGENDA_CASE_TIPIFICATION


def canal_agrupado_agendamiento(canal):
    canal_normalizado = normalizar_texto(canal)
    if any(pista in canal_normalizado for pista in AGENDA_DIRECT_CHANNEL_HINTS):
        return "Agenda directa"
    if canal_normalizado in AGENDA_HELP_DESK_CHANNELS:
        return "Mesa de ayuda"
    return "Otro canal"


def texto_analisis_agendamiento(row):
    campos = [
        TEXT_DESCRIPCION_2,
        TEXT_CAUSA,
        TEXT_CODIGO_RESOLUCION,
        "notas_resolucion",
        TEXT_OBSERVACIONES_ADICIONALES,
        TEXT_OBSERVACIONES_TRABAJO,
        TEXT_PRODUCTO,
    ]
    return " ".join(normalizar_texto(row.get(campo)) for campo in campos).strip()


def inferir_motivo_agendamiento(row):
    texto = texto_analisis_agendamiento(row)
    for motivo, palabras in AGENDA_REASON_RULES:
        if any(palabra in texto for palabra in palabras):
            return motivo
    return "Sin detalle suficiente"


def cliente_agendamiento(row):
    cuenta = valor_limpio(row.get(TEXT_CUENTA))
    return cuenta if cuenta else SIN_CUENTA


def fecha_corta(valor):
    if pd.isna(valor):
        return ""
    return pd.Timestamp(valor).strftime("%Y-%m-%d")


def inicio_periodo_agendamiento(agenda, mes_dashboard):
    if mes_dashboard != TEXT_TODOS:
        try:
            return pd.Period(mes_dashboard, freq="M").to_timestamp()
        except Exception:
            pass
    fechas = agenda.get(TEXT_CREADO_DT_DASHBOARD, pd.Series(dtype=PANDAS_DATETIME_DTYPE)).dropna()
    if fechas.empty:
        return None
    return fechas.min().normalize()


AGENDAMIENTO_COLUMNAS_BASE = [
    COL_CANAL_AGRUPADO,
    COL_MOTIVO_INFERIDO,
    COL_CLIENTE_AGENDA,
    COL_CICLO_CLIENTE,
    COL_CLIENTE_RECURRENTE_AGENDA,
    COL_REINCIDENTE_AGENDA,
    COL_CASOS_HISTORICOS_CLIENTE,
    COL_AGENDAS_HISTORICAS_CLIENTE,
]


def agregar_columnas_agendamiento_vacias(df):
    trabajo = df.copy()
    for columna in AGENDAMIENTO_COLUMNAS_BASE:
        trabajo[columna] = pd.Series(dtype=TEXT_OBJECT)
    return trabajo


def preparar_fuentes_agendamiento(df_periodo, df_historico):
    periodo = normalizar_tipificaciones_casos_df(df_periodo).copy()
    historico = normalizar_tipificaciones_casos_df(df_historico).copy()
    if TEXT_CREADO_DT_DASHBOARD not in periodo.columns:
        periodo = preparar_fechas_dashboard(periodo)
    if TEXT_CREADO_DT_DASHBOARD not in historico.columns:
        historico = preparar_fechas_dashboard(historico)
    return periodo, historico


def asegurar_columnas_agenda(agenda, historico):
    for columna in [TEXT_CANAL, TEXT_CUENTA, TEXT_NUMERO]:
        if columna not in agenda.columns:
            agenda[columna] = ""
        if columna not in historico.columns:
            historico[columna] = ""


def agregar_datos_basicos_agenda(agenda):
    agenda[COL_CANAL_AGRUPADO] = agenda[TEXT_CANAL].apply(canal_agrupado_agendamiento)
    agenda[COL_MOTIVO_INFERIDO] = agenda.apply(inferir_motivo_agendamiento, axis=1)
    agenda[COL_CLIENTE_AGENDA] = agenda.apply(cliente_agendamiento, axis=1)
    agenda[TEXT_CLIENTE_NORM_AGENDA] = agenda[TEXT_CUENTA].apply(normalizar_texto)


def agregar_historial_agendamiento(agenda, historico):
    historico[TEXT_CLIENTE_NORM_AGENDA] = historico[TEXT_CUENTA].apply(normalizar_texto)
    historico_validos = historico[historico[TEXT_CLIENTE_NORM_AGENDA] != ""].copy()
    if historico_validos.empty:
        agenda[COL_PRIMERA_ATENCION_CLIENTE] = pd.NaT
        agenda[COL_CASOS_HISTORICOS_CLIENTE] = 0
        agenda[COL_AGENDAS_HISTORICAS_CLIENTE] = 0
        return

    historia_cliente = historico_validos.groupby(TEXT_CLIENTE_NORM_AGENDA).agg(
        Primera_atencion_cliente=(TEXT_CREADO_DT_DASHBOARD, "min"),
        Casos_historicos_cliente=(TEXT_NUMERO, "nunique"),
    )
    historico_agenda = historico_validos[
        historico_validos[TEXT_TIPIFICACION_2].apply(es_tipificacion_agendamiento)
    ].copy()
    agendas_historicas = historico_agenda.groupby(TEXT_CLIENTE_NORM_AGENDA)[TEXT_NUMERO].nunique()
    agenda[COL_PRIMERA_ATENCION_CLIENTE] = agenda[TEXT_CLIENTE_NORM_AGENDA].map(
        historia_cliente["Primera_atencion_cliente"]
    )
    agenda[COL_CASOS_HISTORICOS_CLIENTE] = (
        agenda[TEXT_CLIENTE_NORM_AGENDA]
        .map(historia_cliente["Casos_historicos_cliente"])
        .fillna(0)
        .astype(int)
    )
    agenda[COL_AGENDAS_HISTORICAS_CLIENTE] = (
        agenda[TEXT_CLIENTE_NORM_AGENDA].map(agendas_historicas).fillna(0).astype(int)
    )


def marcar_reincidencia_agendamiento(agenda, historico):
    agenda[COL_REINCIDENTE_AGENDA] = "No"
    if agenda.empty or TEXT_CLIENTE_NORM_AGENDA not in agenda.columns:
        return

    historico_agenda = historico[
        historico[TEXT_TIPIFICACION_2].apply(es_tipificacion_agendamiento)
    ].copy()
    if historico_agenda.empty:
        return
    if TEXT_CLIENTE_NORM_AGENDA not in historico_agenda.columns:
        historico_agenda[TEXT_CLIENTE_NORM_AGENDA] = historico_agenda[TEXT_CUENTA].apply(normalizar_texto)

    historico_agenda = historico_agenda[historico_agenda[TEXT_CLIENTE_NORM_AGENDA] != ""].copy()
    if historico_agenda.empty:
        return

    historico_agenda = historico_agenda.sort_values(
        by=[TEXT_CLIENTE_NORM_AGENDA, TEXT_CREADO_DT_DASHBOARD, TEXT_NUMERO],
        na_position="last",
    ).copy()
    historico_agenda["_agenda_previa_cliente"] = historico_agenda.groupby(TEXT_CLIENTE_NORM_AGENDA).cumcount()
    previas_por_numero = (
        historico_agenda.dropna(subset=[TEXT_NUMERO])
        .set_index(TEXT_NUMERO)["_agenda_previa_cliente"]
        .to_dict()
    )
    fechas_por_cliente = {
        cliente: grupo[TEXT_CREADO_DT_DASHBOARD].dropna()
        for cliente, grupo in historico_agenda.groupby(TEXT_CLIENTE_NORM_AGENDA)
    }

    def es_reincidente(row):
        numero = valor_limpio(row.get(TEXT_NUMERO))
        if numero in previas_por_numero:
            return previas_por_numero[numero] > 0

        cliente = row.get(TEXT_CLIENTE_NORM_AGENDA)
        if not cliente:
            return False

        creado = row.get(TEXT_CREADO_DT_DASHBOARD)
        fechas_cliente = fechas_por_cliente.get(cliente, pd.Series(dtype=PANDAS_DATETIME_DTYPE))
        if pd.notna(creado) and not fechas_cliente.empty:
            return bool((fechas_cliente < creado).any())

        return row.get(COL_AGENDAS_HISTORICAS_CLIENTE, 0) > 1

    agenda[COL_REINCIDENTE_AGENDA] = agenda.apply(
        lambda row: "Si" if es_reincidente(row) else "No",
        axis=1,
    )


def clasificar_ciclo_agendamiento(row, inicio_periodo):
    if row[TEXT_CLIENTE_NORM_AGENDA] == "":
        return SIN_CUENTA
    primera = row.get(COL_PRIMERA_ATENCION_CLIENTE)
    if inicio_periodo is not None and pd.notna(primera) and primera < inicio_periodo:
        return "Cliente con historial previo"
    if row.get(COL_AGENDAS_HISTORICAS_CLIENTE, 0) > 1:
        return "Cliente recurrente en el periodo"
    return "Cliente nuevo en la base"


def agregar_ciclo_agendamiento(agenda, mes_dashboard):
    inicio_periodo = inicio_periodo_agendamiento(agenda, mes_dashboard)
    agenda[COL_CICLO_CLIENTE] = agenda.apply(
        lambda row: clasificar_ciclo_agendamiento(row, inicio_periodo),
        axis=1,
    )
    agenda[COL_CLIENTE_RECURRENTE_AGENDA] = agenda[COL_AGENDAS_HISTORICAS_CLIENTE].apply(
        lambda valor: "Si" if valor > 1 else "No"
    )


def preparar_analisis_agendamiento(df_periodo, df_historico, mes_dashboard):
    if df_periodo.empty:
        return agregar_columnas_agendamiento_vacias(df_periodo)

    periodo, historico = preparar_fuentes_agendamiento(df_periodo, df_historico)
    agenda = periodo[periodo[TEXT_TIPIFICACION_2].apply(es_tipificacion_agendamiento)].copy()
    if agenda.empty:
        return agregar_columnas_agendamiento_vacias(agenda)

    asegurar_columnas_agenda(agenda, historico)
    agregar_datos_basicos_agenda(agenda)
    agregar_historial_agendamiento(agenda, historico)
    marcar_reincidencia_agendamiento(agenda, historico)
    agregar_ciclo_agendamiento(agenda, mes_dashboard)
    return agenda


def metricas_reincidencia_agendamiento(agenda):
    if agenda.empty or COL_REINCIDENTE_AGENDA not in agenda.columns:
        return 0, 0
    reincidentes = int((agenda[COL_REINCIDENTE_AGENDA] == "Si").sum())
    return reincidentes, porcentaje(reincidentes, len(agenda))

def lectura_ejecutiva_agendamiento(agenda, total_casos_periodo):
    columnas = [TEXT_PREGUNTA, TEXT_LECTURA, TEXT_EVIDENCIA, COL_ACCION_SUGERIDA]
    if agenda.empty:
        return pd.DataFrame(columns=columnas)

    total_agenda = len(agenda)
    mesa = int((agenda[COL_CANAL_AGRUPADO] == "Mesa de ayuda").sum())
    directa = int((agenda[COL_CANAL_AGRUPADO] == "Agenda directa").sum())
    token_fisico = int((agenda[COL_MOTIVO_INFERIDO] == AGENDA_MOTIVO_TOKEN_FISICO).sum())
    reincidentes, porcentaje_reincidencia = metricas_reincidencia_agendamiento(agenda)
    clientes_reincidentes = agenda[agenda[COL_REINCIDENTE_AGENDA] == "Si"][COL_CLIENTE_AGENDA].nunique()
    motivo_principal = valor_mas_frecuente(agenda[COL_MOTIVO_INFERIDO])
    canal_principal = valor_mas_frecuente(agenda[TEXT_CANAL].replace("", "Sin canal"))
    producto_principal = valor_mas_frecuente(agenda.get(TEXT_PRODUCTO, pd.Series(dtype=TEXT_OBJECT)))

    return pd.DataFrame(
        [
            {
                TEXT_PREGUNTA: "Por que siguen entrando por mesa de ayuda?",
                TEXT_LECTURA: (
                    "La mayoria no esta entrando por el calendario directo, sino por canales operativos "
                    "que terminan en mesa de ayuda."
                ),
                TEXT_EVIDENCIA: (
                    f"{mesa} de {total_agenda} casos de agendamiento ({porcentaje(mesa, total_agenda)}%) "
                    f"entraron por mesa de ayuda. Canal principal: {canal_principal}. "
                    f"Agenda directa: {directa} casos."
                ),
                COL_ACCION_SUGERIDA: (
                    "Redirigir Web/telefono/correo al enlace de agenda antes de crear caso y marcar excepciones "
                    "cuando mesa de ayuda deba tomarlo."
                ),
            },
            {
                TEXT_PREGUNTA: "Cual es la razon de tantos casos?",
                TEXT_LECTURA: (
                    f"El motivo dominante es {motivo_principal}; esta tipologia concentra "
                    f"{porcentaje(total_agenda, total_casos_periodo)}% de los casos del periodo."
                ),
                TEXT_EVIDENCIA: f"Producto principal asociado: {producto_principal}.",
                COL_ACCION_SUGERIDA: (
                    "Separar en el cierre si fue instalacion, activacion/descarga, token fisico o firma; "
                    "asi la causa queda accionable y no solo como agendamiento."
                ),
            },
            {
                TEXT_PREGUNTA: "Cuanto pesa token fisico/ePass?",
                TEXT_LECTURA: (
                    "Token fisico/ePass es un motivo operativo relevante porque suele requerir validacion "
                    "del dispositivo, pruebas o una sesion asistida."
                ),
                TEXT_EVIDENCIA: (
                    f"{token_fisico} de {total_agenda} casos de agendamiento "
                    f"({porcentaje(token_fisico, total_agenda)}%) fueron inferidos como {AGENDA_MOTIVO_TOKEN_FISICO}."
                ),
                COL_ACCION_SUGERIDA: (
                    "Marcar token fisico/ePass de forma explicita en causa o resolucion para separar este motivo "
                    "de una agenda generica."
                ),
            },
            {
                TEXT_PREGUNTA: "Que porcentaje es reincidente?",
                TEXT_LECTURA: (
                    "La reincidencia muestra casos de agenda que vuelven a aparecer para un cliente que ya tenia "
                    "un redireccionamiento previo."
                ),
                TEXT_EVIDENCIA: (
                    f"{reincidentes} de {total_agenda} casos de agendamiento son reincidentes "
                    f"({porcentaje_reincidencia}%). Clientes involucrados: {clientes_reincidentes}."
                ),
                COL_ACCION_SUGERIDA: (
                    "Revisar los clientes con mas de una agenda para separar si falta comunicacion del enlace, "
                    "si hay bloqueo tecnico recurrente o si el cierre esta quedando generico."
                ),
            },
        ],
        columns=columnas,
    )


def resumen_clientes_agendamiento(agenda):
    columnas = [
        TEXT_CLIENTE,
        COL_CASOS_AGENDA,
        COL_CASOS_REINCIDENTES_AGENDA,
        COL_REINCIDENCIA_AGENDAMIENTO,
        "Canal principal",
        "Motivo principal",
        COL_CICLO_CLIENTE,
        "Agendas historicas",
        "Casos historicos",
        COL_PRIMERA_AGENDA,
        COL_ULTIMA_AGENDA,
    ]
    if agenda.empty:
        return pd.DataFrame(columns=columnas)

    resumen = (
        agenda.groupby(COL_CLIENTE_AGENDA, dropna=False)
        .agg(
            Casos_agenda=(TEXT_NUMERO, TEXT_COUNT),
            Casos_reincidentes_agenda=(COL_REINCIDENTE_AGENDA, lambda serie: int((serie == "Si").sum())),
            Canal_principal=(COL_CANAL_AGRUPADO, valor_mas_frecuente),
            Motivo_principal=(COL_MOTIVO_INFERIDO, valor_mas_frecuente),
            Ciclo_cliente=(COL_CICLO_CLIENTE, valor_mas_frecuente),
            Agendas_historicas=(COL_AGENDAS_HISTORICAS_CLIENTE, "max"),
            Casos_historicos=(COL_CASOS_HISTORICOS_CLIENTE, "max"),
            Primera_agenda=(TEXT_CREADO_DT_DASHBOARD, "min"),
            Ultima_agenda=(TEXT_CREADO_DT_DASHBOARD, "max"),
        )
        .reset_index()
        .rename(
            columns={
                COL_CLIENTE_AGENDA: TEXT_CLIENTE,
                "Casos_agenda": COL_CASOS_AGENDA,
                "Casos_reincidentes_agenda": COL_CASOS_REINCIDENTES_AGENDA,
                "Canal_principal": "Canal principal",
                "Motivo_principal": "Motivo principal",
                "Ciclo_cliente": COL_CICLO_CLIENTE,
                "Agendas_historicas": "Agendas historicas",
                "Casos_historicos": "Casos historicos",
                "Primera_agenda": COL_PRIMERA_AGENDA,
                "Ultima_agenda": COL_ULTIMA_AGENDA,
            }
        )
    )
    resumen[COL_REINCIDENCIA_AGENDAMIENTO] = resumen.apply(
        lambda row: porcentaje(row[COL_CASOS_REINCIDENTES_AGENDA], row[COL_CASOS_AGENDA]),
        axis=1,
    )
    resumen[COL_PRIMERA_AGENDA] = resumen[COL_PRIMERA_AGENDA].apply(fecha_corta)
    resumen[COL_ULTIMA_AGENDA] = resumen[COL_ULTIMA_AGENDA].apply(fecha_corta)
    return resumen.sort_values(by=[COL_CASOS_AGENDA, TEXT_CLIENTE], ascending=[False, True])[columnas]


def resumen_canales_agendamiento(agenda):
    columnas = [COL_CANAL_AGRUPADO, TEXT_CANAL_2, TEXT_CASOS, "% agendamiento"]
    if agenda.empty:
        return pd.DataFrame(columns=columnas)

    resumen = (
        agenda.assign(Canal=agenda[TEXT_CANAL].replace("", "Sin canal"))
        .groupby([COL_CANAL_AGRUPADO, TEXT_CANAL_2], dropna=False)
        .size()
        .reset_index(name=TEXT_CASOS)
        .sort_values(by=[TEXT_CASOS, TEXT_CANAL_2], ascending=[False, True])
    )
    total = len(agenda)
    resumen["% agendamiento"] = resumen[TEXT_CASOS].apply(lambda valor: porcentaje(valor, total))
    return resumen[columnas]


def data_reincidencia_agendamiento(agenda):
    columnas = [
        TEXT_NUMERO,
        COL_CLIENTE_AGENDA,
        TEXT_CUENTA,
        TEXT_CANAL,
        COL_CANAL_AGRUPADO,
        COL_MOTIVO_INFERIDO,
        COL_CICLO_CLIENTE,
        COL_AGENDAS_HISTORICAS_CLIENTE,
        COL_CASOS_HISTORICOS_CLIENTE,
        TEXT_PRODUCTO,
        TEXT_ASIGNADO,
        TEXT_CREADO,
        TEXT_ESTADO,
        TEXT_DESCRIPCION_2,
        TEXT_CAUSA,
    ]
    if agenda.empty or COL_REINCIDENTE_AGENDA not in agenda.columns:
        return pd.DataFrame(columns=columnas)

    reincidentes = agenda[agenda[COL_REINCIDENTE_AGENDA] == "Si"].copy()
    if reincidentes.empty:
        return pd.DataFrame(columns=columnas)

    visibles = [col for col in columnas if col in reincidentes.columns]
    return reincidentes.sort_values(by=TEXT_CREADO_DT_DASHBOARD, ascending=False)[visibles]


def render_analisis_agendamiento_mesa(df_periodo, df_historico, mes_dashboard):
    agenda = preparar_analisis_agendamiento(df_periodo, df_historico, mes_dashboard)

    st.subheader("Analisis de agendamiento por mesa de ayuda")
    st.caption(
        "Cruce para explicar por que los casos de redireccionamiento a agenda siguen llegando por mesa de ayuda, "
        "que motivo operativo se repite y cuanto pesa token fisico/ePass."
    )

    if agenda.empty:
        st.info("No hay casos de redireccionamiento a agenda en el periodo seleccionado.")
        return

    total_agenda = len(agenda)
    mesa = int((agenda[COL_CANAL_AGRUPADO] == "Mesa de ayuda").sum())
    token_fisico = int((agenda[COL_MOTIVO_INFERIDO] == AGENDA_MOTIVO_TOKEN_FISICO).sum())
    reincidentes, porcentaje_reincidencia = metricas_reincidencia_agendamiento(agenda)

    render_tarjetas(
        [
            ("Agendamiento", total_agenda),
            ("Mesa ayuda", f"{mesa} ({porcentaje(mesa, total_agenda)}%)"),
            ("Token fisico/ePass", f"{token_fisico} ({porcentaje(token_fisico, total_agenda)}%)"),
            ("Reincidencia", f"{reincidentes} ({porcentaje_reincidencia}%)"),
        ]
    )

    tab_lectura, tab_motivos, tab_canales, tab_reincidencias, tab_clientes, tab_detalle = st.tabs(
        [TEXT_LECTURA, "Motivos", "Canales", "Reincidencias", TEXT_CLIENTE, "Detalle"]
    )

    with tab_lectura:
        st.dataframe(
            lectura_ejecutiva_agendamiento(agenda, len(df_periodo)),
            use_container_width=True,
            hide_index=True,
        )

    with tab_motivos:
        motivos = (
            agenda[COL_MOTIVO_INFERIDO]
            .value_counts()
            .rename_axis("Motivo")
            .reset_index(name=TEXT_CASOS)
            .sort_values(by=TEXT_CASOS, ascending=True)
        )
        fig = px.bar(
            motivos,
            x=TEXT_CASOS,
            y="Motivo",
            orientation="h",
            text=TEXT_CASOS,
            color_discrete_sequence=[UI_PALETTE[TEXT_YELLOW]],
        )
        fig.update_traces(marker_color=UI_PALETTE[TEXT_YELLOW], textposition=TEXT_OUTSIDE)
        st.plotly_chart(aplicar_estilo_figura(fig, COL_MOTIVO_INFERIDO), use_container_width=True)

    with tab_canales:
        canales = resumen_canales_agendamiento(agenda)
        st.caption("Entrada por canal de los casos mostrados en este analisis.")
        st.dataframe(canales, use_container_width=True, hide_index=True)

    with tab_reincidencias:
        data_reincidentes = data_reincidencia_agendamiento(agenda)
        st.caption("Data base del porcentaje: casos con una agenda previa del mismo cliente.")
        if data_reincidentes.empty:
            st.info("No hay casos reincidentes de agendamiento en el periodo seleccionado.")
        else:
            st.dataframe(data_reincidentes, use_container_width=True, hide_index=True)
            st.download_button(
                "Descargar data reincidencias",
                data_reincidentes.to_csv(index=False).encode("utf-8"),
                file_name="reincidencias_agendamiento.csv",
                mime="text/csv",
            )

    with tab_clientes:
        st.caption(
            "Reincidencia por cliente: casos de agenda con un redireccionamiento previo del mismo cliente."
        )
        st.dataframe(resumen_clientes_agendamiento(agenda), use_container_width=True, hide_index=True)

    with tab_detalle:
        columnas = [
            TEXT_NUMERO,
            TEXT_CUENTA,
            TEXT_CANAL,
            COL_CANAL_AGRUPADO,
            COL_MOTIVO_INFERIDO,
            COL_REINCIDENTE_AGENDA,
            COL_CLIENTE_RECURRENTE_AGENDA,
            COL_AGENDAS_HISTORICAS_CLIENTE,
            TEXT_PRODUCTO,
            TEXT_ASIGNADO,
            TEXT_CREADO,
            TEXT_ESTADO,
            TEXT_DESCRIPCION_2,
            TEXT_CAUSA,
        ]
        visibles = [col for col in columnas if col in agenda.columns]
        st.dataframe(
            agenda.sort_values(by=TEXT_CREADO_DT_DASHBOARD, ascending=False)[visibles],
            use_container_width=True,
            hide_index=True,
        )


COLUMNAS_RESUMEN_TIPOLOGIAS = [
    TEXT_TIPOLOGIA,
    COL_LECTURA_EJECUTIVA,
    TEXT_TOTAL,
    TEXT_CERRADOS,
    TEXT_ABIERTOS,
    "SLA aplica",
    "Objetivo atencion",
    COL_CUMPLE_SLA,
    COL_NO_CUMPLE_SLA,
    "SLA %",
    COL_PROM_HORAS,
    COL_PROM_DIAS,
]


def preparar_base_tipologias_incidentes(df):
    trabajo = df.copy()
    trabajo[TEXT_TIPIFICACION_AUTO] = trabajo[TEXT_TIPIFICACION_AUTO].fillna(TIPIFICACION_INCIDENTE_INTERNO)
    trabajo[TEXT_CERRADO_2] = mascara_cerrados(trabajo)
    duracion_base = TEXT_DURACION_SLA_HORAS if TEXT_DURACION_SLA_HORAS in trabajo.columns else TEXT_DURACION_HORAS
    trabajo["_duracion_horas_num"] = pd.to_numeric(trabajo[duracion_base], errors=TEXT_COERCE)
    return trabajo


def ordenar_tipologias_incidentes(trabajo):
    tipologias_disponibles = set(trabajo[TEXT_TIPIFICACION_AUTO].dropna())
    tipologias = [tipologia for tipologia in INCIDENT_TIPIFICATION_ORDER if tipologia in tipologias_disponibles]
    tipologias.extend(
        sorted(
            tipologia
            for tipologia in trabajo[TEXT_TIPIFICACION_AUTO].dropna().unique().tolist()
            if tipologia not in tipologias
        )
    )
    return tipologias


def objetivo_atencion_tipologia(objetivos):
    if len(objetivos) == 1:
        return formato_horas_dias(objetivos[0])
    if len(objetivos) > 1:
        return "Segun prioridad"
    return "Sin objetivo"


def calcular_sla_tipologia(cerrados, objetivos):
    sla_aplica = len(objetivos) > 0
    con_sla = (
        cerrados[cerrados[TEXT_ESTADO_SLA].isin([ESTADO_SLA_CUMPLE, ESTADO_SLA_NO_CUMPLE])].copy()
        if sla_aplica
        else cerrados.iloc[0:0].copy()
    )
    cumple = int((con_sla[TEXT_ESTADO_SLA] == ESTADO_SLA_CUMPLE).sum())
    no_cumple = int((con_sla[TEXT_ESTADO_SLA] == ESTADO_SLA_NO_CUMPLE).sum())
    total_sla = cumple + no_cumple
    if total_sla:
        sla_porcentaje = f"{porcentaje(cumple, total_sla)}%"
    elif sla_aplica:
        sla_porcentaje = "Sin cierre"
    else:
        sla_porcentaje = "No aplica"
    return sla_aplica, cumple, no_cumple, sla_porcentaje


def fila_resumen_tipologia(trabajo, tipologia):
    grupo = trabajo[trabajo[TEXT_TIPIFICACION_AUTO] == tipologia].copy()
    cerrados = grupo[grupo[TEXT_CERRADO_2]].copy()
    objetivos = pd.to_numeric(
        grupo.get(TEXT_SLA_OBJETIVO_HORAS, pd.Series(dtype=TEXT_FLOAT)),
        errors=TEXT_COERCE,
    ).dropna().unique()
    sla_aplica, cumple, no_cumple, sla_porcentaje = calcular_sla_tipologia(cerrados, objetivos)
    promedio_horas = cerrados["_duracion_horas_num"].dropna().mean()

    return {
        TEXT_TIPOLOGIA: tipologia,
        COL_LECTURA_EJECUTIVA: INCIDENT_TIPIFICATION_GUIDE.get(
            tipologia,
            "Tipologia detectada en el archivo cargado.",
        ),
        TEXT_TOTAL: len(grupo),
        TEXT_CERRADOS: len(cerrados),
        TEXT_ABIERTOS: len(grupo) - len(cerrados),
        "SLA aplica": "Si" if sla_aplica else "No aplica",
        "Objetivo atencion": objetivo_atencion_tipologia(objetivos),
        COL_CUMPLE_SLA: cumple if sla_aplica else pd.NA,
        COL_NO_CUMPLE_SLA: no_cumple if sla_aplica else pd.NA,
        "SLA %": sla_porcentaje,
        COL_PROM_HORAS: round(promedio_horas, 2) if pd.notna(promedio_horas) else pd.NA,
        COL_PROM_DIAS: round(promedio_horas / 24, 2) if pd.notna(promedio_horas) else pd.NA,
    }


def convertir_tipos_resumen_tipologias(resumen):
    for columna in [TEXT_TOTAL, TEXT_CERRADOS, TEXT_ABIERTOS, COL_CUMPLE_SLA, COL_NO_CUMPLE_SLA]:
        resumen[columna] = resumen[columna].astype("Int64")
    for columna in [COL_PROM_HORAS, COL_PROM_DIAS]:
        resumen[columna] = resumen[columna].astype("Float64")
    return resumen


def tabla_resumen_tipologias_incidentes(df):
    if df.empty:
        return pd.DataFrame(columns=COLUMNAS_RESUMEN_TIPOLOGIAS)

    trabajo = preparar_base_tipologias_incidentes(df)
    filas = [fila_resumen_tipologia(trabajo, tipologia) for tipologia in ordenar_tipologias_incidentes(trabajo)]
    return convertir_tipos_resumen_tipologias(pd.DataFrame(filas))

def texto_incidente_para_tema(row):
    campos = [
        TEXT_CAUSA_RAIZ_AUTO,
        TEXT_SERVICIO_NEGOCIO,
        TEXT_TIPO_FALLA,
        TEXT_BREVE_DESCRIPCION,
        TEXT_DESCRIPCION_2,
        TEXT_OBSERVACIONES_TRABAJO,
        TEXT_OBSERVACIONES_ADICIONALES,
        "categoria",
        "impacto",
        "actualizaciones",
        "lista_notas_trabajo",
    ]
    return " ".join(normalizar_texto(row.get(campo)) for campo in campos).strip()


INCIDENT_DETAIL_NOISE_VALUES = {
    "",
    "all",
    "todos",
    "todo",
    "na",
    "n/a",
    "no aplica",
    "no apica",
    "no definido",
    "sin dato",
    "sin datos",
    "sin informacion",
    "sin informacion suficiente",
    "sin inferencia",
    "sin patron",
    "sin patron concluyente",
    "null",
    "none",
    "-",
    ".",
}


def es_detalle_incidente_util(valor):
    valor_texto = valor_limpio(valor)
    normalizado = normalizar_texto(valor_texto)
    if not normalizado or normalizado in INCIDENT_DETAIL_NOISE_VALUES:
        return False
    if "sin patron" in normalizado or "sin inferencia" in normalizado:
        return False
    return True


def valor_detalle_incidente(row):
    for campo in [
        TEXT_SERVICIO_NEGOCIO,
        TEXT_TIPO_FALLA,
        TEXT_CAUSA_RAIZ_AUTO,
        TEXT_BREVE_DESCRIPCION,
        TEXT_DESCRIPCION_2,
    ]:
        valor = valor_limpio(row.get(campo))
        if es_detalle_incidente_util(valor):
            return valor
    return "Validacion tecnica complementaria"


def tema_revision_especifica(row):
    return "Validacion tecnica complementaria"


def clasificacion_tema_incidente(row):
    texto = texto_incidente_para_tema(row)
    detalle = valor_detalle_incidente(row)

    reglas = [
        (
            ["phishing", "pishing", "suplantacion", "fraude", "malicioso", "correo sospechoso"],
            "Infraestructura, accesos, seguridad y proveedor",
            "Operacion tecnica",
            "Reporte asociado a seguridad, acceso, suplantacion, fraude o correo sospechoso.",
            "Validar origen, bloquear indicadores, revisar accesos y reforzar comunicacion preventiva.",
        ),
        (
            [
                "ocsp",
                "rpost",
                "portal rpost",
                "caida",
                "indisponibilidad",
                "degradacion",
                "no disponible",
                "disponibilidad",
                "monitoreo",
                "noc",
                "servicio caido",
                "fuera de servicio",
            ],
            "Disponibilidad del servicio",
            "Disponibilidad",
            "Caida, indisponibilidad o degradacion de un servicio, plataforma o componente monitoreado.",
            "Revisar ventana de afectacion, recurrencia, dependencias y comunicacion a clientes.",
        ),
        (
            [
                "certimail",
                "certi mail",
                "certicmal",
                "correo",
                "notificacion",
                "notificaciones",
                "smtp",
                "mail",
                "envio",
                "recepcion",
                "procesamiento",
                "acuse",
                "acuses",
            ],
            "Comunicaciones y acuses",
            "Comunicaciones",
            "Fallas en envio, recepcion, acuses o procesamiento de notificaciones al cliente.",
            "Revisar colas, rebotes, trazabilidad, acuses y proveedor de correo.",
        ),
        (
            [
                "firma",
                "firmar",
                "validacion",
                "validar",
                "documento no firma",
                "tsa",
                "certificado",
                "cadena de confianza",
                "ssl",
                "certificados",
            ],
            "Firma, certificados y validacion",
            "Firma y validacion",
            "Dificultades para firmar, validar, sellar tiempo o operar certificados digitales.",
            "Revisar flujo de firma, cadena de confianza, TSA, mensaje de error y recurrencia por producto.",
        ),
        (
            [
                "base de datos",
                "database",
                "sql",
                " bd ",
                "red",
                "conectividad",
                "vpn",
                "latencia",
                "enlace",
                "comunicacion",
                "servidor",
                "cpu",
                "memoria",
                "disco",
                "infraestructura",
            ],
            "Infraestructura, accesos, seguridad y proveedor",
            "Operacion tecnica",
            "Afectacion asociada a infraestructura, conectividad, recursos, accesos o componentes tecnicos.",
            "Validar capacidad, trazas, conectividad, eventos de sistema y recurrencia del componente.",
        ),
        (
            [
                "ldap",
                "directorio activo",
                "active directory",
                "login",
                "autenticacion",
                "acceso",
                "usuario",
                "clave",
                "password",
                "permiso",
            ],
            "Infraestructura, accesos, seguridad y proveedor",
            "Operacion tecnica",
            "Problemas de acceso, autenticacion, directorio o permisos.",
            "Validar permisos, autenticacion, trazas del usuario y dependencia tecnica.",
        ),
        (
            ["duplicad"],
            "Validacion tecnica complementaria",
            "Validacion tecnica",
            "Registros repetidos que pueden distorsionar la lectura operativa.",
            "Depurar duplicados y ajustar reglas de cargue o cierre.",
        ),
        (
            ["proveedor", "tercero", "escalado", "escalamiento"],
            "Infraestructura, accesos, seguridad y proveedor",
            "Operacion tecnica",
            "Incidentes relacionados con escalamiento, dependencia o gestion de proveedor.",
            "Validar responsable, tiempos de respuesta del proveedor y acuerdos de escalamiento.",
        ),
    ]

    for palabras, tema, familia, lectura, accion in reglas:
        if any(palabra in texto for palabra in palabras):
            return tema, familia, lectura, accion, detalle

    tema = tema_revision_especifica(row)
    return (
        tema,
        "Validacion tecnica",
        "La informacion disponible no permite inferir con certeza una familia tecnica principal.",
        "Revisar servicio afectado, causa registrada y detalle de cierre para fortalecer la lectura de causa raiz.",
        detalle,
    )


def clasificar_causa_incidente(causa):
    row = {TEXT_CAUSA_RAIZ_AUTO: causa}
    tema, _familia, lectura, accion, _detalle = clasificacion_tema_incidente(row)
    return tema, lectura, accion


def clasificar_causa_cliente_externo(causa):
    return clasificar_causa_incidente(causa)


def resumen_causas_incidentes(df, porcentaje_columna="% incidentes"):
    columnas = [
        COL_CAUSA_RAIZ,
        COL_FAMILIA_INCIDENTE,
        TEXT_CANTIDAD,
        porcentaje_columna,
        COL_LECTURA_EJECUTIVA,
        COL_EVIDENCIA_INCIDENTE,
        COL_ACCION_SUGERIDA,
        "Detalle tecnico observado",
    ]
    if df.empty:
        return pd.DataFrame(columns=columnas)

    trabajo = df.copy()
    trabajo[TEXT_CAUSA_TECNICA] = trabajo[TEXT_CAUSA_RAIZ_AUTO].replace("", pd.NA).fillna("Sin inferencia")
    clasificacion = trabajo.apply(clasificacion_tema_incidente, axis=1)
    trabajo[
        [
            COL_CAUSA_RAIZ,
            COL_FAMILIA_INCIDENTE,
            COL_LECTURA_EJECUTIVA,
            COL_ACCION_SUGERIDA,
            COL_EVIDENCIA_INCIDENTE,
        ]
    ] = pd.DataFrame(
        clasificacion.tolist(),
        index=trabajo.index,
    )

    total = len(trabajo)

    def detalles_tecnicos(serie):
        valores_utiles = [
            valor_limpio(valor)
            for valor in serie.dropna().astype(str).tolist()
            if es_detalle_incidente_util(valor)
        ]
        if not valores_utiles:
            return "Validacion tecnica complementaria"
        valores = pd.Series(valores_utiles).value_counts().head(2).index.tolist()
        return "; ".join(valores)

    def primer_valor_util(serie):
        for valor in serie.dropna().astype(str).tolist():
            limpio = valor_limpio(valor)
            if limpio:
                return limpio
        return ""

    resumen = (
        trabajo.groupby(
            [
                COL_CAUSA_RAIZ,
                COL_FAMILIA_INCIDENTE,
            ],
            dropna=False,
        )
        .agg(
            Cantidad=(TEXT_NUMERO, TEXT_COUNT),
            Lectura_ejecutiva=(COL_LECTURA_EJECUTIVA, primer_valor_util),
            Evidencia_observada=(COL_EVIDENCIA_INCIDENTE, detalles_tecnicos),
            Accion_sugerida=(COL_ACCION_SUGERIDA, primer_valor_util),
            Detalle_tecnico_observado=(TEXT_CAUSA_TECNICA, detalles_tecnicos),
        )
        .reset_index()
        .sort_values(by=[TEXT_CANTIDAD, COL_CAUSA_RAIZ], ascending=[False, True])
    )
    resumen[porcentaje_columna] = resumen[TEXT_CANTIDAD].apply(lambda valor: porcentaje(valor, total))
    resumen = resumen.rename(
        columns={
            "Lectura_ejecutiva": COL_LECTURA_EJECUTIVA,
            "Evidencia_observada": COL_EVIDENCIA_INCIDENTE,
            "Accion_sugerida": COL_ACCION_SUGERIDA,
            "Detalle_tecnico_observado": "Detalle tecnico observado",
        }
    )
    return resumen[columnas]


def resumen_causas_cliente_externo(df):
    return resumen_causas_incidentes(df, "% cliente externo")


def causas_relevantes_incidentes(causas):
    if causas.empty:
        return causas
    excluidas = [
        "Casos repetidos o registros duplicados que distorsionan la lectura operativa.",
        "Hallazgos tecnicos con bajo volumen o descripcion no estandarizada.",
        "No hay informacion suficiente para explicar la causa.",
    ]
    relevantes = causas[~causas[COL_LECTURA_EJECUTIVA].isin(excluidas)].copy()
    return relevantes[relevantes[TEXT_CANTIDAD] >= 2].copy()


def causas_relevantes_cliente_externo(causas):
    return causas_relevantes_incidentes(causas)


def resumen_motivos_caso_cliente_externo(df):
    columnas = [COL_MOTIVO_CASO, TEXT_CANTIDAD, "% casos"]
    if df.empty:
        return pd.DataFrame(columns=columnas)
    trabajo = df.copy()
    trabajo[COL_MOTIVO_CASO] = (
        trabajo[TEXT_CAUSA_RAIZ_AUTO]
        .replace("", pd.NA)
        .fillna("Otro caso cliente externo")
    )
    total = len(trabajo)
    resumen = (
        trabajo.groupby(COL_MOTIVO_CASO, dropna=False)
        .size()
        .reset_index(name=TEXT_CANTIDAD)
        .sort_values(by=[TEXT_CANTIDAD, COL_MOTIVO_CASO], ascending=[False, True])
    )
    resumen["% casos"] = resumen[TEXT_CANTIDAD].apply(lambda valor: porcentaje(valor, total))
    return resumen[columnas]


def render_causas_incidentes(df, titulo, porcentaje_columna):
    causas = resumen_causas_incidentes(df, porcentaje_columna)
    relevantes = causas_relevantes_incidentes(causas)
    if relevantes.empty:
        st.info(f"No hay causas raiz relevantes para {titulo.lower()} en el periodo.")
        return

    grafico = (
        relevantes.groupby(COL_LECTURA_EJECUTIVA, as_index=False)
        .agg(Cantidad=(TEXT_CANTIDAD, "sum"))
        .sort_values(by=TEXT_CANTIDAD, ascending=True)
    )
    fig = px.bar(
        grafico,
        x=TEXT_CANTIDAD,
        y=COL_LECTURA_EJECUTIVA,
        orientation="h",
        text=TEXT_CANTIDAD,
        color_discrete_sequence=[UI_PALETTE[TEXT_PURPLE]],
    )
    fig.update_traces(textposition=TEXT_OUTSIDE)
    fig = aplicar_estilo_figura(fig, titulo)
    fig.update_layout(height=max(320, 56 * len(grafico)),yaxis={"automargin": True})
    st.plotly_chart(fig, use_container_width=True)
    st.dataframe(relevantes, use_container_width=True, hide_index=True)


def dashboard_casos():
    anio, mes, periodo_label = selector_periodo_sql("cases", "dashboard_casos_periodo")
    if not periodo_sql_valido(anio, "casos"):
        return
    df = cargar_casos_soporte_filtrados_cache(anio, mes)
    if df.empty:
        st.info(f"No hay casos cargados para {periodo_label}.")
        return

    df = preparar_fechas_dashboard(df)
    if df.empty:
        st.info(f"No hay casos cargados para {periodo_label}.")
        return

    total = len(df)
    cerrados_mask = mascara_cerrados(df)
    cerrados = len(df[cerrados_mask])
    abiertos = total - cerrados

    df_cerrados = df[cerrados_mask]
    tiempos_cerrados = pd.to_numeric(df_cerrados[TEXT_TIEMPO_RESPUESTA], errors=TEXT_COERCE).dropna()
    promedio = round(tiempos_cerrados.mean(), 2) if len(tiempos_cerrados) > 0 else 0

    total_cerrados = len(tiempos_cerrados)
    cumplen = len(tiempos_cerrados[tiempos_cerrados < SLA_CASOS_HORAS])
    porcentaje_sla = round((cumplen / total_cerrados) * 100, 2) if total_cerrados > 0 else 0
    incumplen = total_cerrados - cumplen

    render_tarjetas(
        [
            ("Total Casos", total),
            (TEXT_CERRADOS, cerrados),
            (TEXT_ABIERTOS, abiertos),
            ("Promedio (h)", promedio),
            (f"SLA <{SLA_CASOS_HORAS}h (%)", f"{porcentaje_sla}%"),
        ]
    )
    st.caption(f"{TEXT_PERIODO}{periodo_label} | Cumplen: {cumplen}{TEXT_NO_CUMPLEN}{incumplen}")

    st.divider()
    render_distribucion_productos_soporte(df, periodo_label)

    st.divider()
    render_carga_agentes(df, TEXT_ASIGNADO, "Carga por agente - casos", TEXT_CASOS)

    st.divider()
    col1, col2 = st.columns(2)

    with col1:
        tip = resumen_tipologias_soporte_casos(df)
        grafico_porcentaje_tipologias_soporte(tip)

    with col2:
        serie = df.copy()
        casos_dia = serie.groupby(serie[TEXT_CREADO_DT_DASHBOARD].dt.date).size().reset_index(name=TEXT_CASOS_2)
        casos_dia.columns = [TEXT_FECHA, TEXT_CASOS_2]
        fig = px.bar(casos_dia, x=TEXT_FECHA, y=TEXT_CASOS_2, color_discrete_sequence=[UI_PALETTE[TEXT_YELLOW]])
        fig.update_traces(marker_color=UI_PALETTE[TEXT_YELLOW])
        st.plotly_chart(aplicar_estilo_figura(fig, "Casos por dia"), use_container_width=True)

    st.divider()
    with st.expander("Analisis de agendamiento con historico"):
        st.caption("Este bloque puede consultar mas datos. Se carga solo cuando lo solicitas.")
        if st.button("Calcular analisis de agendamiento", key="calcular_agendamiento_dashboard_casos"):
            historico = preparar_fechas_dashboard(cargar_casos_soporte_cache())
            render_analisis_agendamiento_mesa(df, historico, periodo_key_sql(anio, mes))

    st.divider()
    st.subheader("Resumen por tipologia de casos")
    st.caption("Agrupacion ejecutiva de casos. La tipificacion original se mantiene como referencia.")
    st.dataframe(resumen_tipologias_soporte_casos(df), use_container_width=True, hide_index=True)
    with st.expander("Ver tipificaciones originales"):
        st.dataframe(tabla_resumen_tipificaciones_casos(df), use_container_width=True, hide_index=True)

    render_seguimiento_casos(df)


def dashboard_kpi_casos_cliente_externo():
    vista = st.radio(
        "Vista",
        ["KPI actual", "Por mes"],
        horizontal=True,
        key="kpi_casos_cliente_externo_vista",
    )
    if vista == "Por mes":
        render_kpi_casos_cliente_externo_comparativo()
        return

    if vista == "KPI actual":
        anio, mes, periodo_label = selector_periodo_sql("cases", "kpi_casos_cliente_externo_periodo")
        if periodo_sql_valido(anio, "casos"):
            df = cargar_casos_soporte_filtrados_cache(anio, mes)
            if df.empty:
                st.info(f"No hay casos cargados para {periodo_label}.")
            else:
                df = preparar_fechas_dashboard(df)
                if df.empty:
                    st.info(f"No hay casos cargados para {periodo_label}.")
                else:
                    render_kpi_casos_cliente_externo(df, periodo_label)


def dashboard_kpi_incidentes():
    anio, mes, periodo_label = selector_periodo_sql("incidents", "kpi_incidentes_periodo")
    if not periodo_sql_valido(anio, "incidentes"):
        return
    df = cargar_incidentes_filtrados_cache(anio, mes)
    if df.empty:
        st.info(f"No hay incidentes cargados para {periodo_label}.")
        return

    df = preparar_fechas_dashboard(df)
    if df.empty:
        st.info(f"No hay incidentes cargados para {periodo_label}.")
        return

    render_kpi_incidentes(df, periodo_label)


def filtrar_anio_dashboard(df, anio, columna_dt=TEXT_CREADO_DT_DASHBOARD):
    if df.empty or columna_dt not in df.columns:
        return df.copy()
    return df[df[columna_dt].dt.year == int(anio)].copy()


def filtrar_rango_dashboard(df, fecha_inicio, fecha_fin, columna_dt=TEXT_CREADO_DT_DASHBOARD):
    if df.empty or columna_dt not in df.columns:
        return df.copy()
    inicio = pd.Timestamp(fecha_inicio).normalize()
    fin = pd.Timestamp(fecha_fin).normalize() + pd.Timedelta(days=1) - pd.Timedelta(microseconds=1)
    fechas = df[columna_dt]
    return df[(fechas >= inicio) & (fechas <= fin)].copy()


def anios_disponibles_kpi_comparativo(*tablas):
    anios = {2025, 2026}
    for tabla in tablas:
        if tabla.empty or TEXT_CREADO_DT_DASHBOARD not in tabla.columns:
            continue
        anios.update(tabla[TEXT_CREADO_DT_DASHBOARD].dropna().dt.year.astype(int).tolist())
    return sorted(anios)


def preparar_casos_kpi_comparativo_ligero(casos):
    if casos.empty:
        return casos.copy()
    trabajo = normalizar_tipificaciones_casos_df(casos)
    trabajo = preparar_fechas_dashboard(trabajo)
    trabajo = trabajo.dropna(subset=[TEXT_CREADO_DT_DASHBOARD])
    trabajo[TEXT_CERRADO_2] = mascara_cerrados(trabajo)
    trabajo[TEXT_ABIERTO] = ~trabajo[TEXT_CERRADO_2]
    trabajo["_tiempo_eval_sla_h"] = pd.to_numeric(
        trabajo.get(TEXT_TIEMPO_RESPUESTA, pd.Series(dtype=TEXT_FLOAT)),
        errors=TEXT_COERCE,
    )
    trabajo["Cumple SLA <=36h"] = trabajo["_tiempo_eval_sla_h"].apply(
        lambda valor: "Si" if pd.notna(valor) and valor <= SLA_CASOS_HORAS else "No"
    )
    return trabajo


def asegurar_columna(df, columna, default=""):
    if columna not in df.columns:
        df[columna] = default
    return df


def preparar_incidentes_kpi_comparativo_ligero(incidentes):
    if incidentes.empty:
        return incidentes.copy()
    trabajo = preparar_fechas_dashboard(incidentes)
    trabajo = trabajo.dropna(subset=[TEXT_CREADO_DT_DASHBOARD])
    for columna in [TEXT_TIPIFICACION_AUTO, TEXT_TIPO_INCIDENTE_AUTO, TEXT_PRIORIDAD, TEXT_DURACION_HORAS, "duracion_segundos"]:
        asegurar_columna(trabajo, columna)

    tipificacion = trabajo[TEXT_TIPIFICACION_AUTO].fillna("")
    tiene_tipificacion_real = tipificacion.isin(
        [TIPIFICACION_INCIDENTE_CLIENTE_EXTERNO, TIPIFICACION_INCIDENTE_INTERNO]
    ).any()
    if tiene_tipificacion_real:
        trabajo = trabajo[
            tipificacion.isin([TIPIFICACION_INCIDENTE_CLIENTE_EXTERNO, TIPIFICACION_INCIDENTE_INTERNO])
        ].copy()
    trabajo[TEXT_CERRADO_2] = mascara_cerrados(trabajo)
    trabajo[TEXT_ABIERTO] = ~trabajo[TEXT_CERRADO_2]
    trabajo[TEXT_SEGMENTO] = trabajo[TEXT_TIPIFICACION_AUTO].apply(segmento_incidente)
    trabajo[TEXT_PRIORIDAD_NORMALIZADA] = trabajo[TEXT_PRIORIDAD].apply(normalizar_prioridad_incidente)
    trabajo["familia_sla"] = trabajo.apply(
        lambda row: familia_sla_incidente(row[TEXT_TIPIFICACION_AUTO], row[TEXT_TIPO_INCIDENTE_AUTO]),
        axis=1,
    )
    trabajo[TEXT_SLA_OBJETIVO_HORAS] = trabajo.apply(sla_objetivo_horas_incidente, axis=1)
    trabajo[TEXT_DURACION_SLA_HORAS] = trabajo.apply(duracion_sla_horas_incidente, axis=1)
    trabajo[TEXT_ESTADO_SLA] = trabajo.apply(estado_sla_incidente, axis=1)
    trabajo[TEXT_DURACION_HORAS_NUM] = pd.to_numeric(trabajo[TEXT_DURACION_SLA_HORAS], errors=TEXT_COERCE)
    return trabajo


def preparar_bases_kpi_comparativo(casos, incidentes):
    return (
        preparar_casos_kpi_comparativo_ligero(casos),
        preparar_incidentes_kpi_comparativo_ligero(incidentes),
    )


def metricas_casos_comparativo(base, anio):
    datos = filtrar_anio_dashboard(base, anio)
    if datos.empty:
        return {
            "Registro": TEXT_CASOS,
            "Anio": anio,
            TEXT_TOTAL: 0,
            TEXT_CERRADOS: 0,
            TEXT_ABIERTOS: 0,
            "SLA %": 0,
            "Cumple SLA": 0,
            "No cumple SLA": 0,
            COL_PROM_HORAS: 0,
        }
    cerrados = int(datos[TEXT_CERRADO_2].sum()) if TEXT_CERRADO_2 in datos.columns else 0
    cerrados_sla = datos[datos[TEXT_CERRADO_2]].copy() if TEXT_CERRADO_2 in datos.columns else pd.DataFrame()
    tiempos = pd.to_numeric(
        cerrados_sla.get("_tiempo_eval_sla_h", pd.Series(dtype=TEXT_FLOAT)),
        errors=TEXT_COERCE,
    ).dropna()
    cumple = int((tiempos <= SLA_CASOS_HORAS).sum())
    no_cumple = len(tiempos) - cumple
    return {
        "Registro": TEXT_CASOS,
        "Anio": anio,
        TEXT_TOTAL: len(datos),
        TEXT_CERRADOS: cerrados,
        TEXT_ABIERTOS: len(datos) - cerrados,
        "SLA %": porcentaje(cumple, len(tiempos)),
        "Cumple SLA": cumple,
        "No cumple SLA": no_cumple,
        COL_PROM_HORAS: round(tiempos.mean(), 2) if not tiempos.empty else 0,
    }


def metricas_incidentes_comparativo(base, anio):
    datos = filtrar_anio_dashboard(base, anio)
    if datos.empty:
        return {
            "Registro": TEXT_INCIDENTES,
            "Anio": anio,
            TEXT_TOTAL: 0,
            TEXT_CERRADOS: 0,
            TEXT_ABIERTOS: 0,
            "SLA %": 0,
            "Cumple SLA": 0,
            "No cumple SLA": 0,
            COL_PROM_HORAS: 0,
            "Reincidentes": 0,
        }
    cerrados = int(datos[TEXT_CERRADO_2].sum()) if TEXT_CERRADO_2 in datos.columns else 0
    cerrados_sla = datos[
        datos.get(TEXT_CERRADO_2, pd.Series(False, index=datos.index))
        & datos[TEXT_ESTADO_SLA].isin([ESTADO_SLA_CUMPLE, ESTADO_SLA_NO_CUMPLE])
    ]
    cumple = int((cerrados_sla[TEXT_ESTADO_SLA] == ESTADO_SLA_CUMPLE).sum())
    no_cumple = int((cerrados_sla[TEXT_ESTADO_SLA] == ESTADO_SLA_NO_CUMPLE).sum())
    tiempos = pd.to_numeric(datos.get(TEXT_DURACION_HORAS_NUM, pd.Series(dtype=TEXT_FLOAT)), errors=TEXT_COERCE).dropna()
    reincidentes = int(datos.get(TEXT_REINCIDENTE, pd.Series(False, index=datos.index)).sum())
    return {
        "Registro": TEXT_INCIDENTES,
        "Anio": anio,
        TEXT_TOTAL: len(datos),
        TEXT_CERRADOS: cerrados,
        TEXT_ABIERTOS: len(datos) - cerrados,
        "SLA %": porcentaje(cumple, cumple + no_cumple),
        "Cumple SLA": cumple,
        "No cumple SLA": no_cumple,
        COL_PROM_HORAS: round(tiempos.mean(), 2) if not tiempos.empty else 0,
        "Reincidentes": reincidentes,
    }


def tabla_metricas_kpi_comparativo(base_casos, base_incidentes, anios):
    filas = []
    for anio in anios:
        filas.append(metricas_casos_comparativo(base_casos, anio))
        filas.append(metricas_incidentes_comparativo(base_incidentes, anio))
    return pd.DataFrame(filas)


def metricas_casos_comparativo_rango(base, etiqueta, fecha_inicio, fecha_fin):
    datos = filtrar_rango_dashboard(base, fecha_inicio, fecha_fin)
    if datos.empty:
        return {
            "Registro": TEXT_CASOS,
            "Periodo": etiqueta,
            TEXT_TOTAL: 0,
            TEXT_CERRADOS: 0,
            TEXT_ABIERTOS: 0,
            "SLA %": 0,
            "Cumple SLA": 0,
            "No cumple SLA": 0,
            COL_PROM_HORAS: 0,
        }

    cerrados = int(datos[TEXT_CERRADO_2].sum()) if TEXT_CERRADO_2 in datos.columns else 0
    cerrados_sla = datos[datos[TEXT_CERRADO_2]].copy() if TEXT_CERRADO_2 in datos.columns else pd.DataFrame()
    tiempos = pd.to_numeric(
        cerrados_sla.get("_tiempo_eval_sla_h", pd.Series(dtype=TEXT_FLOAT)),
        errors=TEXT_COERCE,
    ).dropna()
    cumple = int((tiempos <= SLA_CASOS_HORAS).sum())
    no_cumple = len(tiempos) - cumple
    return {
        "Registro": TEXT_CASOS,
        "Periodo": etiqueta,
        TEXT_TOTAL: len(datos),
        TEXT_CERRADOS: cerrados,
        TEXT_ABIERTOS: len(datos) - cerrados,
        "SLA %": porcentaje(cumple, len(tiempos)),
        "Cumple SLA": cumple,
        "No cumple SLA": no_cumple,
        COL_PROM_HORAS: round(tiempos.mean(), 2) if not tiempos.empty else 0,
    }


def metricas_incidentes_comparativo_rango(base, etiqueta, fecha_inicio, fecha_fin):
    datos = filtrar_rango_dashboard(base, fecha_inicio, fecha_fin)
    if datos.empty:
        return {
            "Registro": TEXT_INCIDENTES,
            "Periodo": etiqueta,
            TEXT_TOTAL: 0,
            TEXT_CERRADOS: 0,
            TEXT_ABIERTOS: 0,
            "SLA %": 0,
            "Cumple SLA": 0,
            "No cumple SLA": 0,
            COL_PROM_HORAS: 0,
            "Reincidentes": 0,
        }

    cerrados = int(datos[TEXT_CERRADO_2].sum()) if TEXT_CERRADO_2 in datos.columns else 0
    cerrados_sla = datos[
        datos.get(TEXT_CERRADO_2, pd.Series(False, index=datos.index))
        & datos[TEXT_ESTADO_SLA].isin([ESTADO_SLA_CUMPLE, ESTADO_SLA_NO_CUMPLE])
    ]
    cumple = int((cerrados_sla[TEXT_ESTADO_SLA] == ESTADO_SLA_CUMPLE).sum())
    no_cumple = int((cerrados_sla[TEXT_ESTADO_SLA] == ESTADO_SLA_NO_CUMPLE).sum())
    tiempos = pd.to_numeric(datos.get(TEXT_DURACION_HORAS_NUM, pd.Series(dtype=TEXT_FLOAT)), errors=TEXT_COERCE).dropna()
    reincidentes = int(datos.get(TEXT_REINCIDENTE, pd.Series(False, index=datos.index)).sum())
    return {
        "Registro": TEXT_INCIDENTES,
        "Periodo": etiqueta,
        TEXT_TOTAL: len(datos),
        TEXT_CERRADOS: cerrados,
        TEXT_ABIERTOS: len(datos) - cerrados,
        "SLA %": porcentaje(cumple, cumple + no_cumple),
        "Cumple SLA": cumple,
        "No cumple SLA": no_cumple,
        COL_PROM_HORAS: round(tiempos.mean(), 2) if not tiempos.empty else 0,
        "Reincidentes": reincidentes,
    }


def tabla_metricas_kpi_comparativo_rangos(base_casos, base_incidentes, rangos):
    filas = []
    for etiqueta, fecha_inicio, fecha_fin in rangos:
        filas.append(metricas_casos_comparativo_rango(base_casos, etiqueta, fecha_inicio, fecha_fin))
        filas.append(metricas_incidentes_comparativo_rango(base_incidentes, etiqueta, fecha_inicio, fecha_fin))
    return pd.DataFrame(filas)


def variacion_porcentual(actual, base):
    if not base:
        return None
    return round(((actual - base) / base) * 100, 2)


def texto_variacion(actual, base):
    if not base and actual:
        return "Sin base anterior"
    if actual > base:
        return "Aumento"
    if actual < base:
        return "Disminucion"
    return "Sin cambio"


def tabla_comparativo_anios(metricas, anio_base, anio_comparado):
    filas = []
    for registro in [TEXT_CASOS, TEXT_INCIDENTES]:
        base = metricas[(metricas["Registro"] == registro) & (metricas["Anio"] == anio_base)]
        actual = metricas[(metricas["Registro"] == registro) & (metricas["Anio"] == anio_comparado)]
        if base.empty or actual.empty:
            continue
        base = base.iloc[0]
        actual = actual.iloc[0]
        filas.append(
            {
                "Registro": registro,
                f"Total {anio_base}": int(base[TEXT_TOTAL]),
                f"Total {anio_comparado}": int(actual[TEXT_TOTAL]),
                "Diferencia total": int(actual[TEXT_TOTAL] - base[TEXT_TOTAL]),
                "Variacion total %": variacion_porcentual(actual[TEXT_TOTAL], base[TEXT_TOTAL]),
                f"SLA {anio_base} %": base["SLA %"],
                f"SLA {anio_comparado} %": actual["SLA %"],
                "Diferencia SLA p.p.": round(actual["SLA %"] - base["SLA %"], 2),
                f"Abiertos {anio_base}": int(base[TEXT_ABIERTOS]),
                f"Abiertos {anio_comparado}": int(actual[TEXT_ABIERTOS]),
                "Lectura": texto_variacion(actual[TEXT_TOTAL], base[TEXT_TOTAL]),
            }
        )
    return pd.DataFrame(filas)


def tabla_comparativo_rangos(metricas, periodo_base, periodo_comparado):
    filas = []
    for registro in [TEXT_CASOS, TEXT_INCIDENTES]:
        base = metricas[(metricas["Registro"] == registro) & (metricas["Periodo"] == periodo_base)]
        actual = metricas[(metricas["Registro"] == registro) & (metricas["Periodo"] == periodo_comparado)]
        if base.empty or actual.empty:
            continue
        base = base.iloc[0]
        actual = actual.iloc[0]
        filas.append(
            {
                "Registro": registro,
                "Total base": int(base[TEXT_TOTAL]),
                "Total comparado": int(actual[TEXT_TOTAL]),
                "Diferencia total": int(actual[TEXT_TOTAL] - base[TEXT_TOTAL]),
                "Variacion total %": variacion_porcentual(actual[TEXT_TOTAL], base[TEXT_TOTAL]),
                "SLA base %": base["SLA %"],
                "SLA comparado %": actual["SLA %"],
                "Diferencia SLA p.p.": round(actual["SLA %"] - base["SLA %"], 2),
                "Abiertos base": int(base[TEXT_ABIERTOS]),
                "Abiertos comparado": int(actual[TEXT_ABIERTOS]),
                "Lectura": texto_variacion(actual[TEXT_TOTAL], base[TEXT_TOTAL]),
            }
        )
    return pd.DataFrame(filas)


def valor_comparativo_rango(comparativo, registro, columna):
    filas = comparativo[comparativo["Registro"] == registro]
    return filas.iloc[0][columna] if not filas.empty else 0


def render_tarjetas_kpi_comparativo_rangos(comparativo):
    render_tarjetas(
        [
            ("Casos base", valor_comparativo_rango(comparativo, TEXT_CASOS, "Total base")),
            ("Casos comparado", valor_comparativo_rango(comparativo, TEXT_CASOS, "Total comparado")),
            ("Var casos", f"{valor_comparativo_rango(comparativo, TEXT_CASOS, 'Diferencia total'):+g}"),
            ("Incidentes base", valor_comparativo_rango(comparativo, TEXT_INCIDENTES, "Total base")),
            ("Incidentes comparado", valor_comparativo_rango(comparativo, TEXT_INCIDENTES, "Total comparado")),
            ("Var incidentes", f"{valor_comparativo_rango(comparativo, TEXT_INCIDENTES, 'Diferencia total'):+g}"),
        ]
    )


def tendencia_diaria_kpi_rango(base, tipo, rangos):
    columnas = ["Periodo", "Dia", "Fecha", "Tipo", TEXT_CANTIDAD]
    if base.empty or TEXT_CREADO_DT_DASHBOARD not in base.columns:
        return pd.DataFrame(columns=columnas)

    filas = []
    for etiqueta, fecha_inicio, fecha_fin in rangos:
        inicio = pd.Timestamp(fecha_inicio).normalize()
        fin = pd.Timestamp(fecha_fin).normalize()
        datos = filtrar_rango_dashboard(base, inicio, fin)
        conteo = (
            datos.groupby(datos[TEXT_CREADO_DT_DASHBOARD].dt.normalize()).size()
            if not datos.empty
            else pd.Series(dtype=TEXT_FLOAT)
        )
        for fecha in pd.date_range(inicio, fin, freq="D"):
            filas.append(
                {
                    "Periodo": etiqueta,
                    "Dia": int((fecha - inicio).days) + 1,
                    "Fecha": fecha.date().isoformat(),
                    "Tipo": tipo,
                    TEXT_CANTIDAD: int(conteo.get(fecha, 0)),
                }
            )
    return pd.DataFrame(filas, columns=columnas)


def render_graficas_kpi_comparativo_rangos(metricas, tendencia):
    if metricas.empty:
        st.info("No hay metricas para graficar.")
        return
    
    # Obtener los periodos Ãºnicos
    periodos = metricas["Periodo"].unique().tolist()
    
    # GrÃ¡ficos de Total por rango
    st.markdown("---")
    st.markdown("<h2 style='text-align: center; color: #f35b04;'>ðŸ“Š AnÃ¡lisis de Casos e Incidentes</h2>", unsafe_allow_html=True)
    
    col_casos, col_incidentes = st.columns(2)
    
    with col_casos:
        st.markdown("<h3 style='text-align: center;'>DistribuciÃ³n de Casos</h3>", unsafe_allow_html=True)
        datos_casos = metricas[metricas["Registro"] == TEXT_CASOS].copy()
        if not datos_casos.empty:
            fig = px.pie(
                datos_casos,
                values=TEXT_TOTAL,
                names="Periodo",
                color_discrete_sequence=[UI_PALETTE[TEXT_PRIMARY], UI_PALETTE[TEXT_LAVENDER]],
                hole=0.35,
            )
            fig.update_traces(
                textposition="inside", 
                textinfo="label+value+percent",
                textfont=dict(size=13, color="white"),
                hovertemplate="<b>%{label}</b><br>Total: %{value}<br>Porcentaje: %{percent}<extra></extra>"
            )
            fig.update_layout(
                height=450,
                showlegend=True,
                font=dict(size=12),
                margin=dict(t=20, b=20, l=20, r=20),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
            )
            st.plotly_chart(fig, use_container_width=True)
    
    with col_incidentes:
        st.markdown("<h3 style='text-align: center;'>DistribuciÃ³n de Incidentes</h3>", unsafe_allow_html=True)
        datos_incidentes = metricas[metricas["Registro"] == TEXT_INCIDENTES].copy()
        if not datos_incidentes.empty:
            fig = px.pie(
                datos_incidentes,
                values=TEXT_TOTAL,
                names="Periodo",
                color_discrete_sequence=[UI_PALETTE[TEXT_PRIMARY], UI_PALETTE[TEXT_LAVENDER]],
                hole=0.35,
            )
            fig.update_traces(
                textposition="inside", 
                textinfo="label+value+percent",
                textfont=dict(size=13, color="white"),
                hovertemplate="<b>%{label}</b><br>Total: %{value}<br>Porcentaje: %{percent}<extra></extra>"
            )
            fig.update_layout(
                height=450,
                showlegend=True,
                font=dict(size=12),
                margin=dict(t=20, b=20, l=20, r=20),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
            )
            st.plotly_chart(fig, use_container_width=True)
    
    # Tarjetas de Cumplimiento SLA
    st.markdown("---")
    
    # Calcular porcentajes de SLA
    tarjetas_sla = []
    
    # SLA Casos
    datos_casos = metricas[metricas["Registro"] == TEXT_CASOS].copy()
    if not datos_casos.empty:
        for _, row in datos_casos.iterrows():
            cumple = row.get("Cumple SLA", 0)
            no_cumple = row.get("No cumple SLA", 0)
            total = cumple + no_cumple
            sla_pct = (cumple / total * 100) if total > 0 else 0
            periodo = row["Periodo"]
            tarjetas_sla.append((f"SLA Casos {periodo}", f"{sla_pct:.1f}%"))
    
    # SLA Incidentes
    datos_incidentes = metricas[metricas["Registro"] == TEXT_INCIDENTES].copy()
    if not datos_incidentes.empty:
        for _, row in datos_incidentes.iterrows():
            cumple = row.get("Cumple SLA", 0)
            no_cumple = row.get("No cumple SLA", 0)
            total = cumple + no_cumple
            sla_pct = (cumple / total * 100) if total > 0 else 0
            periodo = row["Periodo"]
            tarjetas_sla.append((f"SLA Incidentes {periodo}", f"{sla_pct:.1f}%"))
    
    if tarjetas_sla:
        render_tarjetas(tarjetas_sla)

    st.divider()
    st.subheader("Tendencia diaria comparativa")
    if tendencia.empty:
        st.info("No hay registros diarios para los rangos seleccionados.")
        return
    for tipo in [TEXT_CASOS, TEXT_INCIDENTES]:
        datos_tipo = tendencia[tendencia["Tipo"] == tipo].copy()
        if datos_tipo.empty:
            st.info(f"No hay datos diarios de {tipo.lower()}.")
            continue
        fig = px.line(
            datos_tipo,
            x="Dia",
            y=TEXT_CANTIDAD,
            color="Periodo",
            markers=True,
            hover_data={"Fecha": True, TEXT_CANTIDAD: True, "Dia": True},
            color_discrete_sequence=[UI_PALETTE[TEXT_PRIMARY], UI_PALETTE[TEXT_LAVENDER]],
        )
        st.plotly_chart(aplicar_estilo_figura(fig, f"{tipo} por dia del rango"), use_container_width=True)


def fechas_disponibles_comparativo_rango(*tablas):
    fechas = []
    for tabla in tablas:
        if tabla.empty or TEXT_CREADO_DT_DASHBOARD not in tabla.columns:
            continue
        serie = tabla[TEXT_CREADO_DT_DASHBOARD].dropna()
        if not serie.empty:
            fechas.append(serie.min())
            fechas.append(serie.max())
    if not fechas:
        return None, None
    return min(fechas).date(), max(fechas).date()


def rangos_default_comparativo(fecha_min, fecha_max):
    inicio_global = pd.Timestamp(fecha_min).normalize()
    fin_global = pd.Timestamp(fecha_max).normalize()
    total_dias = max(int((fin_global - inicio_global).days) + 1, 1)
    ventana = max(1, min(30, total_dias // 2 or 1))

    comparado_fin = fin_global
    comparado_inicio = max(inicio_global, comparado_fin - pd.Timedelta(days=ventana - 1))
    base_fin = comparado_inicio - pd.Timedelta(days=1)
    if base_fin < inicio_global:
        base_fin = comparado_inicio
    base_inicio = max(inicio_global, base_fin - pd.Timedelta(days=ventana - 1))
    return (
        (base_inicio.date(), base_fin.date()),
        (comparado_inicio.date(), comparado_fin.date()),
    )


def normalizar_rango_fecha_input(rango):
    if isinstance(rango, tuple) and len(rango) == 2:
        return rango[0], rango[1]
    if isinstance(rango, list) and len(rango) == 2:
        return rango[0], rango[1]
    return None


def etiqueta_rango_fechas(fecha_inicio, fecha_fin):
    ts_inicio = pd.Timestamp(fecha_inicio)
    ts_fin = pd.Timestamp(fecha_fin)
    meses_es = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 
                'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
    
    mes_inicio = meses_es[ts_inicio.month - 1]
    mes_fin = meses_es[ts_fin.month - 1]
    anio_inicio = ts_inicio.year
    anio_fin = ts_fin.year
    
    # Si es el mismo mes y aÃ±o
    if ts_inicio.month == ts_fin.month and anio_inicio == anio_fin:
        return f"{mes_inicio} {anio_inicio}"
    
    # Si es el mismo aÃ±o pero diferente mes
    if anio_inicio == anio_fin:
        return f"{mes_inicio} - {mes_fin} {anio_fin}"
    
    # Si son aÃ±os diferentes
    return f"{mes_inicio} {anio_inicio} - {mes_fin} {anio_fin}"


def selector_rangos_kpi_comparativo(fecha_min, fecha_max, key_prefix="kpi_comparativo"):
    rango_base_default, rango_comparado_default = rangos_default_comparativo(fecha_min, fecha_max)
    col_base, col_comparado = st.columns(2)
    with col_base:
        rango_base = st.date_input(
            "Rango base",
            value=rango_base_default,
            min_value=fecha_min,
            max_value=fecha_max,
            key=f"{key_prefix}_rango_base",
        )
    with col_comparado:
        rango_comparado = st.date_input(
            "Rango comparado",
            value=rango_comparado_default,
            min_value=fecha_min,
            max_value=fecha_max,
            key=f"{key_prefix}_rango_comparado",
        )

    rango_base = normalizar_rango_fecha_input(rango_base)
    rango_comparado = normalizar_rango_fecha_input(rango_comparado)
    if not rango_base or not rango_comparado:
        st.warning("Selecciona fecha inicial y final para ambos rangos.")
        return None

    base_inicio, base_fin = rango_base
    comparado_inicio, comparado_fin = rango_comparado
    if pd.Timestamp(base_inicio) > pd.Timestamp(base_fin) or pd.Timestamp(comparado_inicio) > pd.Timestamp(comparado_fin):
        st.warning("La fecha inicial debe ser menor o igual a la fecha final en cada rango.")
        return None

    return [
        ("Base", base_inicio, base_fin),
        ("Comparado", comparado_inicio, comparado_fin),
    ]


def tendencia_mensual_kpi(base, tipo, anios):
    columnas = ["Anio", "Mes", "Tipo", TEXT_CANTIDAD]
    if base.empty or TEXT_CREADO_DT_DASHBOARD not in base.columns:
        return pd.DataFrame(columns=columnas)
    trabajo = base[base[TEXT_CREADO_DT_DASHBOARD].dt.year.isin(anios)].copy()
    if trabajo.empty:
        return pd.DataFrame(columns=columnas)
    trabajo["Anio"] = trabajo[TEXT_CREADO_DT_DASHBOARD].dt.year.astype(str)
    trabajo["Mes"] = trabajo[TEXT_CREADO_DT_DASHBOARD].dt.month
    resumen = trabajo.groupby(["Anio", "Mes"]).size().reset_index(name=TEXT_CANTIDAD)
    resumen["Tipo"] = tipo
    return resumen[columnas]


def render_tarjetas_kpi_comparativo(comparativo, anio_base, anio_comparado):
    def valor_registro(registro, columna):
        filas = comparativo[comparativo["Registro"] == registro]
        return filas.iloc[0][columna] if not filas.empty else 0

    render_tarjetas(
        [
            (f"Casos {anio_base}", valor_registro(TEXT_CASOS, f"Total {anio_base}")),
            (f"Casos {anio_comparado}", valor_registro(TEXT_CASOS, f"Total {anio_comparado}")),
            ("Var casos", f"{valor_registro(TEXT_CASOS, 'Diferencia total'):+g}"),
            (f"Incidentes {anio_base}", valor_registro(TEXT_INCIDENTES, f"Total {anio_base}")),
            (f"Incidentes {anio_comparado}", valor_registro(TEXT_INCIDENTES, f"Total {anio_comparado}")),
            ("Var incidentes", f"{valor_registro(TEXT_INCIDENTES, 'Diferencia total'):+g}"),
        ]
    )


def render_graficas_kpi_comparativo(metricas, tendencia, anios):
    col_totales, col_sla = st.columns(2)
    with col_totales:
        if metricas.empty:
            st.info("No hay metricas para graficar.")
        else:
            fig = px.bar(
                metricas,
                x="Registro",
                y=TEXT_TOTAL,
                color="Anio",
                barmode="group",
                text=TEXT_TOTAL,
                color_discrete_sequence=[UI_PALETTE[TEXT_PRIMARY], UI_PALETTE[TEXT_LAVENDER]],
            )
            fig.update_traces(textposition=TEXT_OUTSIDE)
            st.plotly_chart(aplicar_estilo_figura(fig, "Total por anio"), use_container_width=True)

    with col_sla:
        if metricas.empty:
            st.info("No hay SLA para graficar.")
        else:
            fig = px.bar(
                metricas,
                x="Registro",
                y="SLA %",
                color="Anio",
                barmode="group",
                text="SLA %",
                color_discrete_sequence=[UI_PALETTE[TEXT_PRIMARY], UI_PALETTE[TEXT_LAVENDER]],
            )
            fig.update_traces(textposition=TEXT_OUTSIDE)
            fig.update_yaxes(range=[0, 100])
            st.plotly_chart(aplicar_estilo_figura(fig, "Cumplimiento SLA"), use_container_width=True)

    st.divider()
    st.subheader("Tendencia mensual")
    if tendencia.empty:
        st.info("No hay registros mensuales para los anios seleccionados.")
        return
    for tipo in [TEXT_CASOS, TEXT_INCIDENTES]:
        datos_tipo = tendencia[tendencia["Tipo"] == tipo].copy()
        if datos_tipo.empty:
            st.info(f"No hay datos mensuales de {tipo.lower()}.")
            continue
        fig = px.line(
            datos_tipo,
            x="Mes",
            y=TEXT_CANTIDAD,
            color="Anio",
            markers=True,
            color_discrete_sequence=[UI_PALETTE[TEXT_PRIMARY], UI_PALETTE[TEXT_LAVENDER]],
        )
        fig.update_xaxes(tickmode="array", tickvals=list(range(1, 13)))
        st.plotly_chart(aplicar_estilo_figura(fig, f"{tipo} por mes"), use_container_width=True)


def valor_metrica_anual(metricas, registro, anio, columna):
    filas = metricas[(metricas["Registro"] == registro) & (metricas["Anio"] == anio)]
    if filas.empty:
        return 0
    return int(filas.iloc[0].get(columna, 0) or 0)


def tarjetas_estado_anual_html(metricas, anios):
    tarjetas = []
    for registro in [TEXT_CASOS, TEXT_INCIDENTES]:
        for anio in anios:
            abiertos = valor_metrica_anual(metricas, registro, anio, TEXT_ABIERTOS)
            cerrados = valor_metrica_anual(metricas, registro, anio, TEXT_CERRADOS)
            total = valor_metrica_anual(metricas, registro, anio, TEXT_TOTAL)
            tarjetas.append(
                '<div class="kpi-card">'
                f'<div class="kpi-title">{html.escape(registro)} {anio}</div>'
                f'<div class="kpi-value">{total}</div>'
                f'<div class="executive-note-line"><strong>Abiertos:</strong> {abiertos}</div>'
                f'<div class="executive-note-line"><strong>Cerrados:</strong> {cerrados}</div>'
                "</div>"
            )
    return '<div class="kpi-grid">' + "".join(tarjetas) + "</div>"


def tarjetas_sla_anual_html(metricas, anios):
    tarjetas = []
    for registro in [TEXT_CASOS, TEXT_INCIDENTES]:
        for anio in anios:
            filas = metricas[(metricas["Registro"] == registro) & (metricas["Anio"] == anio)]
            sla = filas.iloc[0].get("SLA %", 0) if not filas.empty else 0
            tarjetas.append(
                '<div class="kpi-card">'
                f'<div class="kpi-title">SLA {html.escape(registro)} {anio}</div>'
                f'<div class="kpi-value">{sla}%</div>'
                "</div>"
            )
    return '<div class="kpi-grid">' + "".join(tarjetas) + "</div>"


def datos_estado_anual_grafico(metricas, registro, anios):
    filas = []
    for anio in anios:
        filas.append({"Anio": str(anio), TEXT_ESTADO_2: TEXT_ABIERTOS, TEXT_CANTIDAD: valor_metrica_anual(metricas, registro, anio, TEXT_ABIERTOS)})
        filas.append({"Anio": str(anio), TEXT_ESTADO_2: TEXT_CERRADOS, TEXT_CANTIDAD: valor_metrica_anual(metricas, registro, anio, TEXT_CERRADOS)})
    return pd.DataFrame(filas)


def render_grafico_estado_anual(metricas, registro, anios):
    datos = datos_estado_anual_grafico(metricas, registro, anios)
    fig = px.bar(
        datos,
        x="Anio",
        y=TEXT_CANTIDAD,
        color=TEXT_ESTADO_2,
        text=TEXT_CANTIDAD,
        barmode="stack",
        labels={"Anio": "AÃ±o"},
        color_discrete_map={
            TEXT_ABIERTOS: UI_PALETTE[TEXT_PRIMARY],
            TEXT_CERRADOS: UI_PALETTE[TEXT_LAVENDER],
        },
    )
    fig.update_traces(textposition="inside")
    fig.update_layout(showlegend=True, height=390)
    st.plotly_chart(aplicar_estilo_figura(fig, registro), use_container_width=True)


def formato_porcentaje_presentacion(valor):
    if valor is None or pd.isna(valor):
        return "Sin dato"
    try:
        numero = float(valor)
    except (TypeError, ValueError):
        return "Sin dato"
    if numero.is_integer():
        return f"{int(numero)}%"
    return f"{numero:.2f}%"


def filtrar_anio_mes_dashboard(df, anio, mes, columna_dt=TEXT_CREADO_DT_DASHBOARD):
    if df.empty or columna_dt not in df.columns:
        return df.copy()
    fechas = df[columna_dt]
    return df[(fechas.dt.year == int(anio)) & (fechas.dt.month == int(mes))].copy()


def filtrar_incidentes_reales_ans(df):
    if df.empty:
        return df.copy()
    if TEXT_TIPIFICACION_AUTO not in df.columns:
        return df.iloc[0:0].copy()
    tipificacion = df[TEXT_TIPIFICACION_AUTO].fillna("").astype(str)
    return df[tipificacion.isin(ANS_INCIDENT_REAL_TYPES)].copy()


def preparar_incidentes_ans_ejecutivo(incidentes):
    if incidentes.empty:
        return incidentes.copy()
    trabajo = preparar_incidentes_kpi_comparativo_ligero(incidentes)
    return filtrar_incidentes_reales_ans(trabajo)


def columnas_faltantes_ans_incidentes(df):
    missing = set(df.attrs.get("missing_columns", []))
    faltantes = []
    columnas_requeridas = [
        (TEXT_TIPIFICACION_AUTO, "tipificacion del incidente"),
        (TEXT_PRIORIDAD, "prioridad"),
        (TEXT_CREADO, "fecha de creacion"),
    ]
    for columna, etiqueta in columnas_requeridas:
        if columna not in df.columns or columna in missing:
            faltantes.append(etiqueta)
    estado_faltante = TEXT_ESTADO not in df.columns or TEXT_ESTADO in missing
    cierre_faltante = TEXT_CERRADO not in df.columns or TEXT_CERRADO in missing
    if estado_faltante and cierre_faltante:
        faltantes.append("estado o fecha de cierre")
    fuentes_duracion = [TEXT_DURACION_HORAS, "duracion_segundos", TEXT_CERRADO]
    if all(columna not in df.columns or columna in missing for columna in fuentes_duracion):
        faltantes.append("duracion_horas, duracion_segundos o fecha de cierre")
    return faltantes


def resumen_ans_incidentes_periodo(datos, periodo):
    if datos.empty:
        return {
            "Periodo": periodo,
            "Total incidentes": 0,
            "Cerrados/resueltos": 0,
            "Cumplen ANS": 0,
            "No cumplen ANS": 0,
            "Sin dato ANS": 0,
            "ANS valor": None,
            "ANS %": "Sin dato",
        }

    cerrados_mask = datos.get(TEXT_CERRADO_2, mascara_cerrados(datos))
    cerrados = datos[cerrados_mask].copy()
    if TEXT_ESTADO_SLA not in cerrados.columns:
        evaluados = cerrados.iloc[0:0].copy()
    else:
        evaluados = cerrados[cerrados[TEXT_ESTADO_SLA].isin([ESTADO_SLA_CUMPLE, ESTADO_SLA_NO_CUMPLE])].copy()
    cumple = int((evaluados.get(TEXT_ESTADO_SLA, pd.Series(dtype=TEXT_OBJECT)) == ESTADO_SLA_CUMPLE).sum())
    no_cumple = int((evaluados.get(TEXT_ESTADO_SLA, pd.Series(dtype=TEXT_OBJECT)) == ESTADO_SLA_NO_CUMPLE).sum())
    total_evaluado = cumple + no_cumple
    ans_valor = porcentaje(cumple, total_evaluado) if total_evaluado else None
    return {
        "Periodo": periodo,
        "Total incidentes": int(len(datos)),
        "Cerrados/resueltos": int(len(cerrados)),
        "Cumplen ANS": cumple,
        "No cumplen ANS": no_cumple,
        "Sin dato ANS": max(int(len(cerrados) - total_evaluado), 0),
        "ANS valor": ans_valor,
        "ANS %": formato_porcentaje_presentacion(ans_valor),
    }


def resumen_ans_incidentes_desde_filas(tabla, periodo):
    if tabla.empty:
        return resumen_ans_incidentes_periodo(pd.DataFrame(), periodo)
    total = int(tabla["Total incidentes"].sum())
    cerrados = int(tabla["Cerrados/resueltos"].sum())
    cumple = int(tabla["Cumplen ANS"].sum())
    no_cumple = int(tabla["No cumplen ANS"].sum())
    sin_dato = int(tabla["Sin dato ANS"].sum())
    total_evaluado = cumple + no_cumple
    ans_valor = porcentaje(cumple, total_evaluado) if total_evaluado else None
    return {
        "Periodo": periodo,
        "Total incidentes": total,
        "Cerrados/resueltos": cerrados,
        "Cumplen ANS": cumple,
        "No cumplen ANS": no_cumple,
        "Sin dato ANS": sin_dato,
        "ANS valor": ans_valor,
        "ANS %": formato_porcentaje_presentacion(ans_valor),
    }


def tabla_ans_incidentes_2026(base):
    filas = []
    for mes in ANS_INCIDENT_MONTHS_FOCUS:
        etiqueta = f"{MONTH_NAMES_ES.get(mes, mes)} {ANS_INCIDENT_YEAR_FOCUS}"
        datos_mes = filtrar_anio_mes_dashboard(base, ANS_INCIDENT_YEAR_FOCUS, mes)
        filas.append(resumen_ans_incidentes_periodo(datos_mes, etiqueta))
    return pd.DataFrame(filas)


def tarjetas_ans_incidentes_html(resumen_2025, resumen_2026):
    items = [
        ("Incidentes 2025", resumen_2025["Total incidentes"]),
        ("ANS 2025", resumen_2025["ANS %"]),
        ("Incidentes marzo-mayo 2026", resumen_2026["Total incidentes"]),
        ("ANS marzo-mayo 2026", resumen_2026["ANS %"]),
    ]
    tarjetas = [
        '<div class="ans-card">'
        f'<div class="ans-card-label">{html.escape(str(titulo))}</div>'
        f'<div class="ans-card-value">{html.escape(str(valor))}</div>'
        "</div>"
        for titulo, valor in items
    ]
    return '<div class="ans-card-grid">' + "".join(tarjetas) + "</div>"


def tabla_ans_html(titulo, subtitulo, tabla):
    columnas = [
        "Periodo",
        "Total incidentes",
        "Cerrados/resueltos",
        "Cumplen ANS",
        "No cumplen ANS",
        "ANS %",
    ]
    if not tabla.empty and int(tabla.get("Sin dato ANS", pd.Series(dtype=TEXT_FLOAT)).sum()) > 0:
        columnas.append("Sin dato ANS")
    encabezado = "".join(f"<th>{html.escape(columna)}</th>" for columna in columnas)
    filas = []
    for _, row in tabla.iterrows():
        celdas = []
        for columna in columnas:
            valor = row.get(columna, "")
            if columna == "Periodo":
                celdas.append(f'<td class="ans-period">{html.escape(str(valor))}</td>')
            elif columna == "ANS %":
                celdas.append(f'<td><span class="ans-pill">{html.escape(str(valor))}</span></td>')
            else:
                celdas.append(f"<td>{html.escape(str(valor))}</td>")
        filas.append("<tr>" + "".join(celdas) + "</tr>")
    cuerpo = "".join(filas) if filas else f'<tr><td colspan="{len(columnas)}">Sin datos</td></tr>'
    return (
        '<div class="ans-panel">'
        f'<div class="ans-panel-title">{html.escape(titulo)}</div>'
        f'<div class="ans-panel-subtitle">{html.escape(subtitulo)}</div>'
        '<table class="ans-table">'
        f"<thead><tr>{encabezado}</tr></thead>"
        f"<tbody>{cuerpo}</tbody>"
        "</table>"
        "</div>"
    )


def dashboard_kpi_comparativo_anual_legacy():
    st.subheader(MENU_KPI_COMPARATIVO_ANUAL)
    anio_actual = pd.Timestamp.now().year
    anios_disponibles = list(range(2024, max(2026, anio_actual) + 1))
    col_base, col_comparado = st.columns(2)
    with col_base:
        anio_base = st.selectbox(
            "AÃ±o base",
            anios_disponibles,
            index=anios_disponibles.index(2025) if 2025 in anios_disponibles else 0,
            key="kpi_comp_anio_base",
        )
    with col_comparado:
        anio_comparado = st.selectbox(
            "AÃ±o comparado",
            anios_disponibles,
            index=anios_disponibles.index(2026) if 2026 in anios_disponibles else len(anios_disponibles) - 1,
            key="kpi_comp_anio_comparado",
        )

    if anio_base == anio_comparado:
        st.warning("Selecciona dos aÃ±os diferentes para comparar.")
        return

    anios = [anio_base, anio_comparado]
    with st.spinner(f"Cargando solo registros de {anio_base} y {anio_comparado}..."):
        casos = cargar_casos_anios_cache(tuple(anios))
        incidentes = cargar_incidentes_anios_cache(tuple(anios))

    if casos.empty and incidentes.empty:
        st.info("No hay casos ni incidentes cargados para los aÃ±os seleccionados.")
        return

    with st.spinner("Preparando resumen anual..."):
        base_casos, base_incidentes = preparar_bases_kpi_comparativo(casos, incidentes)
    metricas = tabla_metricas_kpi_comparativo(base_casos, base_incidentes, anios)

    st.markdown(tarjetas_estado_anual_html(metricas, anios), unsafe_allow_html=True)
    st.markdown(tarjetas_sla_anual_html(metricas, anios), unsafe_allow_html=True)

    graf_col1, graf_col2 = st.columns(2)
    with graf_col1:
        render_grafico_estado_anual(metricas, TEXT_CASOS, anios)
    with graf_col2:
        render_grafico_estado_anual(metricas, TEXT_INCIDENTES, anios)

    with st.expander("Ver tabla resumen"):
        columnas = ["Anio", "Registro", TEXT_TOTAL, TEXT_ABIERTOS, TEXT_CERRADOS]
        tabla = metricas[[col for col in columnas if col in metricas.columns]].rename(columns={"Anio": "AÃ±o"})
        st.dataframe(tabla, use_container_width=True, hide_index=True)


def dashboard_ans_incidentes_anual():
    st.subheader("ANS de incidentes 2025 vs 2026")
    st.caption(
        "Reporte ejecutivo solo con incidentes reales. Para 2026 se muestran exclusivamente marzo, abril y mayo."
    )

    anios = [ANS_INCIDENT_YEAR_BASE, ANS_INCIDENT_YEAR_FOCUS]
    with st.spinner("Cargando incidentes para el reporte ANS..."):
        incidentes = cargar_incidentes_anios_cache(tuple(anios))

    if incidentes.empty:
        st.info("No hay incidentes cargados para 2025 o 2026.")
        return

    faltantes = columnas_faltantes_ans_incidentes(incidentes)
    if faltantes:
        st.warning(
            "Faltan campos para calcular el ANS completo: "
            + ", ".join(sorted(set(faltantes)))
            + ". Se mostraran los datos disponibles sin inventar valores."
        )

    with st.spinner("Preparando ANS solo de incidentes reales..."):
        base_incidentes = preparar_incidentes_ans_ejecutivo(incidentes)

    if base_incidentes.empty:
        st.info(
            "No hay registros clasificados como Incidente Cliente Externo o Incidente Interno para este reporte."
        )
        return

    base_2025 = filtrar_anio_dashboard(base_incidentes, ANS_INCIDENT_YEAR_BASE)
    resumen_2025 = resumen_ans_incidentes_periodo(base_2025, str(ANS_INCIDENT_YEAR_BASE))
    tabla_2025 = pd.DataFrame([resumen_2025])
    tabla_2026 = tabla_ans_incidentes_2026(base_incidentes)
    resumen_2026 = resumen_ans_incidentes_desde_filas(tabla_2026, "Marzo a mayo 2026")

    st.markdown(tarjetas_ans_incidentes_html(resumen_2025, resumen_2026), unsafe_allow_html=True)
    st.markdown(
        tabla_ans_html(
            "Referencia 2025",
            "Bloque anual conservado solo con incidentes reales.",
            tabla_2025,
        ),
        unsafe_allow_html=True,
    )
    st.markdown(
        tabla_ans_html(
            "Detalle 2026",
            "Solo marzo, abril y mayo de 2026.",
            tabla_2026,
        ),
        unsafe_allow_html=True,
    )

    sin_dato_ans = int(tabla_2025["Sin dato ANS"].sum() + tabla_2026["Sin dato ANS"].sum())
    if sin_dato_ans:
        st.info(
            f"{sin_dato_ans} incidentes cerrados/resueltos no tienen datos suficientes para evaluar ANS "
            "y no se incluyen en el porcentaje."
        )


def dashboard_kpi_comparativo_rango():
    st.subheader("Comparativo KPI por rango de fechas")
    st.caption("Compara casos e incidentes entre dos rangos definidos manualmente.")

    with st.spinner("Cargando casos e incidentes para el comparativo por rango..."):
        casos = cargar_casos_cache()
        incidentes = cargar_incidentes_cache()

    if casos.empty and incidentes.empty:
        st.info("No hay casos ni incidentes cargados para comparar.")
        return

    with st.spinner("Preparando metricas KPI por fecha..."):
        base_casos, base_incidentes = preparar_bases_kpi_comparativo(casos, incidentes)

    fecha_min, fecha_max = fechas_disponibles_comparativo_rango(base_casos, base_incidentes)
    if not fecha_min or not fecha_max:
        st.info("No hay fechas validas para construir el comparativo.")
        return

    rangos = selector_rangos_kpi_comparativo(fecha_min, fecha_max)
    if not rangos:
        return

    etiqueta_base = etiqueta_rango_fechas(rangos[0][1], rangos[0][2])
    etiqueta_comparado = etiqueta_rango_fechas(rangos[1][1], rangos[1][2])
    rangos_metricas = [
        (etiqueta_base, rangos[0][1], rangos[0][2]),
        (etiqueta_comparado, rangos[1][1], rangos[1][2]),
    ]

    metricas = tabla_metricas_kpi_comparativo_rangos(base_casos, base_incidentes, rangos_metricas)
    comparativo = tabla_comparativo_rangos(metricas, etiqueta_base, etiqueta_comparado)
    tendencia = pd.concat(
        [
            tendencia_diaria_kpi_rango(base_casos, TEXT_CASOS, rangos_metricas),
            tendencia_diaria_kpi_rango(base_incidentes, TEXT_INCIDENTES, rangos_metricas),
        ],
        ignore_index=True,
    )

    render_tarjetas_kpi_comparativo_rangos(comparativo)
    st.caption(f"Base: {etiqueta_base} | Comparado: {etiqueta_comparado}")
    render_graficas_kpi_comparativo_rangos(metricas, tendencia)

    st.divider()
    st.subheader("Tabla comparativa")
    tabla_visible = comparativo.copy()
    if not tabla_visible.empty:
        tabla_visible["Variacion total %"] = tabla_visible["Variacion total %"].apply(
            lambda valor: "Sin base" if pd.isna(valor) else formato_porcentaje_presentacion(valor)
        )
    st.dataframe(tabla_visible, use_container_width=True, hide_index=True)

    with st.expander("Ver metricas base por rango"):
        st.dataframe(metricas, use_container_width=True, hide_index=True)


def dashboard_kpi_comparativo_anual():
    tab_rango, tab_anual = st.tabs(["Rango de fechas", "ANS 2025 vs 2026"])
    with tab_rango:
        dashboard_kpi_comparativo_rango()
    with tab_anual:
        dashboard_ans_incidentes_anual()


def opciones_filtro_reincidencias(base, columna):
    if base.empty or columna not in base.columns:
        return []
    return sorted(base[columna].replace("", pd.NA).dropna().astype(str).unique().tolist())


def construir_filtros_reincidencias(base):
    filtros = {}
    st.caption("Filtros aplicados al calculo de reincidencias y problemas sugeridos.")

    fecha_col, tipo_col = st.columns([1.4, 1])
    fechas = base.get("fecha_dt", pd.Series(dtype=PANDAS_DATETIME_DTYPE)).dropna()
    with fecha_col:
        if fechas.empty:
            st.caption("No hay fechas validas para filtrar.")
        else:
            fecha_min = fechas.min().date()
            fecha_max = fechas.max().date()
            rango = st.date_input(
                "Rango de fechas",
                value=(fecha_min, fecha_max),
                min_value=fecha_min,
                max_value=fecha_max,
                key="reincidencias_rango_fechas",
            )
            if isinstance(rango, tuple) and len(rango) == 2:
                filtros["fecha_inicio"], filtros["fecha_fin"] = rango
    with tipo_col:
        tipos = opciones_filtro_reincidencias(base, "tipo_registro")
        seleccion = st.multiselect("Tipo de registro", tipos, default=tipos, key="reincidencias_tipo")
        if seleccion and set(seleccion) != set(tipos):
            filtros["tipo_registro"] = seleccion

    filtro_col1, filtro_col2, filtro_col3, filtro_col4 = st.columns(4)
    filtros_config = [
        (filtro_col1, "Cliente", "cliente_analisis", "reincidencias_cliente"),
        (filtro_col2, "Servicio/producto", "servicio_producto", "reincidencias_servicio"),
        (filtro_col3, TEXT_TIPIFICACION, "tipificacion", "reincidencias_tipificacion"),
        (filtro_col4, "Causa", "causa", "reincidencias_causa"),
    ]
    for columna_ui, etiqueta, columna, key in filtros_config:
        with columna_ui:
            opciones = opciones_filtro_reincidencias(base, columna)
            seleccion = st.multiselect(etiqueta, opciones, key=key)
            if seleccion:
                filtros[columna] = seleccion
    return filtros


def render_tabla_reincidencias(reincidencias):
    if reincidencias.empty:
        st.info("No se identificaron reincidencias con los filtros seleccionados.")
        return
    dataframe_liviano(reincidencias)


def render_tabla_problemas_sugeridos(problemas):
    if problemas.empty:
        st.info("No se generaron problemas sugeridos con los filtros seleccionados.")
        return
    dataframe_liviano(problemas)


def render_detalle_reincidencias(base, problemas=None):
    if base.empty or "nivel_reincidencia" not in base.columns:
        st.info("No hay registros asociados para mostrar.")
        return
    if problemas is not None and not problemas.empty and TEXT_CAUSA in problemas.columns:
        causas_problema = problemas[TEXT_CAUSA].replace("", pd.NA).dropna().unique().tolist()
        detalle = base[base[TEXT_CAUSA].isin(causas_problema)].copy()
    else:
        detalle = base[base["nivel_reincidencia"].fillna("").ne("")].copy()
    if detalle.empty:
        st.info("No hay casos o incidentes asociados con los filtros seleccionados.")
        return
    columnas = [
        "tipo_registro",
        TEXT_NUMERO,
        "cliente_analisis",
        "servicio_producto",
        TEXT_TIPIFICACION_2,
        TEXT_CAUSA,
        "nivel_reincidencia",
        "fecha",
        TEXT_ESTADO,
        TEXT_PRIORIDAD,
        TEXT_DESCRIPCION_2,
        TEXT_ESTADO_SLA,
        "es_alerta",
        "tipo_incidente",
    ]
    columnas = [col for col in columnas if col in detalle.columns]
    dataframe_liviano(detalle.sort_values(by=["fecha_dt"], ascending=False)[columnas])


def preparar_fechas_reincidencias(df):
    if df.empty:
        return df
    if TEXT_CREADO not in df.columns:
        trabajo = df.copy()
        trabajo[TEXT_CREADO_DT_DASHBOARD] = pd.NaT
        return trabajo
    return preparar_fechas_dashboard(df)


def seleccionar_meses_reincidencias(casos, incidentes):
    casos = preparar_fechas_reincidencias(casos)
    incidentes = preparar_fechas_reincidencias(incidentes)

    st.caption("Selecciona el mes antes de calcular. La vista solo analiza ese corte para mantenerla rapida.")
    col_casos, col_incidentes = st.columns(2)
    with col_casos:
        if casos.empty:
            mes_casos = TEXT_TODOS
            st.caption("No hay casos cargados.")
        else:
            mes_casos = selector_mes_dashboard(
                casos,
                "reincidencias_mes_casos",
                label="Mes casos",
                incluir_todos=False,
            )
    with col_incidentes:
        if incidentes.empty:
            mes_incidentes = TEXT_TODOS
            st.caption("No hay incidentes cargados.")
        else:
            mes_incidentes = selector_mes_dashboard(
                incidentes,
                "reincidencias_mes_incidentes",
                label="Mes incidentes",
                incluir_todos=False,
            )

    casos = filtrar_mes_dashboard(casos, mes_casos)
    incidentes = filtrar_mes_dashboard(incidentes, mes_incidentes)
    return casos, incidentes, mes_casos, mes_incidentes


def dashboard_reincidencias_problemas():
    st.subheader(MENU_REINCIDENCIAS_PROBLEMAS)
    st.caption("Selecciona el periodo antes de calcular. La vista solo consulta esos cortes para mantenerla rapida.")
    col_casos, col_incidentes = st.columns(2)
    with col_casos:
        st.markdown("##### Casos")
        anio_casos, mes_casos_num, periodo_casos = selector_periodo_sql("cases", "reincidencias_periodo_casos")
    with col_incidentes:
        st.markdown("##### Incidentes")
        anio_incidentes, mes_incidentes_num, periodo_incidentes = selector_periodo_sql(
            "incidents",
            "reincidencias_periodo_incidentes",
        )

    with st.spinner("Cargando casos e incidentes del periodo seleccionado..."):
        casos = (
            cargar_casos_filtrados_cache(anio_casos, mes_casos_num)
            if anio_casos is not None
            else pd.DataFrame()
        )
        incidentes = (
            cargar_incidentes_filtrados_cache(anio_incidentes, mes_incidentes_num)
            if anio_incidentes is not None
            else pd.DataFrame()
        )
    if casos.empty and incidentes.empty:
        st.info("No hay casos ni incidentes para los periodos seleccionados.")
        return

    with st.spinner("Calculando reincidencias y problemas sugeridos..."):
        base, reincidencias, problemas = analizar_reincidencias_y_problemas(casos, incidentes)
    if base.empty:
        st.info("No hay registros analizables para construir reincidencias.")
        return

    clientes_reincidentes = (
        reincidencias["cliente_analisis"].replace("", pd.NA).dropna().nunique()
        if not reincidencias.empty
        else 0
    )
    registros_asociados = 0
    if not problemas.empty:
        registros_asociados = int(pd.to_numeric(problemas["total_registros"], errors=TEXT_COERCE).sum())
    elif not reincidencias.empty:
        registros_asociados = int(pd.to_numeric(reincidencias["total_registros"], errors=TEXT_COERCE).sum())
    render_tarjetas(
        [
            ("Clientes con reincidencia", clientes_reincidentes),
            ("Problemas sugeridos", len(problemas)),
            ("Registros asociados", registros_asociados),
        ]
    )
    st.caption(
        f"Periodo analizado: casos {periodo_casos} | incidentes {periodo_incidentes}. "
        "Lectura simple: reincidencias de casos/incidentes y posibles problemas cuando una misma causa raiz se repite. "
        "No modifica datos ni crea registros nuevos."
    )

    tab_reincidencias, tab_problemas, tab_detalle = st.tabs(
        ["Reincidencias casos/incidentes", "Problemas por causa", "Detalle asociado"]
    )
    with tab_reincidencias:
        render_tabla_reincidencias(reincidencias)
    with tab_problemas:
        render_tabla_problemas_sugeridos(problemas)
    with tab_detalle:
        render_detalle_reincidencias(base, problemas)


def dashboard_incidentes():
    anio, mes, periodo_label = selector_periodo_sql("incidents", "dashboard_incidentes_periodo")
    if not periodo_sql_valido(anio, "incidentes"):
        return
    df = cargar_incidentes_sla_filtrados_cache(anio, mes)
    if df.empty:
        st.info(f"No hay incidentes cargados para {periodo_label}.")
        return

    df = preparar_fechas_dashboard(df)
    if df.empty:
        st.info(f"No hay incidentes cargados para {periodo_label}.")
        return

    total = len(df)
    cerrados_mask = mascara_cerrados(df)
    cerrados = len(df[cerrados_mask])
    abiertos = total - cerrados

    incidentes_reales = df[df[TEXT_APLICA_SLA_INCIDENTE]].copy()
    atenciones_noc = df[df[TEXT_TIPIFICACION_AUTO].fillna("") == NOC_TIPIFICATION].copy()
    incidentes_externos = df[df[TEXT_TIPIFICACION_AUTO].fillna("") == TIPIFICACION_INCIDENTE_CLIENTE_EXTERNO].copy()
    incidentes_internos = df[df[TEXT_TIPIFICACION_AUTO].fillna("") == TIPIFICACION_INCIDENTE_INTERNO].copy()
    casos_cliente_externo = df[df[TEXT_TIPIFICACION_AUTO].fillna("") == TIPIFICACION_CASO_CLIENTE_EXTERNO].copy()

    duraciones = pd.to_numeric(incidentes_reales[TEXT_DURACION_SLA_HORAS], errors=TEXT_COERCE).dropna()
    promedio = round(duraciones.mean(), 2) if len(duraciones) > 0 else 0

    df_cerrados = incidentes_reales[mascara_cerrados(incidentes_reales)].copy()
    df_cerrados = df_cerrados[df_cerrados[TEXT_ESTADO_SLA].isin([ESTADO_SLA_CUMPLE, ESTADO_SLA_NO_CUMPLE])]
    total_cerrados = len(df_cerrados)
    cumplen = len(df_cerrados[df_cerrados[TEXT_ESTADO_SLA] == ESTADO_SLA_CUMPLE])
    porcentaje_sla = round((cumplen / total_cerrados) * 100, 2) if total_cerrados > 0 else 0
    incumplen = total_cerrados - cumplen
    sla_base = df_cerrados.copy()
    if not sla_base.empty:
        sla_base[TEXT_DURACION_HORAS_NUM] = pd.to_numeric(sla_base[TEXT_DURACION_SLA_HORAS], errors=TEXT_COERCE)
        sla_base[TEXT_DURACION_DIAS_NUM] = (sla_base[TEXT_DURACION_HORAS_NUM] / 24).round(2)
        sla_base["sla_objetivo_dias"] = (pd.to_numeric(sla_base[TEXT_SLA_OBJETIVO_HORAS], errors=TEXT_COERCE) / 24).round(2)
        sla_base[COL_SLA_OBJETIVO] = sla_base[TEXT_SLA_OBJETIVO_HORAS].apply(formato_horas_dias)
        sla_base["Rango resolucion"] = sla_base[TEXT_DURACION_HORAS_NUM].apply(clasificar_rango_resolucion)

    casos_cerrados = casos_cliente_externo[mascara_cerrados(casos_cliente_externo)].copy()
    casos_cerrados = casos_cerrados[casos_cerrados[TEXT_ESTADO_SLA].isin([ESTADO_SLA_CUMPLE, ESTADO_SLA_NO_CUMPLE])]
    casos_cumplen = int((casos_cerrados[TEXT_ESTADO_SLA] == ESTADO_SLA_CUMPLE).sum())
    casos_total_sla = len(casos_cerrados)
    casos_sla = round((casos_cumplen / casos_total_sla) * 100, 2) if casos_total_sla else 0

    render_tarjetas(
        [
            (TEXT_ATENCIONES, total),
            ("Incidentes reales", len(incidentes_reales)),
            ("Alertas y Consultas NOC", len(atenciones_noc)),
            (LABEL_CASOS_CLIENTE_EXTERNO, len(casos_cliente_externo)),
            ("SLA incidentes", f"{porcentaje_sla}%"),
        ]
    )
    st.caption(
        f"{TEXT_PERIODO}{periodo_label} | Cerrados: {cerrados} | Abiertos: {abiertos} | "
        f"Promedio incidentes: {promedio}h | Cumplen: {cumplen}{TEXT_NO_CUMPLEN}{incumplen} | "
        f"Externos: {len(incidentes_externos)} | Internos: {len(incidentes_internos)} | "
        f"SLA casos cliente externo: {casos_sla}%"
    )

    st.divider()
    render_carga_agentes(df, TEXT_ASIGNADO_A, "Carga por agente - incidentes", TEXT_INCIDENTES)

    st.divider()
    st.subheader("Clasificacion de atenciones")
    st.caption(
        "Vista consolidada de lo recibido: NOC, casos cliente externo e incidentes reales internos o externos."
    )
    st.dataframe(
        tabla_resumen_tipologias_incidentes(df),
        use_container_width=True,
        hide_index=True,
    )

    st.divider()
    fila1_col1, fila1_col2 = st.columns(2)

    with fila1_col1:
        tip = df[TEXT_TIPIFICACION_AUTO].fillna("Sin clasificacion").value_counts().reset_index()
        tip.columns = [TEXT_TIPOLOGIA, TEXT_CANTIDAD]
        tip = tip.sort_values(by=TEXT_CANTIDAD, ascending=True)
        fig = px.bar(
            tip,
            x=TEXT_CANTIDAD,
            y=TEXT_TIPOLOGIA,
            orientation="h",
            text=TEXT_CANTIDAD,
            color=TEXT_TIPOLOGIA,
            color_discrete_sequence=CHART_COLORS,
        )
        fig.update_traces(textposition=TEXT_OUTSIDE)
        fig.update_layout(showlegend=False)
        st.plotly_chart(aplicar_estilo_figura(fig, "Clasificacion general"), use_container_width=True)

    with fila1_col2:
        serie = df.copy()
        atenciones_dia = serie.groupby(serie[TEXT_CREADO_DT_DASHBOARD].dt.date).size().reset_index(name=TEXT_ATENCIONES)
        atenciones_dia.columns = [TEXT_FECHA, TEXT_ATENCIONES]
        fig = px.bar(
            atenciones_dia,
            x=TEXT_FECHA,
            y=TEXT_ATENCIONES,
            color_discrete_sequence=[UI_PALETTE[TEXT_PURPLE]],
        )
        fig.update_traces(marker_color=UI_PALETTE[TEXT_PURPLE])
        st.plotly_chart(aplicar_estilo_figura(fig, "Atenciones por dia"), use_container_width=True)

    st.divider()
    st.subheader("SLA de incidentes reales")
    if sla_base.empty:
        st.info("No hay incidentes cerrados con objetivo SLA para el periodo seleccionado.")
        diagnostico_sla = pd.DataFrame(
            [
                ("Incidentes reales", len(incidentes_reales)),
                ("Incidentes reales cerrados", len(incidentes_reales[mascara_cerrados(incidentes_reales)])),
                (
                    "Con objetivo SLA",
                    int(pd.to_numeric(incidentes_reales[TEXT_SLA_OBJETIVO_HORAS], errors=TEXT_COERCE).notna().sum())
                    if TEXT_SLA_OBJETIVO_HORAS in incidentes_reales.columns
                    else 0,
                ),
                (
                    "Con duracion SLA",
                    int(pd.to_numeric(incidentes_reales[TEXT_DURACION_SLA_HORAS], errors=TEXT_COERCE).notna().sum())
                    if TEXT_DURACION_SLA_HORAS in incidentes_reales.columns
                    else 0,
                ),
            ],
            columns=["Validacion", TEXT_CANTIDAD],
        )
        st.dataframe(diagnostico_sla, use_container_width=True, hide_index=True)
    else:
        st.caption(
            "Solo incluye incidentes reales. Seguridad se muestra como tipo de incidente dentro de cliente externo o interno."
        )
        sla_base[TEXT_SEGMENTO] = sla_base[TEXT_TIPIFICACION_AUTO] + " - " + sla_base[TEXT_TIPO_INCIDENTE_AUTO]
        cumplimiento_segmento = (
            sla_base.groupby([TEXT_SEGMENTO, TEXT_ESTADO_SLA])
            .size()
            .reset_index(name=TEXT_CANTIDAD)
            .sort_values(by=TEXT_CANTIDAD, ascending=True)
        )
        fig = px.bar(
            cumplimiento_segmento,
            x=TEXT_CANTIDAD,
            y=TEXT_SEGMENTO,
            orientation="h",
            text=TEXT_CANTIDAD,
            color=TEXT_ESTADO_SLA,
            barmode="stack",
            color_discrete_map={
                ESTADO_SLA_CUMPLE: UI_PALETTE[TEXT_LAVENDER],
                ESTADO_SLA_NO_CUMPLE: UI_PALETTE[TEXT_PRIMARY],
            },
            labels={TEXT_ESTADO_SLA: "Estado SLA"},
        )
        fig.update_traces(textposition="inside")
        st.plotly_chart(aplicar_estilo_figura(fig, "SLA por tipo de incidente"), use_container_width=True)

        sla_resumen = (
            sla_base.groupby(
                [TEXT_TIPIFICACION_AUTO, TEXT_TIPO_INCIDENTE_AUTO, TEXT_PRIORIDAD_NORMALIZADA, TEXT_SLA_OBJETIVO_HORAS, COL_SLA_OBJETIVO],
                dropna=False,
            )
            .agg(
                Total=(TEXT_NUMERO, TEXT_COUNT),
                Cumple=(TEXT_ESTADO_SLA, lambda serie: int((serie == ESTADO_SLA_CUMPLE).sum())),
                No_cumple=(TEXT_ESTADO_SLA, lambda serie: int((serie == ESTADO_SLA_NO_CUMPLE).sum())),
                Prom_horas=(TEXT_DURACION_HORAS_NUM, "mean"),
                Prom_dias=(TEXT_DURACION_DIAS_NUM, "mean"),
                Max_horas=(TEXT_DURACION_HORAS_NUM, "max"),
                Max_dias=(TEXT_DURACION_DIAS_NUM, "max"),
            )
            .reset_index()
        )
        sla_resumen["Cumplimiento %"] = sla_resumen.apply(
            lambda row: porcentaje(row[ESTADO_SLA_CUMPLE], row[TEXT_TOTAL]),
            axis=1,
        )
        sla_resumen = sla_resumen.rename(
            columns={
                TEXT_TIPIFICACION_AUTO: TEXT_TIPIFICACION,
                TEXT_TIPO_INCIDENTE_AUTO: "Tipo incidente",
                TEXT_PRIORIDAD_NORMALIZADA: TEXT_PRIORIDAD_2,
                TEXT_SLA_OBJETIVO_HORAS: COL_SLA_OBJETIVO_H,
                "No_cumple": ESTADO_SLA_NO_CUMPLE,
                "Prom_horas": COL_PROM_HORAS,
                "Prom_dias": COL_PROM_DIAS,
                "Max_horas": COL_MAX_HORAS,
                "Max_dias": COL_MAX_DIAS,
            }
        )
        columnas_numericas = [COL_PROM_HORAS, COL_PROM_DIAS, COL_MAX_HORAS, COL_MAX_DIAS, COL_SLA_OBJETIVO_H]
        for columna in columnas_numericas:
            sla_resumen[columna] = pd.to_numeric(sla_resumen[columna], errors=TEXT_COERCE).round(2)
        st.dataframe(
            sla_resumen[
                [
                    TEXT_TIPIFICACION,
                    "Tipo incidente",
                    TEXT_PRIORIDAD_2,
                    COL_SLA_OBJETIVO,
                    COL_SLA_OBJETIVO_H,
                    COL_PROM_HORAS,
                    COL_PROM_DIAS,
                    COL_MAX_HORAS,
                    COL_MAX_DIAS,
                    ESTADO_SLA_CUMPLE,
                    ESTADO_SLA_NO_CUMPLE,
                    TEXT_TOTAL,
                    "Cumplimiento %",
                ]
            ],
            use_container_width=True,
            hide_index=True,
        )

    st.divider()
    st.subheader("Causas raiz de incidentes")
    causa_col1, causa_col2 = st.columns(2)
    with causa_col1:
        render_causas_incidentes(incidentes_externos, "Causas cliente externo", "% cliente externo")
    with causa_col2:
        render_causas_incidentes(incidentes_internos, "Causas cliente interno", "% cliente interno")

    st.divider()
    st.subheader(LABEL_CASOS_CLIENTE_EXTERNO)
    if casos_cliente_externo.empty:
        st.info("No hay casos cliente externo para el periodo seleccionado.")
        return

    caso_col1, caso_col2 = st.columns(2)
    with caso_col1:
        st.markdown(tarjeta(LABEL_CASOS_CLIENTE_EXTERNO, len(casos_cliente_externo)), unsafe_allow_html=True)
        st.caption("SLA propio de mesa de ayuda: 36 horas habiles.")
    with caso_col2:
        st.markdown(tarjeta("SLA casos", f"{casos_sla}%"), unsafe_allow_html=True)
        st.caption(f"Cumplen: {casos_cumplen}{TEXT_NO_CUMPLEN}{casos_total_sla - casos_cumplen}")

    motivos = resumen_motivos_caso_cliente_externo(casos_cliente_externo)
    if motivos.empty:
        st.info("No hay motivos suficientes para graficar casos cliente externo.")
    else:
        grafico_motivos = motivos.sort_values(by=TEXT_CANTIDAD, ascending=True)
        fig = px.bar(
            grafico_motivos,
            x=TEXT_CANTIDAD,
            y=COL_MOTIVO_CASO,
            orientation="h",
            text=TEXT_CANTIDAD,
            color_discrete_sequence=[UI_PALETTE[TEXT_YELLOW]],
        )
        fig.update_traces(textposition=TEXT_OUTSIDE)
        st.plotly_chart(aplicar_estilo_figura(fig, "Motivos de casos cliente externo"), use_container_width=True)
        st.dataframe(motivos, use_container_width=True, hide_index=True)


def texto_normalizado_campos(row, campos):
    return " ".join(normalizar_texto(row.get(campo)) for campo in campos).strip()


def contiene_no_recibio_acuse(texto):
    if not texto:
        return False
    return any(re.search(patron, texto) for patron in PATRONES_NO_RECIBIO_ACUSE)


def motivo_caso_seguimiento_rpost(texto):
    if contiene_no_recibio_acuse(texto):
        return "No recibio acuse"
    for motivo, palabras in CASE_RPOST_RELATION_RULES:
        if any(palabra in texto for palabra in palabras):
            return motivo
    return ""


def preparar_fechas_seguimiento_rpost(df):
    trabajo = df.copy()
    if trabajo.empty or TEXT_CREADO not in trabajo.columns:
        trabajo[TEXT_CREADO_DT_DASHBOARD] = pd.Series(dtype=PANDAS_DATETIME_DTYPE)
        return trabajo
    return preparar_fechas_dashboard(trabajo)


def cliente_caso_seguimiento_rpost(row):
    cuenta = valor_limpio(row.get(TEXT_CUENTA))
    texto_cliente = " ".join(
        [cuenta] + [valor_limpio(row.get(campo)) for campo in CASE_FIELDS_SEGUIMIENTO_RPOST]
    )
    cliente_detectado = detectar_cliente_clave(texto_cliente)
    if cliente_detectado:
        return cliente_detectado
    if cuenta:
        return cuenta
    return SIN_DATO


def cliente_incidente_seguimiento_rpost(row):
    texto_cliente = " ".join(valor_limpio(row.get(campo)) for campo in INCIDENT_FIELDS_SEGUIMIENTO_RPOST)
    cliente_detectado = detectar_cliente_clave(texto_cliente)
    if cliente_detectado:
        return cliente_detectado
    empresa = valor_limpio(row.get(TEXT_EMPRESA))
    if empresa:
        return empresa
    solicitante = valor_limpio(row.get(TEXT_SOLICITANTE))
    return solicitante or SIN_DATO


def filtrar_casos_no_recibio_acuse(casos):
    trabajo = preparar_fechas_seguimiento_rpost(normalizar_tipificaciones_casos_df(casos))
    if trabajo.empty:
        trabajo[TEXT_CLIENTE] = pd.Series(dtype=TEXT_OBJECT)
        return trabajo

    trabajo["_texto_seguimiento_rpost"] = trabajo.apply(
        lambda row: texto_normalizado_campos(row, CASE_FIELDS_SEGUIMIENTO_RPOST),
        axis=1,
    )
    trabajo["Relacion RPost"] = trabajo["_texto_seguimiento_rpost"].apply(motivo_caso_seguimiento_rpost)
    trabajo = trabajo[trabajo["Relacion RPost"] != ""].copy()
    if trabajo.empty:
        trabajo[TEXT_CLIENTE] = pd.Series(dtype=TEXT_OBJECT)
        return trabajo
    trabajo[TEXT_CLIENTE] = trabajo.apply(cliente_caso_seguimiento_rpost, axis=1)
    return trabajo


def filtrar_incidentes_rpost(incidentes):
    trabajo = preparar_fechas_seguimiento_rpost(incidentes)
    if trabajo.empty:
        trabajo[TEXT_CLIENTE] = pd.Series(dtype=TEXT_OBJECT)
        return trabajo

    trabajo["_texto_seguimiento_rpost"] = trabajo.apply(
        lambda row: texto_normalizado_campos(row, INCIDENT_FIELDS_SEGUIMIENTO_RPOST),
        axis=1,
    )
    trabajo = trabajo[
        trabajo["_texto_seguimiento_rpost"].str.contains(
            r"(?<![a-z0-9])rpost(?![a-z0-9])",
            regex=True,
            na=False,
        )
    ].copy()
    if trabajo.empty:
        trabajo[TEXT_CLIENTE] = pd.Series(dtype=TEXT_OBJECT)
        return trabajo
    trabajo[TEXT_CLIENTE] = trabajo.apply(cliente_incidente_seguimiento_rpost, axis=1)
    return trabajo


@st.cache_data(ttl=CACHE_TTL_SEGUNDOS, show_spinner=False)
def cargar_casos_rpost_cache():
    return filtrar_casos_no_recibio_acuse(cargar_casos_cache())


@st.cache_data(ttl=CACHE_TTL_SEGUNDOS, show_spinner=False)
def cargar_casos_rpost_filtrados_cache(anio=None, mes=None):
    return filtrar_casos_no_recibio_acuse(cargar_casos_filtrados_cache(anio, mes))


@st.cache_data(ttl=CACHE_TTL_SEGUNDOS, show_spinner=False)
def cargar_incidentes_rpost_cache():
    return filtrar_incidentes_rpost(cargar_incidentes_cache())


@st.cache_data(ttl=CACHE_TTL_SEGUNDOS, show_spinner=False)
def cargar_incidentes_rpost_filtrados_cache(anio=None, mes=None):
    return filtrar_incidentes_rpost(cargar_incidentes_filtrados_cache(anio, mes))


def fechas_seguimiento_rpost(casos, incidentes):
    fechas = []
    for df in [casos, incidentes]:
        if not df.empty and TEXT_CREADO_DT_DASHBOARD in df.columns:
            fechas.append(df[TEXT_CREADO_DT_DASHBOARD])
    return pd.concat(fechas).dropna() if fechas else pd.Series(dtype=PANDAS_DATETIME_DTYPE)


def clientes_seguimiento_rpost(casos, incidentes):
    clientes = []
    for df in [casos, incidentes]:
        if not df.empty and TEXT_CLIENTE in df.columns:
            clientes.extend(df[TEXT_CLIENTE].replace("", pd.NA).dropna().astype(str).tolist())
    return sorted(set(clientes))


def filtrar_busqueda_seguimiento_rpost(df, busqueda, campos):
    if df.empty or not busqueda:
        return df
    texto_busqueda = normalizar_texto(busqueda)
    campos_busqueda = [TEXT_NUMERO, TEXT_CLIENTE] + campos
    mascara = df.apply(
        lambda row: texto_busqueda in texto_normalizado_campos(row, campos_busqueda),
        axis=1,
    )
    return df[mascara].copy()


def aplicar_filtros_seguimiento_rpost(casos, incidentes, mes_dashboard, cliente, busqueda):
    casos = filtrar_mes_dashboard(casos, mes_dashboard)
    incidentes = filtrar_mes_dashboard(incidentes, mes_dashboard)

    if cliente != TEXT_TODOS:
        casos = casos[casos[TEXT_CLIENTE] == cliente].copy()
        incidentes = incidentes[incidentes[TEXT_CLIENTE] == cliente].copy()
    if busqueda:
        casos = filtrar_busqueda_seguimiento_rpost(casos, busqueda, CASE_FIELDS_SEGUIMIENTO_RPOST)
        incidentes = filtrar_busqueda_seguimiento_rpost(
            incidentes,
            busqueda,
            INCIDENT_FIELDS_SEGUIMIENTO_RPOST,
        )
    return casos, incidentes


def base_eventos_seguimiento_rpost(casos, incidentes):
    tablas = []
    if not casos.empty:
        casos_eventos = casos[[TEXT_CREADO_DT_DASHBOARD, TEXT_CLIENTE, TEXT_NUMERO]].copy()
        casos_eventos["Tipo"] = "Caso RPost / acuses"
        tablas.append(casos_eventos)
    if not incidentes.empty:
        incidentes_eventos = incidentes[[TEXT_CREADO_DT_DASHBOARD, TEXT_CLIENTE, TEXT_NUMERO]].copy()
        incidentes_eventos["Tipo"] = "Incidente RPost"
        tablas.append(incidentes_eventos)
    if not tablas:
        return pd.DataFrame(columns=[TEXT_CREADO_DT_DASHBOARD, TEXT_CLIENTE, TEXT_NUMERO, "Tipo"])
    return pd.concat(tablas, ignore_index=True)


def lista_numeros_resumida(serie, limite=18):
    numeros = []
    for valor in serie.tolist():
        numero = valor_limpio(valor)
        if numero and numero not in numeros:
            numeros.append(numero)
    if len(numeros) > limite:
        return ", ".join(numeros[:limite]) + f" +{len(numeros) - limite} mas"
    return ", ".join(numeros)


def resumen_clientes_seguimiento_rpost(casos, incidentes):
    filas = []
    clientes = clientes_seguimiento_rpost(casos, incidentes)
    for cliente in clientes:
        casos_cliente = casos[casos[TEXT_CLIENTE] == cliente] if not casos.empty else pd.DataFrame()
        incidentes_cliente = (
            incidentes[incidentes[TEXT_CLIENTE] == cliente] if not incidentes.empty else pd.DataFrame()
        )
        total_casos = len(casos_cliente)
        total_incidentes = len(incidentes_cliente)
        filas.append(
            {
                TEXT_CLIENTE: cliente,
                "Casos RPost/acuses": total_casos,
                "Incidentes RPost": total_incidentes,
                TEXT_TOTAL: total_casos + total_incidentes,
                "Numeros casos": lista_numeros_resumida(casos_cliente.get(TEXT_NUMERO, pd.Series(dtype=TEXT_OBJECT))),
                "Numeros incidentes": lista_numeros_resumida(
                    incidentes_cliente.get(TEXT_NUMERO, pd.Series(dtype=TEXT_OBJECT))
                ),
            }
        )
    if not filas:
        return pd.DataFrame()
    return pd.DataFrame(filas).sort_values(by=[TEXT_TOTAL, TEXT_CLIENTE], ascending=[False, True])


def render_graficas_seguimiento_rpost(casos, incidentes):
    eventos = base_eventos_seguimiento_rpost(casos, incidentes)
    if eventos.empty:
        st.info("No hay informacion para graficar con los filtros seleccionados.")
        return

    col_clientes, col_fechas = st.columns(2)
    with col_clientes:
        resumen_cliente = (
            eventos.groupby([TEXT_CLIENTE, "Tipo"])
            .size()
            .reset_index(name=TEXT_CANTIDAD)
            .sort_values(by=TEXT_CANTIDAD, ascending=True)
        )
        fig = px.bar(
            resumen_cliente,
            x=TEXT_CANTIDAD,
            y=TEXT_CLIENTE,
            color="Tipo",
            orientation="h",
            text=TEXT_CANTIDAD,
            barmode="group",
            color_discrete_sequence=[UI_PALETTE[TEXT_PRIMARY], UI_PALETTE[TEXT_LAVENDER]],
        )
        fig.update_traces(textposition=TEXT_OUTSIDE)
        st.plotly_chart(aplicar_estilo_figura(fig, "Volumen por cliente"), use_container_width=True)

    with col_fechas:
        eventos_fecha = eventos.dropna(subset=[TEXT_CREADO_DT_DASHBOARD]).copy()
        if eventos_fecha.empty:
            st.info("No hay fechas validas para graficar la tendencia.")
        else:
            eventos_fecha[TEXT_FECHA] = eventos_fecha[TEXT_CREADO_DT_DASHBOARD].dt.date
            resumen_fecha = eventos_fecha.groupby([TEXT_FECHA, "Tipo"]).size().reset_index(name=TEXT_CANTIDAD)
            fig = px.bar(
                resumen_fecha,
                x=TEXT_FECHA,
                y=TEXT_CANTIDAD,
                color="Tipo",
                barmode="group",
                color_discrete_sequence=[UI_PALETTE[TEXT_PRIMARY], UI_PALETTE[TEXT_LAVENDER]],
            )
            st.plotly_chart(aplicar_estilo_figura(fig, "Actividad por fecha"), use_container_width=True)


def render_detalle_casos_seguimiento_rpost(casos):
    st.subheader("Casos relacionados con RPost y acuses")
    if casos.empty:
        st.info("No hay casos RPost/acuses para los filtros seleccionados.")
        return
    columnas = [
        TEXT_NUMERO,
        TEXT_CLIENTE,
        "Relacion RPost",
        TEXT_ESTADO,
        TEXT_PRIORIDAD,
        TEXT_CREADO,
        TEXT_CERRADO,
        TEXT_CUENTA,
        TEXT_PRODUCTO,
        TEXT_TIPIFICACION_2,
        TEXT_DESCRIPCION_2,
        "notas_resolucion",
        TEXT_OBSERVACIONES_TRABAJO,
    ]
    visibles = [col for col in columnas if col in casos.columns]
    tabla = casos.sort_values(by=TEXT_CREADO_DT_DASHBOARD, ascending=False)[visibles].rename(
        columns={
            TEXT_NUMERO: "Numero caso",
            TEXT_CUENTA: "Cuenta",
            TEXT_PRODUCTO: "Producto",
            TEXT_TIPIFICACION_2: TEXT_TIPIFICACION,
        }
    )
    dataframe_liviano(tabla)


def render_detalle_incidentes_seguimiento_rpost(incidentes):
    st.subheader("Incidentes RPost")
    if incidentes.empty:
        st.info("No hay incidentes RPost para los filtros seleccionados.")
        return
    columnas = [
        TEXT_NUMERO,
        TEXT_CLIENTE,
        TEXT_ESTADO,
        TEXT_PRIORIDAD,
        TEXT_CREADO,
        TEXT_CERRADO,
        TEXT_EMPRESA,
        TEXT_SOLICITANTE,
        TEXT_SERVICIO_NEGOCIO,
        TEXT_TIPIFICACION_AUTO,
        TEXT_TIPO_INCIDENTE_AUTO,
        TEXT_CAUSA_RAIZ_AUTO,
        TEXT_BREVE_DESCRIPCION,
        TEXT_DESCRIPCION_2,
    ]
    visibles = [col for col in columnas if col in incidentes.columns]
    tabla = incidentes.sort_values(by=TEXT_CREADO_DT_DASHBOARD, ascending=False)[visibles].rename(
        columns={
            TEXT_NUMERO: "Numero incidente",
            TEXT_SERVICIO_NEGOCIO: "Servicio",
            TEXT_TIPIFICACION_AUTO: TEXT_TIPIFICACION,
            TEXT_TIPO_INCIDENTE_AUTO: "Tipo incidente",
            TEXT_CAUSA_RAIZ_AUTO: COL_CAUSA_RAIZ,
            TEXT_BREVE_DESCRIPCION: "Breve descripcion",
        }
    )
    dataframe_liviano(tabla)


def render_resumen_disponibilidad_rpost(resumen):
    disponibilidad = float(resumen.get("disponibilidad", 100.0) or 100.0)
    caidas = int(resumen.get("caidas", 0) or 0)
    cumple_sla = bool(resumen.get("cumple_sla", True))
    estado = "Cumple SLA" if cumple_sla else "No cumple SLA"
    estado_clase = "ok" if cumple_sla else "bad"

    st.markdown(
        f"""
        <style>
        .rpost-sla-grid {{
            display: grid;
            grid-template-columns: minmax(0, 760px);
            justify-content: center;
            gap: 1rem;
            margin: 0.3rem 0 0.6rem;
        }}
        .rpost-sla-panel {{
            min-height: 178px;
            border: 1px solid var(--border);
            border-radius: 8px;
            background: var(--surface);
            box-shadow: 0 10px 24px rgba(20, 20, 20, 0.06);
            padding: 22px 24px;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
        }}
        .rpost-sla-eyebrow {{
            color: var(--muted);
            font-size: 0.82rem;
            font-weight: 900;
            line-height: 1.2;
            text-transform: uppercase;
            letter-spacing: 0;
        }}
        .rpost-sla-value {{
            color: var(--primary);
            font-size: 3.25rem;
            font-weight: 900;
            line-height: 1;
            font-variant-numeric: tabular-nums;
            overflow-wrap: anywhere;
        }}
        .rpost-sla-footer {{
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 0.75rem;
            border-top: 1px solid rgba(20, 20, 20, 0.1);
            padding-top: 14px;
        }}
        .rpost-sla-status {{
            border-radius: 999px;
            padding: 6px 10px;
            font-size: 0.82rem;
            font-weight: 900;
            white-space: nowrap;
        }}
        .rpost-sla-status.ok {{
            background: rgba(22, 163, 74, 0.1);
            color: #166534;
        }}
        .rpost-sla-status.bad {{
            background: rgba(220, 38, 38, 0.1);
            color: #991b1b;
        }}
        .rpost-sla-caption {{
            color: var(--muted);
            font-size: 0.9rem;
            font-weight: 800;
            text-align: right;
        }}
        .rpost-sla-caption strong {{
            color: var(--text);
            font-size: 1.35rem;
            font-weight: 900;
            font-variant-numeric: tabular-nums;
        }}
        @media (max-width: 760px) {{
            .rpost-sla-grid {{
                grid-template-columns: 1fr;
            }}
            .rpost-sla-value {{
                font-size: 2.65rem;
            }}
        }}
        </style>
        <div class="rpost-sla-grid">
            <section class="rpost-sla-panel">
                <div class="rpost-sla-eyebrow">SLA disponibilidad</div>
                <div class="rpost-sla-value">{disponibilidad:.2f}%</div>
                <div class="rpost-sla-footer">
                    <span class="rpost-sla-status {estado_clase}">{html.escape(estado)}</span>
                    <span class="rpost-sla-caption">Incidentes con caida<br><strong>{caidas}</strong></span>
                </div>
            </section>
        </div>
        """,
        unsafe_allow_html=True,
    )


def dashboard_seguimiento_rpost():
    st.subheader(MENU_SEGUIMIENTO_RPOST)
    anio, mes, periodo_label = selector_periodo_multi_sql(
        ["cases", "incidents"],
        "seguimiento_rpost_periodo",
    )
    if not periodo_sql_valido(anio, "seguimiento RPost"):
        return
    with st.spinner("Cargando seguimiento de RPost..."):
        casos = cargar_casos_rpost_filtrados_cache(anio, mes)
        incidentes = cargar_incidentes_rpost_filtrados_cache(anio, mes)

    if casos.empty and incidentes.empty:
        st.info(f"No hay casos RPost/acuses ni incidentes RPost para {periodo_label}.")
        return

    filtros_col1, filtros_col2 = st.columns([1.4, 1.4])
    clientes = clientes_seguimiento_rpost(casos, incidentes)
    with filtros_col1:
        cliente = st.selectbox("Cliente", [TEXT_TODOS] + clientes, key="seguimiento_rpost_cliente")
    with filtros_col2:
        busqueda = st.text_input("Buscar numero o texto", key="seguimiento_rpost_busqueda")

    casos, incidentes = aplicar_filtros_seguimiento_rpost(
        casos,
        incidentes,
        TEXT_TODOS,
        cliente,
        busqueda,
    )

    if casos.empty and incidentes.empty:
        st.info("No hay registros para los filtros seleccionados.")
        return

    total_clientes = len(clientes_seguimiento_rpost(casos, incidentes))
    render_tarjetas(
        [
            ("Casos RPost/acuses", len(casos)),
            ("Clientes en casos", casos[TEXT_CLIENTE].nunique() if not casos.empty else 0),
            ("Incidentes RPost", len(incidentes)),
            ("Clientes en incidentes", incidentes[TEXT_CLIENTE].nunique() if not incidentes.empty else 0),
            ("Clientes total", total_clientes),
        ]
    )
    st.caption(
        f"{TEXT_PERIODO}{periodo_label} | Casos: no recibio acuse o Certimail/Certicmal | "
        "Incidentes: cualquier registro que mencione RPost."
    )

    st.divider()
    st.subheader("Resumen por cliente")
    resumen = resumen_clientes_seguimiento_rpost(casos, incidentes)
    if resumen.empty:
        st.info("No hay clientes para resumir.")
    else:
        st.dataframe(resumen, use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("Disponibilidad RPost")
    
    if not incidentes.empty:
        mes_ts = pd.Timestamp(year=anio, month=mes, day=1)
        render_resumen_disponibilidad_rpost(resumir_disponibilidad_mes(incidentes, mes_ts))
        
    else:
        st.info("No hay incidentes para calcular disponibilidad.")

    st.divider()
    render_graficas_seguimiento_rpost(casos, incidentes)

    st.divider()
    render_detalle_casos_seguimiento_rpost(casos)

    st.divider()
    render_detalle_incidentes_seguimiento_rpost(incidentes)


def fechas_clientes_clave(casos, incidentes):
    fechas = []
    if not casos.empty:
        fechas.append(casos[TEXT_CREADO_DT])
    if not incidentes.empty:
        fechas.append(incidentes[TEXT_CREADO_DT])
    return pd.concat(fechas).dropna() if fechas else pd.Series(dtype=PANDAS_DATETIME_DTYPE)


def seleccionar_filtros_clientes_clave(fechas, key_prefix="clientes_clave"):
    filtro_col1, filtro_col2, filtro_col3 = st.columns([2, 1, 1])
    with filtro_col1:
        clientes_seleccionados = st.multiselect(
            "Clientes",
            CLIENTES_CLAVE,
            default=CLIENTES_CLAVE,
            key=f"{key_prefix}_filtro",
        )
    with filtro_col2:
        base_meses = pd.DataFrame({TEXT_CREADO_DT: fechas})
        mes_dashboard = selector_mes_dashboard(base_meses, f"{key_prefix}_mes", TEXT_CREADO_DT)
    with filtro_col3:
        rango_fechas = selector_rango_fechas_clientes(fechas, key_prefix)
    return clientes_seleccionados, mes_dashboard, rango_fechas


def selector_rango_fechas_clientes(fechas, key_prefix="clientes_clave"):
    if fechas.empty:
        return None
    fecha_min = fechas.min().date()
    fecha_max = fechas.max().date()
    return st.date_input(
        "Rango de fechas",
        value=(fecha_min, fecha_max),
        min_value=fecha_min,
        max_value=fecha_max,
        key=f"{key_prefix}_rango",
    )


def filtrar_clientes_seleccionados(casos, incidentes, clientes_seleccionados):
    casos = casos[casos[TEXT_CLIENTE_CLAVE].isin(clientes_seleccionados)].copy()
    incidentes = incidentes[incidentes[TEXT_CLIENTE_CLAVE].isin(clientes_seleccionados)].copy()
    return casos, incidentes


def aplicar_filtro_mes_clientes(casos, incidentes, mes_dashboard):
    if mes_dashboard == TEXT_TODOS:
        return casos, incidentes
    if not casos.empty:
        casos = filtrar_mes_dashboard(casos, mes_dashboard, TEXT_CREADO_DT)
    if not incidentes.empty:
        incidentes = filtrar_mes_dashboard(incidentes, mes_dashboard, TEXT_CREADO_DT)
    return casos, incidentes


def aplicar_filtro_rango_clientes(casos, incidentes, rango_fechas):
    if not (rango_fechas and isinstance(rango_fechas, tuple) and len(rango_fechas) == 2):
        return casos, incidentes
    fecha_inicio = pd.Timestamp(rango_fechas[0])
    fecha_fin = pd.Timestamp(rango_fechas[1]) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
    if not casos.empty:
        casos = casos[casos[TEXT_CREADO_DT].between(fecha_inicio, fecha_fin, inclusive="both")].copy()
    if not incidentes.empty:
        incidentes = incidentes[incidentes[TEXT_CREADO_DT].between(fecha_inicio, fecha_fin, inclusive="both")].copy()
    return casos, incidentes


def aplicar_filtros_clientes_clave(casos, incidentes, clientes_seleccionados, mes_dashboard, rango_fechas):
    casos, incidentes = filtrar_clientes_seleccionados(casos, incidentes, clientes_seleccionados)
    casos, incidentes = aplicar_filtro_mes_clientes(casos, incidentes, mes_dashboard)
    return aplicar_filtro_rango_clientes(casos, incidentes, rango_fechas)


def calcular_sla_casos_clientes(casos):
    casos_cerrados = casos[mascara_cerrados(casos)] if not casos.empty else pd.DataFrame()
    tiempos_casos = casos_cerrados.get(TEXT_TIEMPO_RESPUESTA_H, pd.Series(dtype=TEXT_FLOAT)).dropna()
    if not len(tiempos_casos):
        return 0
    return porcentaje(len(tiempos_casos[tiempos_casos < SLA_CASOS_HORAS]), len(tiempos_casos))


def preparar_casos_clientes_clave_comparativo(casos):
    trabajo = preparar_casos_clientes_clave(casos)
    if trabajo.empty:
        return trabajo
    trabajo[TEXT_CERRADO_2] = mascara_cerrados(trabajo)
    trabajo[TEXT_ABIERTO] = ~trabajo[TEXT_CERRADO_2]
    trabajo["_tiempo_eval_sla_h"] = pd.to_numeric(trabajo.get(TEXT_TIEMPO_RESPUESTA_H), errors=TEXT_COERCE)
    trabajo["Cumple SLA"] = trabajo["_tiempo_eval_sla_h"].apply(
        lambda valor: "Si" if pd.notna(valor) and valor < SLA_CASOS_HORAS else "No"
    )
    return trabajo


def resumen_casos_clientes_clave_periodo(casos, periodo):
    columnas = [
        "Periodo",
        TEXT_CLIENTE,
        TEXT_CASOS,
        TEXT_CERRADOS,
        TEXT_ABIERTOS,
        COL_CUMPLE_SLA,
        COL_NO_CUMPLE_SLA,
        "SLA %",
    ]
    if casos.empty:
        return pd.DataFrame(columns=columnas)

    filas = []
    for cliente in CLIENTES_CLAVE:
        datos = casos[casos[TEXT_CLIENTE_CLAVE] == cliente].copy()
        if datos.empty:
            continue
        cerrados = datos[datos[TEXT_CERRADO_2]].copy()
        tiempos = pd.to_numeric(cerrados.get("_tiempo_eval_sla_h", pd.Series(dtype=TEXT_FLOAT)), errors=TEXT_COERCE).dropna()
        cumple = int((tiempos < SLA_CASOS_HORAS).sum())
        no_cumple = int(len(tiempos) - cumple)
        filas.append(
            {
                "Periodo": periodo,
                TEXT_CLIENTE: cliente,
                TEXT_CASOS: int(len(datos)),
                TEXT_CERRADOS: int(len(cerrados)),
                TEXT_ABIERTOS: int(len(datos) - len(cerrados)),
                COL_CUMPLE_SLA: cumple,
                COL_NO_CUMPLE_SLA: no_cumple,
                "SLA %": formato_porcentaje_presentacion(porcentaje(cumple, len(tiempos)) if len(tiempos) else None),
            }
        )
    if not filas:
        return pd.DataFrame(columns=columnas)
    return pd.DataFrame(filas, columns=columnas).sort_values(by=[TEXT_CASOS, TEXT_CLIENTE], ascending=[False, True])


def resumen_total_casos_clientes_clave(tabla, periodo):
    if tabla.empty:
        return {
            "Periodo": periodo,
            TEXT_CASOS: 0,
            TEXT_CERRADOS: 0,
            TEXT_ABIERTOS: 0,
            COL_CUMPLE_SLA: 0,
            COL_NO_CUMPLE_SLA: 0,
            "SLA valor": None,
            "SLA %": "Sin dato",
        }
    casos = int(tabla[TEXT_CASOS].sum())
    cerrados = int(tabla[TEXT_CERRADOS].sum())
    abiertos = int(tabla[TEXT_ABIERTOS].sum())
    cumple = int(tabla[COL_CUMPLE_SLA].sum())
    no_cumple = int(tabla[COL_NO_CUMPLE_SLA].sum())
    evaluados = cumple + no_cumple
    sla_valor = porcentaje(cumple, evaluados) if evaluados else None
    return {
        "Periodo": periodo,
        TEXT_CASOS: casos,
        TEXT_CERRADOS: cerrados,
        TEXT_ABIERTOS: abiertos,
        COL_CUMPLE_SLA: cumple,
        COL_NO_CUMPLE_SLA: no_cumple,
        "SLA valor": sla_valor,
        "SLA %": formato_porcentaje_presentacion(sla_valor),
    }


def tabla_casos_clientes_clave_2026(base):
    tablas = []
    for mes in KEY_CLIENT_CASE_MONTHS_FOCUS:
        etiqueta = f"{MONTH_NAMES_ES.get(mes, mes)} {KEY_CLIENT_CASE_YEAR_FOCUS}"
        datos_mes = filtrar_anio_mes_dashboard(base, KEY_CLIENT_CASE_YEAR_FOCUS, mes, TEXT_CREADO_DT)
        tabla_mes = resumen_casos_clientes_clave_periodo(datos_mes, etiqueta)
        if not tabla_mes.empty:
            tablas.append(tabla_mes)
    if not tablas:
        return pd.DataFrame()
    return pd.concat(tablas, ignore_index=True)


def tarjetas_casos_clientes_clave_html(resumen_2025, resumen_2026):
    items = [
        ("Casos 2025", resumen_2025[TEXT_CASOS]),
        ("Abiertos 2025", resumen_2025[TEXT_ABIERTOS]),
        ("SLA 2025", resumen_2025["SLA %"]),
        ("Casos marzo-abril-junio 2026", resumen_2026[TEXT_CASOS]),
        ("Abiertos 2026", resumen_2026[TEXT_ABIERTOS]),
        ("SLA 2026", resumen_2026["SLA %"]),
    ]
    tarjetas = [
        '<div class="ans-card">'
        f'<div class="ans-card-label">{html.escape(str(titulo))}</div>'
        f'<div class="ans-card-value">{html.escape(str(valor))}</div>'
        "</div>"
        for titulo, valor in items
    ]
    return '<div class="ans-card-grid">' + "".join(tarjetas) + "</div>"


def tabla_casos_clientes_html(titulo, subtitulo, tabla):
    columnas = [
        "Periodo",
        TEXT_CLIENTE,
        TEXT_CASOS,
        TEXT_CERRADOS,
        TEXT_ABIERTOS,
        COL_CUMPLE_SLA,
        COL_NO_CUMPLE_SLA,
        "SLA %",
    ]
    encabezado = "".join(f"<th>{html.escape(columna)}</th>" for columna in columnas)
    filas = []
    for _, row in tabla.iterrows():
        celdas = []
        for columna in columnas:
            valor = row.get(columna, "")
            if columna in ["Periodo", TEXT_CLIENTE]:
                celdas.append(f'<td class="ans-period">{html.escape(str(valor))}</td>')
            elif columna == "SLA %":
                celdas.append(f'<td><span class="ans-pill">{html.escape(str(valor))}</span></td>')
            else:
                celdas.append(f"<td>{html.escape(str(valor))}</td>")
        filas.append("<tr>" + "".join(celdas) + "</tr>")
    cuerpo = "".join(filas) if filas else f'<tr><td colspan="{len(columnas)}">Sin datos</td></tr>'
    return (
        '<div class="ans-panel">'
        f'<div class="ans-panel-title">{html.escape(titulo)}</div>'
        f'<div class="ans-panel-subtitle">{html.escape(subtitulo)}</div>'
        '<table class="ans-table">'
        f"<thead><tr>{encabezado}</tr></thead>"
        f"<tbody>{cuerpo}</tbody>"
        "</table>"
        "</div>"
    )


def dashboard_casos_clientes_clave_comparativo():
    st.subheader("Casos clientes clave 2025 vs 2026")
    st.caption(
        "Reporte ejecutivo solo con casos de clientes clave. Para 2026 se muestran marzo, abril y junio."
    )

    with st.spinner("Cargando casos de clientes clave para el comparativo..."):
        casos = cargar_casos_anios_cache((KEY_CLIENT_CASE_YEAR_BASE, KEY_CLIENT_CASE_YEAR_FOCUS))

    if casos.empty:
        st.info("No hay casos cargados para 2025 o 2026.")
        return

    faltantes = set(casos.attrs.get("missing_columns", []))
    requeridas = {TEXT_CUENTA, TEXT_CREADO, TEXT_ESTADO, TEXT_TIEMPO_RESPUESTA}
    faltantes_visibles = sorted(requeridas.intersection(faltantes))
    if faltantes_visibles:
        st.warning(
            "Faltan campos para calcular el reporte completo: "
            + ", ".join(faltantes_visibles)
            + ". Se mostraran los datos disponibles sin inventar valores."
        )

    base = preparar_casos_clientes_clave_comparativo(casos)
    if base.empty:
        st.info("No hay casos asociados a la lista de clientes clave para los periodos indicados.")
        return

    tabla_2025 = resumen_casos_clientes_clave_periodo(
        filtrar_anio_dashboard(base, KEY_CLIENT_CASE_YEAR_BASE, TEXT_CREADO_DT),
        str(KEY_CLIENT_CASE_YEAR_BASE),
    )
    tabla_2026 = tabla_casos_clientes_clave_2026(base)
    resumen_2025 = resumen_total_casos_clientes_clave(tabla_2025, str(KEY_CLIENT_CASE_YEAR_BASE))
    resumen_2026 = resumen_total_casos_clientes_clave(tabla_2026, "Marzo, abril y junio 2026")

    st.markdown(tarjetas_casos_clientes_clave_html(resumen_2025, resumen_2026), unsafe_allow_html=True)
    st.markdown(
        tabla_casos_clientes_html(
            "Referencia 2025",
            "Casos anuales por cliente clave.",
            tabla_2025,
        ),
        unsafe_allow_html=True,
    )
    st.markdown(
        tabla_casos_clientes_html(
            "Detalle 2026",
            "Solo marzo, abril y junio de 2026.",
            tabla_2026,
        ),
        unsafe_allow_html=True,
    )

    with st.expander("Lista de clientes clave"):
        st.dataframe(pd.DataFrame({TEXT_CLIENTE: CLIENTES_CLAVE}), use_container_width=True, hide_index=True)


def calcular_sla_incidentes_clientes(incidentes):
    incidentes_cerrados = incidentes[mascara_cerrados(incidentes)] if not incidentes.empty else pd.DataFrame()
    if not incidentes_cerrados.empty and TEXT_APLICA_SLA_INCIDENTE in incidentes_cerrados.columns:
        incidentes_cerrados = incidentes_cerrados[incidentes_cerrados[TEXT_APLICA_SLA_INCIDENTE].fillna(False)]
    if incidentes_cerrados.empty or TEXT_ESTADO_SLA not in incidentes_cerrados.columns:
        return 0
    sla_base = incidentes_cerrados[
        incidentes_cerrados[TEXT_ESTADO_SLA].isin([ESTADO_SLA_CUMPLE, ESTADO_SLA_NO_CUMPLE])
    ]
    if not len(sla_base):
        return 0
    return porcentaje(len(sla_base[sla_base[TEXT_ESTADO_SLA] == ESTADO_SLA_CUMPLE]), len(sla_base))


def metricas_dashboard_clientes(casos, incidentes, resumen_actividad):
    abiertos_casos = len(casos[~mascara_cerrados(casos)]) if not casos.empty else 0
    abiertos_incidentes = len(incidentes[~mascara_cerrados(incidentes)]) if not incidentes.empty else 0
    clientes_seguimiento = len(
        resumen_actividad[
            resumen_actividad[TEXT_NIVEL].isin([TEXT_AMARILLO, "Rojo"])
            & (resumen_actividad[TEXT_ABIERTOS] > 0)
        ]
    )
    return {
        TEXT_TOTAL_CASOS: len(casos),
        TEXT_TOTAL_INCIDENTES: len(incidentes),
        "abiertos": abiertos_casos + abiertos_incidentes,
        "clientes_activos": len(resumen_actividad),
        "clientes_seguimiento": clientes_seguimiento,
        "sla_casos": calcular_sla_casos_clientes(casos),
        "sla_incidentes": calcular_sla_incidentes_clientes(incidentes),
    }


def render_kpis_clientes_clave(metricas, mes_dashboard):
    render_tarjetas(
        [
            ("Clientes activos", metricas["clientes_activos"]),
            (TEXT_ATENCIONES, metricas[TEXT_TOTAL_CASOS] + metricas[TEXT_TOTAL_INCIDENTES]),
            (TEXT_ABIERTOS, metricas["abiertos"]),
            (f"SLA casos <{SLA_CASOS_HORAS}h", f"{metricas['sla_casos']}%"),
            ("SLA incidentes", f"{metricas['sla_incidentes']}%"),
        ]
    )
    st.caption(
        f"{TEXT_PERIODO}{mes_dashboard} | Casos: {metricas[TEXT_TOTAL_CASOS]} | "
        f"Incidentes: {metricas[TEXT_TOTAL_INCIDENTES]} | "
        f"Clientes en seguimiento: {metricas['clientes_seguimiento']}"
    )


def render_clientes_sin_actividad(resumen):
    clientes_sin_actividad = resumen[resumen[COL_TOTAL_ATENCIONES] == 0][TEXT_CLIENTE].tolist()
    if clientes_sin_actividad:
        st.caption("Sin actividad en el periodo: " + ", ".join(clientes_sin_actividad))


def render_grafico_atenciones_cliente(resumen_actividad, color_estado=True):
    grafico = resumen_actividad.sort_values(by=COL_TOTAL_ATENCIONES, ascending=True)
    if not color_estado:
        render_ranking_kpi(grafico, TEXT_CLIENTE, COL_TOTAL_ATENCIONES, "Atenciones por cliente clave")
        return

    parametros_color = (
        {
            "color": TEXT_NIVEL,
            "color_discrete_map": {
                "Verde": UI_PALETTE[TEXT_LAVENDER],
                TEXT_AMARILLO: UI_PALETTE[TEXT_YELLOW],
                "Rojo": UI_PALETTE[TEXT_PRIMARY],
            },
        }
        if color_estado
        else {"color_discrete_sequence": [UI_PALETTE["mustard"]]}
    )
    fig = px.bar(
        grafico,
        x=COL_TOTAL_ATENCIONES,
        y=TEXT_CLIENTE,
        orientation="h",
        text=COL_TOTAL_ATENCIONES,
        **parametros_color,
    )
    fig.update_traces(textposition=TEXT_OUTSIDE)
    fig = aplicar_estilo_figura(fig, "Atenciones por cliente clave")
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def actividad_diaria_clientes(casos, incidentes):
    actividad = []
    if not casos.empty:
        casos_dia = casos[[TEXT_CREADO_DT]].dropna().copy()
        casos_dia[TEXT_FECHA] = casos_dia[TEXT_CREADO_DT].dt.date
        casos_dia["Tipo"] = TEXT_CASOS
        actividad.append(casos_dia.groupby([TEXT_FECHA, "Tipo"]).size().reset_index(name=TEXT_CANTIDAD))
    if not incidentes.empty:
        incidentes_dia = incidentes[[TEXT_CREADO_DT]].dropna().copy()
        incidentes_dia[TEXT_FECHA] = incidentes_dia[TEXT_CREADO_DT].dt.date
        incidentes_dia["Tipo"] = TEXT_INCIDENTES
        actividad.append(incidentes_dia.groupby([TEXT_FECHA, "Tipo"]).size().reset_index(name=TEXT_CANTIDAD))
    return pd.concat(actividad, ignore_index=True) if actividad else pd.DataFrame()


def render_grafico_actividad_clientes(casos, incidentes):
    actividad_dia = actividad_diaria_clientes(casos, incidentes)
    if actividad_dia.empty:
        st.info("No hay fechas validas para graficar actividad.")
        return
    fig = px.bar(
        actividad_dia,
        x=TEXT_FECHA,
        y=TEXT_CANTIDAD,
        color="Tipo",
        barmode="group",
        color_discrete_sequence=[UI_PALETTE[TEXT_PRIMARY], UI_PALETTE[TEXT_PURPLE]],
    )
    st.plotly_chart(aplicar_estilo_figura(fig, "Actividad por dia"), use_container_width=True)


def render_grafico_productos_clientes(casos):
    if casos.empty:
        st.info("No hay casos asociados a clientes clave.")
        return
    productos = casos[TEXT_PRODUCTO].replace("", pd.NA).fillna("Sin producto").value_counts().reset_index()
    productos.columns = ["Producto", TEXT_CANTIDAD]
    productos = productos.head(10).sort_values(by=TEXT_CANTIDAD, ascending=True)
    fig = px.bar(
        productos,
        x=TEXT_CANTIDAD,
        y="Producto",
        orientation="h",
        text=TEXT_CANTIDAD,
        color_discrete_sequence=[UI_PALETTE[TEXT_YELLOW]],
    )
    fig.update_traces(textposition=TEXT_OUTSIDE)
    st.plotly_chart(aplicar_estilo_figura(fig, "Productos con mas casos"), use_container_width=True)


def render_grafico_causas_clientes(incidentes):
    if incidentes.empty:
        st.info("No hay incidentes asociados a clientes clave.")
        return
    causas = (
        incidentes[TEXT_CAUSA_RAIZ_AUTO]
        .replace("", pd.NA)
        .fillna("Sin inferencia")
        .value_counts()
        .reset_index()
    )
    causas.columns = ["Causa incidente", TEXT_CANTIDAD]
    causas = causas.head(10).sort_values(by=TEXT_CANTIDAD, ascending=True)
    fig = px.bar(
        causas,
        x=TEXT_CANTIDAD,
        y="Causa incidente",
        orientation="h",
        text=TEXT_CANTIDAD,
        color_discrete_sequence=[UI_PALETTE[TEXT_PURPLE]],
    )
    fig.update_traces(textposition=TEXT_OUTSIDE)
    st.plotly_chart(aplicar_estilo_figura(fig, "Causas en incidentes"), use_container_width=True)


def render_graficas_clientes_clave(casos, incidentes, resumen_actividad):
    st.divider()
    graf_col1, graf_col2 = st.columns(2)
    with graf_col1:
        render_grafico_atenciones_cliente(resumen_actividad)
    with graf_col2:
        render_grafico_actividad_clientes(casos, incidentes)

    graf_col3, graf_col4 = st.columns(2)
    with graf_col3:
        render_grafico_productos_clientes(casos)
    with graf_col4:
        render_grafico_causas_clientes(incidentes)


def columnas_resumen_clientes():
    return [
        TEXT_CLIENTE,
        TEXT_NIVEL,
        COL_ESTADO_ATENCION,
        TEXT_SCORE,
        TEXT_CASOS,
        TEXT_INCIDENTES,
        COL_TOTAL_ATENCIONES,
        TEXT_ABIERTOS,
        COL_SLA_CASOS_PCT,
        COL_SLA_INCIDENTES_PCT,
        COL_ALERTAS_INCIDENTES,
        COL_CASOS_SIN_CAUSA,
        COL_PRODUCTO_PRINCIPAL,
        COL_CAUSA_INCIDENTE_PRINCIPAL,
        COL_ULTIMA_ATENCION,
    ]


def render_tab_resumen_clientes(resumen):
    st.dataframe(
        resumen[columnas_resumen_clientes()].sort_values(by=[COL_TOTAL_ATENCIONES, TEXT_SCORE], ascending=[False, True]),
        use_container_width=True,
        hide_index=True,
    )


def render_tab_abiertos_clientes(casos, incidentes):
    abiertos_detalle = tabla_atenciones_abiertas_clientes(casos, incidentes)
    if abiertos_detalle.empty:
        st.success("No hay casos ni incidentes abiertos para los clientes seleccionados.")
    else:
        st.caption("Detalle de atenciones abiertas. El campo Tipo indica si corresponde a caso o incidente.")
        dataframe_liviano(abiertos_detalle)


def render_tab_casos_clientes(casos):
    if casos.empty:
        st.info("No hay casos asociados a los clientes seleccionados.")
        return
    columnas_casos = [
        TEXT_CLIENTE_CLAVE,
        TEXT_NUMERO,
        TEXT_CUENTA,
        TEXT_ESTADO,
        TEXT_PRIORIDAD,
        TEXT_PRODUCTO,
        TEXT_TIPIFICACION_2,
        TEXT_TIEMPO_RESPUESTA,
        TEXT_CREADO,
        TEXT_CERRADO,
        TEXT_CAUSA,
        TEXT_CODIGO_RESOLUCION,
        TEXT_FUENTE_CLIENTE,
    ]
    columnas_casos = [col for col in columnas_casos if col in casos.columns]
    dataframe_liviano(casos.sort_values(by=TEXT_CREADO_DT, ascending=False)[columnas_casos])


def render_tab_incidentes_clientes(incidentes):
    if incidentes.empty:
        st.info("No hay incidentes asociados a los clientes seleccionados.")
        return
    columnas_incidentes = [
        TEXT_CLIENTE_CLAVE,
        TEXT_NUMERO,
        TEXT_EMPRESA,
        TEXT_SOLICITANTE,
        TEXT_ESTADO,
        TEXT_PRIORIDAD,
        TEXT_SERVICIO_NEGOCIO,
        TEXT_TIPIFICACION_AUTO,
        TEXT_ES_ALERTA_AUTO,
        TEXT_CAUSA_RAIZ_AUTO,
        TEXT_DURACION_HORAS,
        TEXT_CREADO,
        TEXT_CERRADO,
        TEXT_BREVE_DESCRIPCION,
        TEXT_FUENTE_CLIENTE,
    ]
    columnas_incidentes = [col for col in columnas_incidentes if col in incidentes.columns]
    dataframe_liviano(incidentes.sort_values(by=TEXT_CREADO_DT, ascending=False)[columnas_incidentes])


def clientes_en_seguimiento(resumen_actividad):
    return resumen_actividad[
        resumen_actividad[TEXT_NIVEL].isin([TEXT_AMARILLO, "Rojo"])
        & (resumen_actividad[TEXT_ABIERTOS] > 0)
    ].copy()


def render_tab_seguimiento_clientes(resumen_actividad):
    seguimiento = clientes_en_seguimiento(resumen_actividad)
    if seguimiento.empty:
        st.success("Los clientes clave con actividad estan en nivel estable para el periodo seleccionado.")
        return
    seguimiento = seguimiento.sort_values(by=[TEXT_NIVEL, TEXT_SCORE, TEXT_ABIERTOS], ascending=[False, True, False])
    st.dataframe(
        seguimiento[
            [
                TEXT_CLIENTE,
                TEXT_NIVEL,
                COL_ESTADO_ATENCION,
                TEXT_SCORE,
                TEXT_ABIERTOS,
                COL_ALERTAS_INCIDENTES,
                PRIORIDAD_ALTA,
                COL_CASOS_SIN_CAUSA,
                COL_PRODUCTO_PRINCIPAL,
                COL_CAUSA_INCIDENTE_PRINCIPAL,
            ]
        ],
        use_container_width=True,
        hide_index=True,
    )
    for _, row in seguimiento.iterrows():
        st.warning(
            f"**{row[TEXT_CLIENTE]}** | {row[COL_ESTADO_ATENCION]} | "
            f"Abiertos: {row[TEXT_ABIERTOS]} | Alertas: {row[COL_ALERTAS_INCIDENTES]} | "
            f"Casos sin causa: {row[COL_CASOS_SIN_CAUSA]}"
        )


def render_tabs_clientes_clave(resumen, casos, incidentes, resumen_actividad):
    st.divider()
    tab_resumen, tab_abiertos, tab_casos, tab_incidentes, tab_seguimiento = st.tabs(
        [TEXT_RESUMEN, TEXT_ABIERTOS, TEXT_CASOS, TEXT_INCIDENTES, "Seguimiento"]
    )
    with tab_resumen:
        render_tab_resumen_clientes(resumen)
    with tab_abiertos:
        render_tab_abiertos_clientes(casos, incidentes)
    with tab_casos:
        render_tab_casos_clientes(casos)
    with tab_incidentes:
        render_tab_incidentes_clientes(incidentes)
    with tab_seguimiento:
        render_tab_seguimiento_clientes(resumen_actividad)


def render_tarjetas_kpi_clientes_clave(metricas):
    render_tarjetas(
        [
            ("Clientes activos", metricas["clientes_activos"]),
            (TEXT_ATENCIONES, metricas[TEXT_TOTAL_CASOS] + metricas[TEXT_TOTAL_INCIDENTES]),
            (TEXT_ABIERTOS, metricas["abiertos"]),
            ("En seguimiento", metricas["clientes_seguimiento"]),
            (f"SLA casos <{SLA_CASOS_HORAS}h", f"{metricas['sla_casos']}%"),
            ("SLA incidentes", f"{metricas['sla_incidentes']}%"),
        ]
    )


def fila_cliente_prioritario(resumen_actividad):
    if resumen_actividad.empty:
        return None
    seguimiento = clientes_en_seguimiento(resumen_actividad)
    base = seguimiento if not seguimiento.empty else resumen_actividad
    return base.sort_values(
        by=[TEXT_SCORE, TEXT_ABIERTOS, COL_TOTAL_ATENCIONES],
        ascending=[True, False, False],
    ).iloc[0]


def lineas_lectura_kpi_clientes_clave(metricas, resumen_actividad):
    mayor_actividad = resumen_actividad.sort_values(by=COL_TOTAL_ATENCIONES, ascending=False).iloc[0]
    prioritario = fila_cliente_prioritario(resumen_actividad)
    score_promedio = round(resumen_actividad[TEXT_SCORE].mean(), 2)

    cliente_top = html.escape(str(mayor_actividad[TEXT_CLIENTE]))
    cliente_prioritario = html.escape(str(prioritario[TEXT_CLIENTE]))
    return [
        (
            "Cliente con mas actividad: "
            f"<strong>{cliente_top}</strong> ({int(mayor_actividad[COL_TOTAL_ATENCIONES])} atenciones)."
        ),
        (
            "Cliente a priorizar: "
            f"<strong>{cliente_prioritario}</strong> "
            f"(score {prioritario[TEXT_SCORE]}, abiertos {int(prioritario[TEXT_ABIERTOS])})."
        ),
        (
            f'<div class="slide-note-muted">Score promedio: {score_promedio} | '
            f"Clientes en seguimiento: {metricas['clientes_seguimiento']}.</div>"
        ),
    ]


def render_lectura_kpi_clientes_clave(metricas, resumen_actividad):
    if resumen_actividad.empty:
        st.info("No hay clientes clave con actividad para generar lectura KPI.")
        return

    lineas = lineas_lectura_kpi_clientes_clave(metricas, resumen_actividad)
    contenido = f"""
    <div class="executive-note">
        <div class="executive-note-title">Lectura</div>
        <div class="executive-note-line">{lineas[0]}</div>
        <div class="executive-note-line">{lineas[1]}</div>
        <div class="executive-note-conclusion">{lineas[2]}</div>
    </div>
    """
    st.markdown(contenido, unsafe_allow_html=True)


def render_detalle_kpi_clientes_clave(resumen):
    columnas = [
        TEXT_CLIENTE,
        TEXT_NIVEL,
        TEXT_SCORE,
        COL_TOTAL_ATENCIONES,
        TEXT_ABIERTOS,
        COL_SLA_CASOS_PCT,
        COL_SLA_INCIDENTES_PCT,
    ]
    with st.expander("Ver resumen por cliente"):
        st.dataframe(
            resumen[columnas].sort_values(by=[COL_TOTAL_ATENCIONES, TEXT_SCORE], ascending=[False, True]),
            use_container_width=True,
            hide_index=True,
        )


def render_slide_kpi_clientes_clave(metricas, resumen_actividad, mes_dashboard, clientes_seleccionados):
    tarjetas = [
        ("Clientes activos", metricas["clientes_activos"]),
        (TEXT_ATENCIONES, metricas[TEXT_TOTAL_CASOS] + metricas[TEXT_TOTAL_INCIDENTES]),
        (TEXT_ABIERTOS, metricas["abiertos"]),
        ("En seguimiento", metricas["clientes_seguimiento"]),
        (f"SLA casos <{SLA_CASOS_HORAS}h", f"{metricas['sla_casos']}%"),
        ("SLA incidentes", f"{metricas['sla_incidentes']}%"),
    ]
    caption = (
        f"Clientes seleccionados: {len(clientes_seleccionados)} | "
        f"Casos: {metricas[TEXT_TOTAL_CASOS]} | Incidentes: {metricas[TEXT_TOTAL_INCIDENTES]}"
    )
    lineas = lineas_lectura_kpi_clientes_clave(metricas, resumen_actividad)
    izquierda = slide_ranking_html(
        resumen_actividad,
        TEXT_CLIENTE,
        COL_TOTAL_ATENCIONES,
        "Atenciones por cliente clave",
        top_n=6,
        limite=58,
    )
    derecha = slide_note_html("Lectura", lineas)
    render_slide_frame_kpi(MENU_KPI_CLIENTES_CLAVE, mes_dashboard, tarjetas, caption, izquierda, derecha)



def dashboard_kpi_clientes_clave():
    st.subheader(MENU_KPI_CLIENTES_CLAVE)
    vista = st.radio(
        "Vista",
        ["Dashboard actual", "Comparativo casos clientes clave"],
        horizontal=True,
        key="kpi_clientes_clave_vista",
    )
    if vista == "Comparativo casos clientes clave":
        dashboard_casos_clientes_clave_comparativo()
        return

    anio, mes, periodo_label = selector_periodo_multi_sql(
        ["cases", "incidents"],
        "kpi_clientes_clave_periodo",
    )
    if not periodo_sql_valido(anio, "clientes clave"):
        return
    clientes_seleccionados = st.multiselect(
        "Clientes",
        CLIENTES_CLAVE,
        default=CLIENTES_CLAVE,
        key="kpi_clientes_clave_filtro",
    )
    if not clientes_seleccionados:
        st.warning("Selecciona al menos un cliente clave.")
        return

    casos = cargar_casos_clientes_clave_filtrados_cache(anio, mes)
    incidentes = cargar_incidentes_clientes_clave_filtrados_cache(anio, mes)
    casos, incidentes = filtrar_clientes_seleccionados(casos, incidentes, clientes_seleccionados)
    resumen = resumen_clientes_clave(casos, incidentes)
    resumen = resumen[resumen[TEXT_CLIENTE].isin(clientes_seleccionados)].copy()
    resumen_actividad = resumen[resumen[COL_TOTAL_ATENCIONES] > 0].copy()
    metricas = metricas_dashboard_clientes(casos, incidentes, resumen_actividad)

    modo_diapositiva = st.toggle("Formato diapositiva 16:9", key="slide_kpi_clientes_clave")
    if modo_diapositiva and not resumen_actividad.empty:
        render_slide_kpi_clientes_clave(metricas, resumen_actividad, periodo_label, clientes_seleccionados)
        return

    render_tarjetas_kpi_clientes_clave(metricas)
    st.caption(
        f"{TEXT_PERIODO}{periodo_label} | Clientes seleccionados: {len(clientes_seleccionados)} | "
        f"Casos: {metricas[TEXT_TOTAL_CASOS]} | Incidentes: {metricas[TEXT_TOTAL_INCIDENTES]}"
    )
    render_clientes_sin_actividad(resumen)

    if resumen_actividad.empty:
        st.info("No hay casos o incidentes asociados a los clientes seleccionados en el periodo.")
        return

    col_grafico, col_lectura = st.columns([2.15, 1])
    with col_grafico:
        render_grafico_atenciones_cliente(resumen_actividad, color_estado=False)
    with col_lectura:
        render_lectura_kpi_clientes_clave(metricas, resumen_actividad)

    render_detalle_kpi_clientes_clave(resumen)


def dashboard_clientes_clave():
    st.subheader(MENU_CLIENTES_CLAVE)

    anio, mes, periodo_label = selector_periodo_multi_sql(
        ["cases", "incidents"],
        "clientes_clave_periodo",
    )
    if not periodo_sql_valido(anio, "clientes clave"):
        return
    clientes_seleccionados = st.multiselect(
        "Clientes",
        CLIENTES_CLAVE,
        default=CLIENTES_CLAVE,
        key="clientes_clave_filtro",
    )
    if not clientes_seleccionados:
        st.warning("Selecciona al menos un cliente clave.")
        return

    casos = cargar_casos_clientes_clave_filtrados_cache(anio, mes)
    incidentes = cargar_incidentes_clientes_clave_filtrados_cache(anio, mes)
    casos, incidentes = filtrar_clientes_seleccionados(casos, incidentes, clientes_seleccionados)
    resumen = resumen_clientes_clave(casos, incidentes)
    resumen = resumen[resumen[TEXT_CLIENTE].isin(clientes_seleccionados)].copy()
    resumen_actividad = resumen[resumen[COL_TOTAL_ATENCIONES] > 0].copy()

    metricas = metricas_dashboard_clientes(casos, incidentes, resumen_actividad)
    render_kpis_clientes_clave(metricas, periodo_label)
    render_clientes_sin_actividad(resumen)

    if resumen_actividad.empty:
        st.info("No hay casos o incidentes asociados a los clientes seleccionados en el periodo.")
        return

    render_graficas_clientes_clave(casos, incidentes, resumen_actividad)
    render_tabs_clientes_clave(resumen, casos, incidentes, resumen_actividad)

def mensaje_carga_casos(cargados, reemplazados, duplicados_archivo, reemplazar_meses, meses_reemplazados, eliminados):
    detalle_reemplazo = ""
    if reemplazar_meses:
        meses = ", ".join(meses_reemplazados) if meses_reemplazados else "sin fechas validas"
        detalle_reemplazo = f" | Meses reemplazados: {meses} | Registros eliminados antes de cargar: {eliminados}"
    return (
        f"Cargados: {cargados} | Registros existentes actualizados: {reemplazados} | "
        f"Duplicados/filas sin numero omitidos del archivo: {duplicados_archivo} | "
        "Los meses no incluidos en el archivo se conservan."
        f"{detalle_reemplazo}"
    )


def crear_barra_progreso_carga(mensaje_inicial):
    texto = st.empty()
    barra = st.progress(0)

    def actualizar(valor, mensaje=None):
        porcentaje_barra = int(round(max(0, min(1, float(valor))) * 100))
        texto.caption(mensaje or mensaje_inicial)
        barra.progress(porcentaje_barra)

    actualizar(0, mensaje_inicial)
    return actualizar


def procesar_archivo_casos(df, reemplazar_meses):
    actualizar_progreso = crear_barra_progreso_carga("Procesando casos...")
    try:
        cargados, reemplazados, eliminados, meses_reemplazados, duplicados_archivo = guardar_casos(
            df,
            reemplazar_meses=reemplazar_meses,
            progress_callback=actualizar_progreso,
        )
    except Exception as error:
        actualizar_progreso(1, "No se pudo finalizar la carga de casos.")
        if es_error_db_transitorio(error):
            st.error(
                "La base de datos estaba ocupada por otra carga. Espera unos segundos y vuelve a procesar el archivo."
            )
            return
        raise
    if cargados == 0:
        st.error("No se guardaron casos. Revisa que el archivo tenga una columna de numero de caso.")
        return
    limpiar_cache_datos()
    actualizar_progreso(1, "Carga de casos finalizada.")
    st.success(
        mensaje_carga_casos(
            cargados,
            reemplazados,
            duplicados_archivo,
            reemplazar_meses,
            meses_reemplazados,
            eliminados,
        )
    )


def vista_cargar_casos():
    archivo = st.file_uploader("Sube Excel de casos", type=["xlsx"], key="casos_upload")
    if not archivo:
        return

    df = pd.read_excel(archivo)
    st.write(f"Filas detectadas: {len(df)}")
    st.dataframe(df.head(), use_container_width=True, hide_index=True)
    reemplazar_meses = st.checkbox(
        "Reemplazar los meses incluidos en este archivo",
        value=True,
        help=(
            "Si esta activo, antes de cargar se eliminan de la base los casos de los meses "
            "presentes en el Excel y luego se carga el archivo. Asi el mes queda igual al corte subido."
        ),
    )
    if st.button("Procesar casos"):
        procesar_archivo_casos(df, reemplazar_meses)

def vista_casos():
    anio, mes, periodo_label = selector_periodo_sql("cases", "vista_casos_periodo")
    if not periodo_sql_valido(anio, "casos"):
        return
    df = cargar_casos_soporte_filtrados_cache(anio, mes)
    filtro_estado = TEXT_TODOS
    filtro_soporte = TEXT_TODOS
    filtro_servicio = TEXT_TODOS
    filtro_cuenta = ""
    if not df.empty:
        df = preparar_fechas_dashboard(df)
        df["mes"] = df[TEXT_CREADO_DT_DASHBOARD].dt.to_period("M").astype(str).replace("NaT", "Sin fecha")

        filtro_col1, filtro_col2, filtro_col3, filtro_col4 = st.columns([1, 1.5, 1.5, 2])
        with filtro_col1:
            estados = sorted(df[TEXT_ESTADO].dropna().unique().tolist())
            filtro_estado = st.selectbox(TEXT_ESTADO_2, [TEXT_TODOS] + estados, key="estado_casos")
        with filtro_col2:
            filtro_soporte = st.selectbox(
                TEXT_TIPOLOGIA_SOPORTE,
                [TEXT_TODOS] + CASE_SUPPORT_TYPOLOGY_ORDER,
                key="tipologia_soporte_casos",
            )
        with filtro_col3:
            servicios = opciones_filtro_servicio(df, TEXT_PRODUCTO)
            filtro_servicio = st.selectbox("Servicio", [TEXT_TODOS] + servicios, key="servicio_casos")
        with filtro_col4:
            filtro_cuenta = st.text_input("Cuenta", key="cuenta_casos")

        filtro_estado_sql = filtro_estado if filtro_estado != TEXT_TODOS else ""
        filtro_servicio_sql = (
            filtro_servicio
            if filtro_servicio not in (TEXT_TODOS, SIN_SERVICIO)
            else ""
        )
        if filtro_estado_sql or filtro_servicio_sql or filtro_cuenta:
            df = cargar_casos_soporte_filtrados_cache(
                anio,
                mes,
                filtro_cuenta,
                filtro_estado_sql,
                filtro_servicio_sql,
            )
            df = preparar_fechas_dashboard(df)
            df["mes"] = df[TEXT_CREADO_DT_DASHBOARD].dt.to_period("M").astype(str).replace("NaT", "Sin fecha")

        if filtro_estado != TEXT_TODOS:
            df = df[df[TEXT_ESTADO] == filtro_estado]
        if filtro_soporte != TEXT_TODOS:
            df = df[df[TEXT_TIPOLOGIA_SOPORTE] == filtro_soporte]
        if filtro_servicio != TEXT_TODOS:
            df = filtrar_por_servicio(df, TEXT_PRODUCTO, filtro_servicio)
        if filtro_cuenta:
            df = df[df[TEXT_CUENTA].fillna("").str.contains(filtro_cuenta, case=False, na=False)]
        df = df.drop(columns=[TEXT_CREADO_DT_DASHBOARD], errors="ignore")
        columnas = [
            TEXT_NUMERO,
            TEXT_ESTADO,
            "mes",
            TEXT_TIPOLOGIA_SOPORTE,
            TEXT_CUENTA,
            "contacto",
            TEXT_DESCRIPCION_2,
            TEXT_PRIORIDAD,
            TEXT_ASIGNADO,
            TEXT_CREADO,
            TEXT_CERRADO,
            TEXT_PRODUCTO,
            TEXT_CAUSA,
            TEXT_TIPIFICACION_2,
            TEXT_TIEMPO_RESPUESTA,
            TEXT_CANAL,
            "creado_por",
            "actualizado",
            TEXT_CODIGO_RESOLUCION,
            "notas_resolucion",
            TEXT_OBSERVACIONES_ADICIONALES,
            TEXT_OBSERVACIONES_TRABAJO,
        ]
        df = df[[col for col in columnas if col in df.columns]]
        st.caption(f"Registros encontrados: {len(df)}")
        st.caption(f"{TEXT_PERIODO}{periodo_label}")
    render_descarga_dataframe(df, "descargar_casos_completos", "casos", periodo_label)
    dataframe_paginado(
        df,
        "vista_casos_tabla",
        reset_token=(periodo_label, filtro_estado, filtro_soporte, filtro_servicio, filtro_cuenta, len(df)),
    )


def vista_cargar_incidentes():
    archivo = st.file_uploader("Sube Excel de incidentes", type=["xlsx"], key="incidentes_upload")
    if archivo:
        df = pd.read_excel(archivo)
        st.write(f"Filas detectadas: {len(df)}")
        st.dataframe(df.head(), use_container_width=True, hide_index=True)
        if st.button("Procesar incidentes"):
            actualizar_progreso = crear_barra_progreso_carga("Procesando incidentes...")
            try:
                cargados, reemplazados = guardar_incidentes(df, progress_callback=actualizar_progreso)
            except Exception as error:
                actualizar_progreso(1, "No se pudo finalizar la carga de incidentes.")
                if es_error_db_transitorio(error):
                    st.error(
                        "La base de datos estaba ocupada por otra carga. Espera unos segundos y vuelve a procesar el archivo."
                    )
                    return
                raise
            if cargados == 0:
                st.error("No se guardaron incidentes. Revisa que el archivo tenga una columna de numero de incidente.")
            else:
                limpiar_cache_datos()
                actualizar_progreso(1, "Carga de incidentes finalizada.")
                st.success(
                    f"Cargados: {cargados} | Registros existentes actualizados: {reemplazados} | "
                    "Los meses anteriores se conservan."
                )


def vista_incidentes():
    anio, mes, periodo_label = selector_periodo_sql("incidents", "vista_incidentes_periodo")
    if not periodo_sql_valido(anio, "incidentes"):
        return
    df = cargar_incidentes_sla_filtrados_cache(anio, mes)
    filtro_estado = TEXT_TODOS
    filtro_tipificacion = TEXT_TODOS
    filtro_servicio = TEXT_TODOS
    filtro_alerta = TEXT_TODOS
    if not df.empty:
        df = preparar_fechas_dashboard(df)
        df["mes"] = df[TEXT_CREADO_DT_DASHBOARD].dt.to_period("M").astype(str).replace("NaT", "Sin fecha")

        filtro_col1, filtro_col2, filtro_col3, filtro_col4 = st.columns([1, 1.4, 1.4, 1])
        with filtro_col1:
            estados = sorted(df[TEXT_ESTADO].dropna().unique().tolist())
            filtro_estado = st.selectbox(TEXT_ESTADO_2, [TEXT_TODOS] + estados, key="estado_inc")
        with filtro_col2:
            filtro_tipificacion = st.selectbox(
                TEXT_TIPIFICACION,
                [TEXT_TODOS] + sorted(df[TEXT_TIPIFICACION_AUTO].dropna().unique().tolist()),
                key="tip_inc",
            )
        with filtro_col3:
            servicios = opciones_filtro_servicio(df, TEXT_SERVICIO_NEGOCIO)
            filtro_servicio = st.selectbox("Servicio", [TEXT_TODOS] + servicios, key="servicio_inc")
        with filtro_col4:
            filtro_alerta = st.selectbox(
                "Es alerta",
                [TEXT_TODOS] + sorted(df[TEXT_ES_ALERTA_AUTO].dropna().unique().tolist()),
                key="alerta_inc",
            )

        filtro_estado_sql = filtro_estado if filtro_estado != TEXT_TODOS else ""
        filtro_tipificacion_sql = filtro_tipificacion if filtro_tipificacion != TEXT_TODOS else ""
        filtro_servicio_sql = (
            filtro_servicio
            if filtro_servicio not in (TEXT_TODOS, SIN_SERVICIO)
            else ""
        )
        filtro_alerta_sql = filtro_alerta if filtro_alerta != TEXT_TODOS else ""
        if filtro_estado_sql or filtro_tipificacion_sql or filtro_servicio_sql or filtro_alerta_sql:
            df = cargar_incidentes_sla_filtrados_cache(
                anio,
                mes,
                "",
                filtro_estado_sql,
                filtro_servicio_sql,
                filtro_tipificacion_sql,
                filtro_alerta_sql,
            )
            df = preparar_fechas_dashboard(df)
            df["mes"] = df[TEXT_CREADO_DT_DASHBOARD].dt.to_period("M").astype(str).replace("NaT", "Sin fecha")

        if filtro_estado != TEXT_TODOS:
            df = df[df[TEXT_ESTADO] == filtro_estado]
        if filtro_tipificacion != TEXT_TODOS:
            df = df[df[TEXT_TIPIFICACION_AUTO] == filtro_tipificacion]
        if filtro_servicio != TEXT_TODOS:
            df = filtrar_por_servicio(df, TEXT_SERVICIO_NEGOCIO, filtro_servicio)
        if filtro_alerta != TEXT_TODOS:
            df = df[df[TEXT_ES_ALERTA_AUTO] == filtro_alerta]
        columnas = [
            TEXT_NUMERO,
            TEXT_SOLICITANTE,
            "categoria",
            TEXT_PRIORIDAD,
            TEXT_ESTADO,
            "mes",
            "grupo_asignacion",
            TEXT_ASIGNADO_A,
            TEXT_DESCRIPCION_2,
            "despues_aprobacion",
            "despues_rechazo",
            "duracion_segundos",
            TEXT_FECHA_VENCIMIENTO_SLA,
            "tipo_falla",
            TEXT_EMPRESA,
            "creado_por",
            TEXT_CERRADO,
            "escalado_proveedor",
            TEXT_SERVICIO_NEGOCIO,
            TEXT_CREADO,
            TEXT_OBSERVACIONES_TRABAJO,
            TEXT_OBSERVACIONES_ADICIONALES,
            "actualizaciones",
            "impacto",
            "lista_notas_trabajo",
            "origen_auto",
            TEXT_TIPIFICACION_AUTO,
            TEXT_TIPO_INCIDENTE_AUTO,
            TEXT_ES_ALERTA_AUTO,
            TEXT_CAUSA_RAIZ_AUTO,
            TEXT_PRIORIDAD_NORMALIZADA,
            "familia_sla",
            TEXT_SLA_OBJETIVO_HORAS,
            TEXT_DURACION_SLA_HORAS,
            TEXT_ESTADO_SLA,
            TEXT_DURACION_HORAS,
        ]
        df = df.drop(columns=[TEXT_CREADO_DT_DASHBOARD], errors="ignore")
        df = df[[col for col in columnas if col in df.columns]]
        st.caption(f"Registros encontrados: {len(df)}")
        st.caption(f"{TEXT_PERIODO}{periodo_label}")
    render_descarga_dataframe(df, "descargar_incidentes_completos", "incidentes", periodo_label)
    dataframe_paginado(
        df,
        "vista_incidentes_tabla",
        reset_token=(periodo_label, filtro_estado, filtro_tipificacion, filtro_servicio, filtro_alerta, len(df)),
    )


def vista_seguimiento_incidentes():
    anio, mes, periodo_label = selector_periodo_sql("incidents", "seguimiento_incidentes_periodo")
    if not periodo_sql_valido(anio, "incidentes"):
        return
    df = cargar_incidentes_filtrados_cache(anio, mes)
    if df.empty:
        st.info(f"No hay incidentes cargados para {periodo_label}.")
        return

    df = preparar_fechas_dashboard(df)
    if df.empty:
        st.info(f"No hay incidentes cargados para {periodo_label}.")
        return

    st.caption(f"{TEXT_PERIODO}{periodo_label}")
    render_seguimiento_operativo_incidentes(df)


def render_tabla_usuarios(usuarios):
    if usuarios.empty:
        st.info("Aun no hay usuarios configurados.")
        return
    tabla = usuarios.copy()
    tabla["active"] = tabla["active"].apply(lambda value: "Activo" if bool(value) else "Inactivo")
    tabla["role"] = tabla["role"].map({TEXT_ADMIN: "Admin", TEXT_VIEWER: "Viewer"}).fillna(tabla["role"])
    st.dataframe(tabla, use_container_width=True, hide_index=True)


def formulario_usuario():
    st.markdown("#### Crear o actualizar usuario")
    with st.form("form_usuario"):
        email = st.text_input("Correo", key="usuario_email")
        password = st.text_input(
            "Contrasena nueva",
            type=TEXT_PASSWORD,
            help="Minimo 8 caracteres. Para actualizar rol/estado sin cambiar contrasena, dejala vacia.",
            key="usuario_password",
        )
        col_rol, col_estado = st.columns(2)
        with col_rol:
            role = st.selectbox("Rol", [TEXT_VIEWER, TEXT_ADMIN], format_func=lambda x: "Viewer" if x == TEXT_VIEWER else "Admin")
        with col_estado:
            active = st.checkbox("Activo", value=True)
        guardar = st.form_submit_button("Guardar usuario")
    return guardar, email, password, role, active


def procesar_formulario_usuario(guardar, email, password, role, active):
    if not guardar:
        return
    if not validar_email(email):
        st.error("Escribe un correo valido.")
        return
    try:
        guardar_usuario(email, password or None, role=role, active=active)
        st.success("Usuario guardado.")
        st.rerun()
    except ValueError as exc:
        st.error(str(exc))


def render_mantenimiento_incidentes():
    st.markdown("#### Mantenimiento de datos")
    total_incidentes = contar_incidentes()
    st.caption(
        f"Incidentes guardados actualmente: {total_incidentes}. "
        "Esta accion elimina solo incidentes; los casos se conservan."
    )
    confirmar_limpieza = st.checkbox(
        "Confirmo que quiero borrar todos los incidentes cargados",
        key="confirmar_limpiar_incidentes",
    )
    clave_limpieza = st.text_input(
        "Clave para limpiar incidentes",
        type=TEXT_PASSWORD,
        key="clave_limpiar_incidentes",
    )
    puede_limpiar = confirmar_limpieza and clave_limpieza == "lina202" and total_incidentes > 0
    if clave_limpieza and clave_limpieza != "lina202":
        st.error("Clave incorrecta.")
    if st.button("Limpiar incidentes", disabled=not puede_limpiar):
        borrados = limpiar_incidentes()
        limpiar_cache_datos()
        st.success(f"Incidentes eliminados: {borrados}. Los casos no fueron modificados.")
        st.rerun()


def candidatos_eliminacion_usuarios(usuarios_actuales):
    email_actual = normalizar_email(st.session_state.get("user"))
    return [
        email
        for email in usuarios_actuales["email"].tolist()
        if normalizar_email(email) != email_actual
    ]


def render_eliminar_usuario():
    st.markdown("#### Quitar acceso")
    usuarios_actuales = listar_usuarios()
    if usuarios_actuales.empty:
        st.info("No hay usuarios para eliminar.")
        return

    candidatos = candidatos_eliminacion_usuarios(usuarios_actuales)
    if not candidatos:
        st.info("No puedes eliminar tu propio usuario desde aqui.")
        return

    usuario_eliminar = st.selectbox("Usuario", candidatos, key="usuario_eliminar")
    confirmar = st.checkbox("Confirmo que quiero eliminar este usuario", key="confirmar_eliminar_usuario")
    if st.button("Eliminar usuario", disabled=not confirmar):
        eliminar_usuario(usuario_eliminar)
        st.success("Usuario eliminado.")
        st.rerun()


def vista_administrar_usuarios():
    st.subheader("Administrar usuarios")
    st.caption("Crea usuarios para dar acceso a los dashboards o cambia su rol y estado.")
    render_tabla_usuarios(listar_usuarios())

    st.divider()
    procesar_formulario_usuario(*formulario_usuario())

    st.divider()
    render_mantenimiento_incidentes()

    st.divider()
    render_eliminar_usuario()

ADMIN_MENU_OPTIONS = [
    "Cargar Excel Casos",
    TEXT_CASOS,
    MENU_DASHBOARD_CASOS_SOPORTE,
    MENU_KPI_CASOS_CLIENTE_EXTERNO,
    "Cargar Excel Incidentes",
    TEXT_INCIDENTES,
    "Dashboard Incidentes",
    MENU_KPI_INCIDENTES,
    MENU_KPI_COMPARATIVO_ANUAL,
    MENU_REINCIDENCIAS_PROBLEMAS,
    MENU_SEGUIMIENTO_RPOST,
    MENU_SEGUIMIENTO_INCIDENTES_ADMIN,
    MENU_KPI_CLIENTES_CLAVE,
    "Clientes Clave",
    "Administrar Usuarios",
]

VIEWER_MENU_OPTIONS = [
    TEXT_CASOS,
    MENU_DASHBOARD_CASOS_SOPORTE,
    MENU_KPI_CASOS_CLIENTE_EXTERNO,
    TEXT_INCIDENTES,
    MENU_KPI_INCIDENTES,
    MENU_KPI_COMPARATIVO_ANUAL,
    MENU_REINCIDENCIAS_PROBLEMAS,
    MENU_SEGUIMIENTO_RPOST,
    MENU_SEGUIMIENTO_INCIDENTES_VIEWER,
    MENU_KPI_CLIENTES_CLAVE,
    MENU_CLIENTES_CLAVE,
]

ADMIN_VIEWS = {
    "Cargar Excel Casos": vista_cargar_casos,
    TEXT_CASOS: vista_casos,
    MENU_DASHBOARD_CASOS_SOPORTE: dashboard_casos,
    MENU_KPI_CASOS_CLIENTE_EXTERNO: dashboard_kpi_casos_cliente_externo,
    "Cargar Excel Incidentes": vista_cargar_incidentes,
    TEXT_INCIDENTES: vista_incidentes,
    "Dashboard Incidentes": dashboard_incidentes,
    MENU_KPI_INCIDENTES: dashboard_kpi_incidentes,
    MENU_KPI_COMPARATIVO_ANUAL: dashboard_kpi_comparativo_anual,
    MENU_REINCIDENCIAS_PROBLEMAS: dashboard_reincidencias_problemas,
    MENU_SEGUIMIENTO_RPOST: dashboard_seguimiento_rpost,
    MENU_SEGUIMIENTO_INCIDENTES_ADMIN: vista_seguimiento_incidentes,
    MENU_KPI_CLIENTES_CLAVE: dashboard_kpi_clientes_clave,
    "Clientes Clave": dashboard_clientes_clave,
    "Administrar Usuarios": vista_administrar_usuarios,
}

VIEWER_VIEWS = {
    TEXT_CASOS: dashboard_casos,
    MENU_DASHBOARD_CASOS_SOPORTE: dashboard_casos,
    MENU_KPI_CASOS_CLIENTE_EXTERNO: dashboard_kpi_casos_cliente_externo,
    TEXT_INCIDENTES: dashboard_incidentes,
    MENU_KPI_INCIDENTES: dashboard_kpi_incidentes,
    MENU_KPI_COMPARATIVO_ANUAL: dashboard_kpi_comparativo_anual,
    MENU_REINCIDENCIAS_PROBLEMAS: dashboard_reincidencias_problemas,
    MENU_SEGUIMIENTO_RPOST: dashboard_seguimiento_rpost,
    MENU_SEGUIMIENTO_INCIDENTES_VIEWER: vista_seguimiento_incidentes,
    MENU_KPI_CLIENTES_CLAVE: dashboard_kpi_clientes_clave,
    MENU_CLIENTES_CLAVE: dashboard_clientes_clave,
}


def cerrar_sesion_boton(en_sidebar=False):
    contenedor = st.sidebar if en_sidebar else st
    if contenedor.button("Cerrar sesion"):
        st.session_state.clear()
        st.rerun()


def render_vista_admin():
    menu = st.sidebar.selectbox("Menu", ADMIN_MENU_OPTIONS)
    st.sidebar.caption(f"Sesion: {st.session_state.user}")
    cerrar_sesion_boton(en_sidebar=True)
    ADMIN_VIEWS[menu]()


def render_vista_viewer():
    cerrar_sesion_boton()
    vista = st.radio(
        "Vista",
        VIEWER_MENU_OPTIONS,
        horizontal=True,
        label_visibility="collapsed",
        key="viewer_vista",
    )
    ejecutar_con_carga(vista, VIEWER_VIEWS[vista])


def run_app():
    aplicar_tema_visual()
    init_db()
    if listar_usuarios().empty:
        configurar_primer_admin()
        return
    if not login():
        return

    if st.session_state.role == TEXT_ADMIN:
        render_vista_admin()
    else:
        render_vista_viewer()
