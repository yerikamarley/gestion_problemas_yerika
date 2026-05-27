import html
import re

import pandas as pd
import plotly.express as px
import streamlit as st

from app_logic import (
    agregar_campos_sla_incidentes,
    autenticar_usuario,
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
    load_incidentes,
    normalizar_fecha,
    normalizar_texto,
    normalizar_email,
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
    "bg": "#fffafa",
    "bg_soft": "#fff3ec",
    "surface": "#fffafa",
    "surface_alt": "#fff4ef",
    "border": "#ead8d1",

    "text": "#141414",
    "muted": "#5a5151",

    TEXT_PRIMARY: TEXT_F35B04,
    "primary_hover": TEXT_F18701,
    TEXT_ORANGE: TEXT_F18701,
    TEXT_YELLOW: "#f7b801",
    "yellow_soft": "#ffe0a1",
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

SLA_CASOS_HORAS = 36

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
CLIENTE_RCI_COLOMBIA = "RCI COLOMBIA S.A COMPAÑÍA DE FINANCIAMIENTO"
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
COL_CUMPLE_SLA = "Cumple SLA"
COL_NO_CUMPLE_SLA = "No cumple SLA"
COL_PROM_HORAS = "Prom. horas"
COL_PROM_DIAS = "Prom. dias"
COL_CAUSA_RAIZ = "Causa raiz"
COL_MOTIVO_CASO = "Motivo del caso"
COL_PRINCIPAL_TIPIFICACION = "Principal tipificacion"
COL_PRINCIPAL_CAUSA_COMUN = "Principal causa comun"
COL_SLA_OBJETIVO = "SLA objetivo"
COL_SLA_OBJETIVO_H = "SLA objetivo h"
COL_MAX_HORAS = "Max. horas"
COL_MAX_DIAS = "Max. dias"
COL_FIRMA_REINCIDENCIA = "Firma reincidencia"
COL_INCIDENTES_REINCIDENTES = "Incidentes reincidentes"
COL_REINCIDENTE = "Reincidente"
COL_REINCIDENTE_AGENDA = "Reincidente agenda"
COL_CASOS_REINCIDENTES_AGENDA = "Casos reincidentes agenda"
COL_REINCIDENCIA_AGENDAMIENTO = "Reincidencia agendamiento %"
MENU_CLIENTES_CLAVE = "Clientes clave"
MENU_KPI_CASOS_CLIENTE_EXTERNO = "KPI Casos Cliente Externo"
MENU_KPI_INCIDENTES = "KPI Incidentes"
MENU_SEGUIMIENTO_INCIDENTES_VIEWER = "Seguimiento incidentes"
MENU_SEGUIMIENTO_INCIDENTES_ADMIN = "Seguimiento Incidentes"
LABEL_CASOS_CLIENTE_EXTERNO = "Casos cliente externo"
CASE_TIPIFICATION_RENAMES = {
    "8 - Instalaciones": TIPIFICACION_REDIRECCIONAMIENTO_AGENDA,
    "8 - Agenda Instalaciones IVR": TIPIFICACION_REDIRECCIONAMIENTO_AGENDA,
    "9 - Agenda Instalaciones IVR": TIPIFICACION_REDIRECCIONAMIENTO_AGENDA,
    "9 - Redireccionamiento Agenda IVR": TIPIFICACION_REDIRECCIONAMIENTO_AGENDA,
    "9 - Agenda sin evidencia": TIPIFICACION_CLIENTE_NO_ASISTIO,
    "10 - Agenda sin evidencia": TIPIFICACION_CLIENTE_NO_ASISTIO,
}

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


def selector_mes_dashboard(df, key, columna_dt=TEXT_CREADO_DT_DASHBOARD):
    meses = meses_disponibles(df, columna_dt)
    if not meses:
        st.caption("No hay fechas validas para filtrar por mes.")
        return TEXT_TODOS
    opciones = [TEXT_TODOS] + meses
    return st.selectbox("Mes del dashboard", opciones, index=len(opciones) - 1, key=key)


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
    CLIENTE_TELEFONICA: [CLIENTE_TELEFONICA, "TELEFÓNICA", "MOVISTAR"],
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
    CLIENTE_COLPENSIONES: [CLIENTE_COLPENSIONES],
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
            background: linear-gradient(180deg, var(--bg-soft) 0%, var(--bg) 42%, var(--surface) 100%);
            color: var(--text) !important;
            color-scheme: light !important;
        }}

        [data-testid="stHeader"], [data-testid="stToolbar"], [data-testid="stDecoration"] {{
            background: transparent !important;
        }}

        [data-testid="stSidebar"] {{
            background: linear-gradient(180deg, var(--surface) 0%, var(--surface-alt) 100%) !important;
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
            font-size: 0.98rem !important;
            line-height: 1.45 !important;
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
            font-weight: 800;
            line-height: 1.25;
            margin-bottom: 12px;
            color: var(--muted);
            text-transform: uppercase;
            letter-spacing: 0;
            max-width: 100%;
            overflow-wrap: anywhere;
        }}

        .kpi-value {{
            font-size: 46px;
            font-weight: 800;
            color: var(--primary);
            line-height: 1.05;
            font-variant-numeric: tabular-nums;
        }}

        .executive-note {{
            background: rgba(255, 250, 250, 0.96);
            border: 1px solid var(--border);
            border-radius: 8px;
            color: var(--text);
            padding: 14px 16px 12px;
            margin: 0.1rem 0 0.25rem;
            line-height: 1.45;
            font-size: 1rem;
            box-shadow: 0 6px 16px rgba(20, 20, 20, 0.04);
        }}

        .executive-note-title {{
            color: var(--primary);
            font-weight: 800;
            margin-bottom: 0.55rem;
        }}

        .executive-note-line {{
            color: var(--muted);
            margin: 0.24rem 0;
        }}

        .executive-note-line strong {{
            color: var(--text);
        }}

        .executive-note-detail {{
            color: var(--muted);
            margin-top: 0.45rem;
            font-size: 0.94rem;
        }}

        .executive-note-detail strong {{
            color: var(--text);
        }}

        .executive-note-conclusion {{
            border-top: 1px solid var(--border);
            color: var(--muted);
            margin-top: 0.7rem;
            padding-top: 0.65rem;
            font-size: 0.94rem;
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
        plot_bgcolor="rgba(255,250,250,0.94)",
        font={"color": UI_PALETTE["text"], "size": 15},
        title_font={"color": UI_PALETTE[TEXT_PRIMARY], "size": 20},
        margin={"l": 12, "r": 12, "t": 52, "b": 12},
        legend={"bgcolor": "rgba(255,250,250,0.82)", "font": {"size": 14}},
    )
    fig.update_xaxes(
        showgrid=True,
        gridcolor="rgba(20, 20, 20, 0.10)",
        zeroline=False,
        tickfont={"size": 14},
        title_font={"size": 15},
        automargin=True,
    )
    fig.update_yaxes(
        showgrid=False,
        zeroline=False,
        tickfont={"size": 14},
        title_font={"size": 15},
        automargin=True,
    )
    return fig


def estilos_login():
    st.markdown(
        """
        <style>
        header {visibility: hidden;}
        footer {visibility: hidden;}

        /* Fondo general */
        .stApp {
            background: linear-gradient(180deg, var(--bg-soft) 0%, var(--bg) 54%, var(--surface) 100%);
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

        /* Título */
        .login-title {
            font-size: 28px;
            font-weight: 800;
            color: var(--primary);
            margin-bottom: 8px;
            text-align: left;
        }

        /* Subtítulo */
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

        /* Botón */
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
    if df.empty or TEXT_ESTADO not in df.columns:
        return pd.Series(False, index=df.index)
    return df[TEXT_ESTADO].apply(normalizar_texto).eq(TEXT_CERRADO)


def mascara_prioridad_alta(df):
    if df.empty or TEXT_PRIORIDAD not in df.columns:
        return pd.Series(False, index=df.index)
    return df[TEXT_PRIORIDAD].fillna("").str.contains(
        r"^(?:1|2\s*-\s*Alta|alta|critica|critico)",
        case=False,
        regex=True,
    )


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
        COL_PRINCIPAL_CAUSA_COMUN: valor_mas_frecuente(trabajo[TEXT_CAUSA_COMUN]),
    }
    return trabajo, metricas


def grafico_barras_kpi(df, x, y, titulo, color):
    if df.empty:
        st.info(f"No hay datos para {titulo.lower()}.")
        return
    grafico = df[df[x] > 0].sort_values(by=x, ascending=True)
    if grafico.empty:
        st.info(f"No hay datos para {titulo.lower()}.")
        return
    fig = px.bar(
        grafico,
        x=x,
        y=y,
        orientation="h",
        text=x,
        color_discrete_sequence=[color],
    )
    fig.update_traces(marker_color=color, textposition=TEXT_OUTSIDE, cliponaxis=False)
    fig.update_layout(height=max(220, 34 * len(grafico) + 110))
    fig = aplicar_estilo_figura(fig, titulo)
    fig.update_layout(margin={"l": 145, "r": 58, "t": 46, "b": 34}, showlegend=False)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def resumen_otras_tipificaciones(base, top_n=3):
    conteo = base["_tipificacion_kpi"].value_counts(dropna=False)
    otras = conteo.iloc[top_n:]
    total_otras = int(otras.sum())
    if total_otras <= 0:
        return "No hay categorias adicionales fuera del top principal."

    principales_otras = ", ".join(f"{tip} ({cantidad})" for tip, cantidad in otras.head(3).items())
    return f"Otras categorias agrupan {total_otras} casos fuera del top 3: {principales_otras}."


def render_lectura_kpi(metricas, base):
    causa_principal = metricas[COL_PRINCIPAL_CAUSA_COMUN]
    detalle_causa = resumen_detalle_causa_principal(base, causa_principal)
    contenido = f"""
    <div class="executive-note">
        <div class="executive-note-title">Lectura</div>
        <div class="executive-note-line">Principal tipificacion: <strong>{html.escape(str(metricas[COL_PRINCIPAL_TIPIFICACION]))}</strong></div>
        <div class="executive-note-line">Causa comun: <strong>{html.escape(str(causa_principal))}</strong></div>
        <div class="executive-note-detail">{html.escape(detalle_causa)}</div>
        <div class="executive-note-conclusion">{html.escape(resumen_otras_tipificaciones(base))}</div>
    </div>
    """
    st.markdown(contenido, unsafe_allow_html=True)


def render_kpi_casos_cliente_externo(df):
    base, metricas = preparar_kpi_casos_cliente_externo(df)
    if base.empty:
        return

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

    col_grafico, col_lectura = st.columns([2.15, 1])
    with col_grafico:
        tipificaciones = conteo_top_con_otras(base["_tipificacion_kpi"], top_n=3)
        grafico_barras_kpi(
            tipificaciones,
            TEXT_CANTIDAD,
            TEXT_TIPOLOGIA,
            "Top 3 tipificaciones + otras",
            UI_PALETTE[TEXT_PRIMARY],
        )
    with col_lectura:
        render_lectura_kpi(metricas, base)

    with st.expander("Detalle completo de casos usados en el calculo"):
        columnas = [
            TEXT_NUMERO,
            TEXT_ESTADO,
            TEXT_CUENTA,
            TEXT_DESCRIPCION_2,
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
        st.dataframe(visible, use_container_width=True, hide_index=True)


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

    metricas = {
        "total": len(trabajo),
        "externos": int((trabajo[TEXT_SEGMENTO] == "Cliente externo").sum()),
        "internos": int((trabajo[TEXT_SEGMENTO] == "Cliente interno").sum()),
        "abiertos": int(trabajo[TEXT_ABIERTO].sum()),
        "cerrados": int(trabajo[TEXT_CERRADO_2].sum()),
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
                TEXT_CANTIDAD,
                "% segmento",
                COL_LECTURA_EJECUTIVA,
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
    if detalle and detalle != "Sin inferencia":
        return f"Casos asociados a {detalle.lower()}."
    return "Causa raiz pendiente de clasificacion en el cierre."


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
        f"{segmento}: la causa principal es {fila[COL_CAUSA_RAIZ]} "
        f"({int(fila[TEXT_CANTIDAD])} casos, {fila['% segmento']}%). "
        f"{fila[COL_LECTURA_EJECUTIVA]} Detalle observado: {fila['Detalle tecnico observado']}."
    )


def render_lectura_kpi_incidentes(causas):
    lectura_externo = texto_lectura_causa_segmento(causas, "Cliente externo")
    lectura_interno = texto_lectura_causa_segmento(causas, "Cliente interno")
    contenido = f"""
    <div class="executive-note">
        <div class="executive-note-title">Lectura</div>
        <div class="executive-note-detail"><strong>Cliente externo:</strong> {html.escape(lectura_externo.replace("Cliente externo: ", ""))}</div>
        <div class="executive-note-detail"><strong>Cliente interno:</strong> {html.escape(lectura_interno.replace("Cliente interno: ", ""))}</div>
    </div>
    """
    st.markdown(contenido, unsafe_allow_html=True)


def render_grafico_causas_kpi_incidentes(causas):
    if causas.empty:
        st.info("No hay causas raiz para graficar en el periodo seleccionado.")
        return

    col_externo, col_interno = st.columns(2)
    segmentos = [
        (col_externo, "Cliente externo", "Causas cliente externo"),
        (col_interno, "Cliente interno", "Causas cliente interno"),
    ]

    for columna, segmento, titulo in segmentos:
        with columna:
            grafico = (
                causas[causas[TEXT_SEGMENTO] == segmento]
                .groupby(COL_LECTURA_EJECUTIVA, as_index=False)
                .agg(Cantidad=(TEXT_CANTIDAD, "sum"))
                .sort_values(by=TEXT_CANTIDAD, ascending=True)
            )
            if grafico.empty:
                st.info(f"No hay causas para {segmento.lower()} en el periodo.")
                continue

            fig = px.bar(
                grafico,
                x=TEXT_CANTIDAD,
                y=COL_LECTURA_EJECUTIVA,
                orientation="h",
                text=TEXT_CANTIDAD,
                color_discrete_sequence=[UI_PALETTE[TEXT_PURPLE]],
            )
            fig.update_traces(marker_color=UI_PALETTE[TEXT_PURPLE], textposition=TEXT_OUTSIDE, cliponaxis=False)
            fig.update_layout(height=max(320, 48 * len(grafico) + 110), yaxis={"automargin": True})
            st.plotly_chart(aplicar_estilo_figura(fig, titulo), use_container_width=True)


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


def render_kpi_incidentes(df):
    base, metricas = preparar_kpi_incidentes(df)
    if base.empty:
        st.info("No hay incidentes cliente interno o externo para el periodo seleccionado.")
        return

    causas = resumen_causas_kpi_incidentes(base)

    st.subheader(MENU_KPI_INCIDENTES)
    render_tarjetas(
        [
            ("Incidentes", metricas["total"]),
            ("Cliente externo", metricas["externos"]),
            ("Cliente interno", metricas["internos"]),
            ("Abiertos", metricas["abiertos"]),
            ("Reincidencia", f"{metricas['tasa_reincidencia']}%"),
            ("SLA incidentes", f"{metricas['cumplimiento_sla']}%"),
        ]
    )
    st.caption(
        f"Cerrados: {metricas['cerrados']} | Reincidentes: {metricas['reincidentes']} | "
        f"Promedio: {metricas['promedio']} h | "
        f"Cumplen SLA: {metricas['cumple_sla']} | No cumplen: {metricas['no_cumple_sla']}"
    )

    col_grafico, col_lectura = st.columns([2.15, 1])
    with col_grafico:
        render_grafico_causas_kpi_incidentes(causas)
    with col_lectura:
        render_lectura_kpi_incidentes(causas)

    tab_resumen, tab_externo, tab_interno, tab_reincidencia, tab_detalle = st.tabs(
        [TEXT_RESUMEN, "Cliente externo", "Cliente interno", "Reincidencia", "Detalle"]
    )
    with tab_resumen:
        st.dataframe(causas, use_container_width=True, hide_index=True)
    with tab_externo:
        st.dataframe(
            causas[causas[TEXT_SEGMENTO] == "Cliente externo"],
            use_container_width=True,
            hide_index=True,
        )
    with tab_interno:
        st.dataframe(
            causas[causas[TEXT_SEGMENTO] == "Cliente interno"],
            use_container_width=True,
            hide_index=True,
        )
    with tab_reincidencia:
        render_reincidencia_kpi_incidentes(base, metricas)
    with tab_detalle:
        columnas = [
            TEXT_NUMERO,
            TEXT_SEGMENTO,
            TEXT_ESTADO,
            TEXT_PRIORIDAD,
            TEXT_SERVICIO_NEGOCIO,
            TEXT_CAUSA_RAIZ_AUTO,
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
        st.dataframe(
            base[[col for col in columnas if col in base.columns]],
            use_container_width=True,
            hide_index=True,
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

def clasificar_causa_incidente(causa):
    texto = normalizar_texto(causa)
    if not texto or "sin inferencia" in texto or "sin patron" in texto:
        return (
            "Sin clasificar",
            "No hay informacion suficiente para explicar la causa.",
            "Completar causa raiz en el cierre del incidente.",
        )

    reglas = [
        (
            ["base de datos", "database", "sql", "bd"],
            "Base de datos",
            "Errores o indisponibilidad asociados a datos o consultas del servicio.",
            "Revisar estabilidad, consultas, bloqueos y eventos recurrentes de base de datos.",
        ),
        (
            ["correo", "notificacion", "smtp", "mail"],
            "Correo / notificaciones",
            "Fallas en envio, recepcion o procesamiento de notificaciones al cliente.",
            "Revisar cola de envio, plantillas, rebotes y proveedores de correo.",
        ),
        (
            ["certificado", "cadena de confianza", "ssl"],
            "Certificados / cadena de confianza",
            "Problemas asociados a certificados, confianza o validacion criptografica.",
            "Validar vigencia, cadena, configuracion y comunicacion preventiva al cliente.",
        ),
        (
            ["firma", "firmar", "validacion", "validar"],
            "Firma digital / validacion",
            "Dificultades para firmar, validar o completar procesos de firma digital.",
            "Revisar flujo de firma, mensajes de error y recurrencia por producto o cliente.",
        ),
        (
            ["duplicad"],
            "Registros duplicados",
            "Casos repetidos o registros duplicados que distorsionan la lectura operativa.",
            "Depurar duplicados y ajustar la regla de clasificacion/cierre.",
        ),
        (
            ["monitoreo", "noc", "caida", "indisponibilidad", "degradacion"],
            "Disponibilidad del servicio",
            "Eventos de disponibilidad, caida o degradacion percibidos por monitoreo o clientes.",
            "Revisar continuidad, ventanas de afectacion y comunicacion a clientes.",
        ),
        (
            ["red", "conectividad", "vpn", "latencia", "enlace"],
            "Conectividad",
            "Problemas de red, comunicacion o acceso al servicio.",
            "Validar conectividad, trazas y posibles dependencias de terceros.",
        ),
    ]

    for palabras, causa_ejecutiva, lectura, accion in reglas:
        if any(palabra in texto for palabra in palabras):
            return causa_ejecutiva, lectura, accion

    return (
        "Otros hallazgos tecnicos",
        "Hallazgos tecnicos con bajo volumen o descripcion no estandarizada.",
        "Normalizar la causa raiz en el cierre para mejorar el analisis mensual.",
    )


def clasificar_causa_cliente_externo(causa):
    return clasificar_causa_incidente(causa)


def resumen_causas_incidentes(df, porcentaje_columna="% incidentes"):
    columnas = [
        COL_CAUSA_RAIZ,
        TEXT_CANTIDAD,
        porcentaje_columna,
        COL_LECTURA_EJECUTIVA,
        COL_ACCION_SUGERIDA,
        "Detalle tecnico observado",
    ]
    if df.empty:
        return pd.DataFrame(columns=columnas)

    trabajo = df.copy()
    trabajo[TEXT_CAUSA_TECNICA] = trabajo[TEXT_CAUSA_RAIZ_AUTO].replace("", pd.NA).fillna("Sin inferencia")
    clasificacion = trabajo[TEXT_CAUSA_TECNICA].apply(clasificar_causa_incidente)
    trabajo[[COL_CAUSA_RAIZ, COL_LECTURA_EJECUTIVA, COL_ACCION_SUGERIDA]] = pd.DataFrame(
        clasificacion.tolist(),
        index=trabajo.index,
    )

    total = len(trabajo)

    def detalles_tecnicos(serie):
        valores = serie.dropna().astype(str).value_counts().head(2).index.tolist()
        return "; ".join(valores)

    resumen = (
        trabajo.groupby([COL_CAUSA_RAIZ, COL_LECTURA_EJECUTIVA, COL_ACCION_SUGERIDA], dropna=False)
        .agg(
            Cantidad=(TEXT_NUMERO, TEXT_COUNT),
            Detalle_tecnico_observado=(TEXT_CAUSA_TECNICA, detalles_tecnicos),
        )
        .reset_index()
        .sort_values(by=[TEXT_CANTIDAD, COL_CAUSA_RAIZ], ascending=[False, True])
    )
    resumen[porcentaje_columna] = resumen[TEXT_CANTIDAD].apply(lambda valor: porcentaje(valor, total))
    resumen = resumen.rename(columns={"Detalle_tecnico_observado": "Detalle tecnico observado"})
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
    df = normalizar_tipificaciones_casos_df(load_casos())
    if df.empty:
        st.info("No hay datos de casos cargados.")
        return

    df_historico = preparar_fechas_dashboard(df)
    mes_dashboard = selector_mes_dashboard(df_historico, "dashboard_casos_mes")
    df = filtrar_mes_dashboard(df_historico, mes_dashboard)
    if df.empty:
        st.info(f"No hay casos cargados para {mes_dashboard}.")
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
    st.caption(f"{TEXT_PERIODO}{mes_dashboard} | Cumplen: {cumplen}{TEXT_NO_CUMPLEN}{incumplen}")

    st.divider()
    col1, col2 = st.columns(2)

    with col1:
        tip = tabla_resumen_tipificaciones_casos(df)[[TEXT_TIPIFICACION, TEXT_CANTIDAD]]
        tip = tip.sort_values(by=TEXT_CANTIDAD, ascending=True)
        fig = px.bar(
            tip,
            x=TEXT_CANTIDAD,
            y=TEXT_TIPIFICACION,
            orientation="h",
            text=TEXT_CANTIDAD,
            color=TEXT_TIPIFICACION,
            color_discrete_sequence=CHART_COLORS,
        )
        fig.update_traces(textposition=TEXT_OUTSIDE)
        fig.update_layout(showlegend=False)
        st.plotly_chart(aplicar_estilo_figura(fig, "Casos por tipificacion"), use_container_width=True)

    with col2:
        serie = df.copy()
        casos_dia = serie.groupby(serie[TEXT_CREADO_DT_DASHBOARD].dt.date).size().reset_index(name=TEXT_CASOS_2)
        casos_dia.columns = [TEXT_FECHA, TEXT_CASOS_2]
        fig = px.bar(casos_dia, x=TEXT_FECHA, y=TEXT_CASOS_2, color_discrete_sequence=[UI_PALETTE[TEXT_YELLOW]])
        fig.update_traces(marker_color=UI_PALETTE[TEXT_YELLOW])
        st.plotly_chart(aplicar_estilo_figura(fig, "Casos por dia"), use_container_width=True)

    st.divider()
    render_analisis_agendamiento_mesa(df, df_historico, mes_dashboard)

    st.divider()
    st.subheader("Resumen de tipificaciones")
    st.caption("Descripcion breve de cada tipificacion y cantidad actual de casos clasificados en el dashboard.")
    st.dataframe(tabla_resumen_tipificaciones_casos(df), use_container_width=True, hide_index=True)

    render_seguimiento_casos(df)


def dashboard_kpi_casos_cliente_externo():
    df = normalizar_tipificaciones_casos_df(load_casos())
    if df.empty:
        st.info("No hay datos de casos cargados.")
        return

    df = preparar_fechas_dashboard(df)
    mes_dashboard = selector_mes_dashboard(df, "kpi_casos_cliente_externo_mes")
    df = filtrar_mes_dashboard(df, mes_dashboard)
    if df.empty:
        st.info(f"No hay casos cargados para {mes_dashboard}.")
        return

    st.caption(f"{TEXT_PERIODO}{mes_dashboard}")
    render_kpi_casos_cliente_externo(df)


def dashboard_kpi_incidentes():
    df = load_incidentes()
    if df.empty:
        st.info("No hay datos de incidentes cargados.")
        return

    df = preparar_fechas_dashboard(df)
    mes_dashboard = selector_mes_dashboard(df, "kpi_incidentes_mes")
    df = filtrar_mes_dashboard(df, mes_dashboard)
    if df.empty:
        st.info(f"No hay incidentes cargados para {mes_dashboard}.")
        return

    st.caption(f"{TEXT_PERIODO}{mes_dashboard}")
    render_kpi_incidentes(df)


def dashboard_incidentes():
    df = load_incidentes()
    if df.empty:
        st.info("No hay datos de incidentes cargados.")
        return

    df = preparar_fechas_dashboard(df)
    mes_dashboard = selector_mes_dashboard(df, "dashboard_incidentes_mes")
    df = filtrar_mes_dashboard(df, mes_dashboard)
    if df.empty:
        st.info(f"No hay incidentes cargados para {mes_dashboard}.")
        return
    df = agregar_campos_sla_incidentes(df)

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
        f"{TEXT_PERIODO}{mes_dashboard} | Cerrados: {cerrados} | Abiertos: {abiertos} | "
        f"Promedio incidentes: {promedio}h | Cumplen: {cumplen}{TEXT_NO_CUMPLEN}{incumplen} | "
        f"Externos: {len(incidentes_externos)} | Internos: {len(incidentes_internos)} | "
        f"SLA casos cliente externo: {casos_sla}%"
    )

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


def fechas_clientes_clave(casos, incidentes):
    fechas = []
    if not casos.empty:
        fechas.append(casos[TEXT_CREADO_DT])
    if not incidentes.empty:
        fechas.append(incidentes[TEXT_CREADO_DT])
    return pd.concat(fechas).dropna() if fechas else pd.Series(dtype=PANDAS_DATETIME_DTYPE)


def seleccionar_filtros_clientes_clave(fechas):
    filtro_col1, filtro_col2, filtro_col3 = st.columns([2, 1, 1])
    with filtro_col1:
        clientes_seleccionados = st.multiselect(
            "Clientes",
            CLIENTES_CLAVE,
            default=CLIENTES_CLAVE,
            key="clientes_clave_filtro",
        )
    with filtro_col2:
        base_meses = pd.DataFrame({TEXT_CREADO_DT: fechas})
        mes_dashboard = selector_mes_dashboard(base_meses, "clientes_clave_mes", TEXT_CREADO_DT)
    with filtro_col3:
        rango_fechas = selector_rango_fechas_clientes(fechas)
    return clientes_seleccionados, mes_dashboard, rango_fechas


def selector_rango_fechas_clientes(fechas):
    if fechas.empty:
        return None
    fecha_min = fechas.min().date()
    fecha_max = fechas.max().date()
    return st.date_input(
        "Rango de fechas",
        value=(fecha_min, fecha_max),
        min_value=fecha_min,
        max_value=fecha_max,
        key="clientes_clave_rango",
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


def render_grafico_atenciones_cliente(resumen_actividad):
    grafico = resumen_actividad.sort_values(by=COL_TOTAL_ATENCIONES, ascending=True)
    fig = px.bar(
        grafico,
        x=COL_TOTAL_ATENCIONES,
        y=TEXT_CLIENTE,
        orientation="h",
        text=COL_TOTAL_ATENCIONES,
        color=TEXT_NIVEL,
        color_discrete_map={
            "Verde": UI_PALETTE[TEXT_LAVENDER],
            TEXT_AMARILLO: UI_PALETTE[TEXT_YELLOW],
            "Rojo": UI_PALETTE[TEXT_PRIMARY],
        },
    )
    fig.update_traces(textposition=TEXT_OUTSIDE)
    st.plotly_chart(aplicar_estilo_figura(fig, "Atenciones por cliente clave"), use_container_width=True)


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
        st.dataframe(abiertos_detalle, use_container_width=True, hide_index=True)


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
    st.dataframe(
        casos.sort_values(by=TEXT_CREADO_DT, ascending=False)[columnas_casos],
        use_container_width=True,
        hide_index=True,
    )


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
    st.dataframe(
        incidentes.sort_values(by=TEXT_CREADO_DT, ascending=False)[columnas_incidentes],
        use_container_width=True,
        hide_index=True,
    )


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


def dashboard_clientes_clave():
    casos = preparar_casos_clientes_clave(load_casos())
    incidentes = preparar_incidentes_clientes_clave(load_incidentes())
    st.subheader(MENU_CLIENTES_CLAVE)

    fechas = fechas_clientes_clave(casos, incidentes)
    clientes_seleccionados, mes_dashboard, rango_fechas = seleccionar_filtros_clientes_clave(fechas)
    if not clientes_seleccionados:
        st.warning("Selecciona al menos un cliente clave.")
        return

    casos, incidentes = aplicar_filtros_clientes_clave(
        casos,
        incidentes,
        clientes_seleccionados,
        mes_dashboard,
        rango_fechas,
    )
    resumen = resumen_clientes_clave(casos, incidentes)
    resumen = resumen[resumen[TEXT_CLIENTE].isin(clientes_seleccionados)].copy()
    resumen_actividad = resumen[resumen[COL_TOTAL_ATENCIONES] > 0].copy()

    metricas = metricas_dashboard_clientes(casos, incidentes, resumen_actividad)
    render_kpis_clientes_clave(metricas, mes_dashboard)
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


def procesar_archivo_casos(df, reemplazar_meses):
    try:
        cargados, reemplazados, eliminados, meses_reemplazados, duplicados_archivo = guardar_casos(
            df,
            reemplazar_meses=reemplazar_meses,
        )
    except Exception as error:
        if es_error_db_transitorio(error):
            st.error(
                "La base de datos estaba ocupada por otra carga. Espera unos segundos y vuelve a procesar el archivo."
            )
            return
        raise
    if cargados == 0:
        st.error("No se guardaron casos. Revisa que el archivo tenga una columna de numero de caso.")
        return
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
    df = normalizar_tipificaciones_casos_df(load_casos())
    if not df.empty:
        df = preparar_fechas_dashboard(df)
        df["mes"] = df[TEXT_CREADO_DT_DASHBOARD].dt.to_period("M").astype(str).replace("NaT", "Sin fecha")

        filtro_col1, filtro_col2, filtro_col3, filtro_col4, filtro_col5 = st.columns([1, 1, 1.5, 1.5, 2])
        with filtro_col1:
            filtro_mes = selector_mes_dashboard(df, "vista_casos_mes")
        if filtro_mes != TEXT_TODOS:
            df = filtrar_mes_dashboard(df, filtro_mes)

        with filtro_col2:
            estados = sorted(df[TEXT_ESTADO].dropna().unique().tolist())
            filtro_estado = st.selectbox(TEXT_ESTADO_2, [TEXT_TODOS] + estados, key="estado_casos")
        with filtro_col3:
            clasificaciones = sorted(df[TEXT_TIPIFICACION_2].dropna().unique().tolist())
            filtro_clasificacion = st.selectbox(
                TEXT_CLASIFICACION,
                [TEXT_TODOS] + clasificaciones,
                key="clasificacion_casos",
            )
        with filtro_col4:
            servicios = opciones_filtro_servicio(df, TEXT_PRODUCTO)
            filtro_servicio = st.selectbox("Servicio", [TEXT_TODOS] + servicios, key="servicio_casos")
        with filtro_col5:
            filtro_cuenta = st.text_input("Cuenta", key="cuenta_casos")

        if filtro_estado != TEXT_TODOS:
            df = df[df[TEXT_ESTADO] == filtro_estado]
        if filtro_clasificacion != TEXT_TODOS:
            df = df[df[TEXT_TIPIFICACION_2] == filtro_clasificacion]
        if filtro_servicio != TEXT_TODOS:
            df = filtrar_por_servicio(df, TEXT_PRODUCTO, filtro_servicio)
        if filtro_cuenta:
            df = df[df[TEXT_CUENTA].fillna("").str.contains(filtro_cuenta, case=False, na=False)]
        df = df.drop(columns=[TEXT_CREADO_DT_DASHBOARD], errors="ignore")
        columnas = [
            TEXT_NUMERO,
            TEXT_ESTADO,
            "mes",
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
    st.dataframe(df, use_container_width=True, hide_index=True)


def vista_cargar_incidentes():
    archivo = st.file_uploader("Sube Excel de incidentes", type=["xlsx"], key="incidentes_upload")
    if archivo:
        df = pd.read_excel(archivo)
        st.write(f"Filas detectadas: {len(df)}")
        st.dataframe(df.head(), use_container_width=True, hide_index=True)
        if st.button("Procesar incidentes"):
            cargados, reemplazados = guardar_incidentes(df)
            if cargados == 0:
                st.error("No se guardaron incidentes. Revisa que el archivo tenga una columna de numero de incidente.")
            else:
                st.success(
                    f"Cargados: {cargados} | Registros existentes actualizados: {reemplazados} | "
                    "Los meses anteriores se conservan."
                )


def vista_incidentes():
    df = agregar_campos_sla_incidentes(load_incidentes())
    if not df.empty:
        df = preparar_fechas_dashboard(df)
        df["mes"] = df[TEXT_CREADO_DT_DASHBOARD].dt.to_period("M").astype(str).replace("NaT", "Sin fecha")

        filtro_col1, filtro_col2, filtro_col3, filtro_col4, filtro_col5 = st.columns([1, 1, 1.4, 1.4, 1])
        with filtro_col1:
            filtro_mes = selector_mes_dashboard(df, "vista_incidentes_mes")
        if filtro_mes != TEXT_TODOS:
            df = filtrar_mes_dashboard(df, filtro_mes)

        with filtro_col2:
            estados = sorted(df[TEXT_ESTADO].dropna().unique().tolist())
            filtro_estado = st.selectbox(TEXT_ESTADO_2, [TEXT_TODOS] + estados, key="estado_inc")
        with filtro_col3:
            filtro_tipificacion = st.selectbox(
                TEXT_TIPIFICACION,
                [TEXT_TODOS] + sorted(df[TEXT_TIPIFICACION_AUTO].dropna().unique().tolist()),
                key="tip_inc",
            )
        with filtro_col4:
            servicios = opciones_filtro_servicio(df, TEXT_SERVICIO_NEGOCIO)
            filtro_servicio = st.selectbox("Servicio", [TEXT_TODOS] + servicios, key="servicio_inc")
        with filtro_col5:
            filtro_alerta = st.selectbox(
                "Es alerta",
                [TEXT_TODOS] + sorted(df[TEXT_ES_ALERTA_AUTO].dropna().unique().tolist()),
                key="alerta_inc",
            )

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
    st.dataframe(df, use_container_width=True, hide_index=True)


def vista_seguimiento_incidentes():
    df = load_incidentes()
    if df.empty:
        st.info("No hay incidentes cargados.")
        return

    df = preparar_fechas_dashboard(df)
    mes_dashboard = selector_mes_dashboard(df, "seguimiento_incidentes_mes")
    df = filtrar_mes_dashboard(df, mes_dashboard)
    if df.empty:
        st.info(f"No hay incidentes cargados para {mes_dashboard}.")
        return

    st.caption(f"{TEXT_PERIODO}{mes_dashboard}")
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
    "Dashboard Casos",
    MENU_KPI_CASOS_CLIENTE_EXTERNO,
    "Cargar Excel Incidentes",
    TEXT_INCIDENTES,
    "Dashboard Incidentes",
    MENU_KPI_INCIDENTES,
    MENU_SEGUIMIENTO_INCIDENTES_ADMIN,
    "Clientes Clave",
    "Administrar Usuarios",
]

VIEWER_MENU_OPTIONS = [
    TEXT_CASOS,
    MENU_KPI_CASOS_CLIENTE_EXTERNO,
    TEXT_INCIDENTES,
    MENU_KPI_INCIDENTES,
    MENU_SEGUIMIENTO_INCIDENTES_VIEWER,
    MENU_CLIENTES_CLAVE,
]

ADMIN_VIEWS = {
    "Cargar Excel Casos": vista_cargar_casos,
    TEXT_CASOS: vista_casos,
    "Dashboard Casos": dashboard_casos,
    MENU_KPI_CASOS_CLIENTE_EXTERNO: dashboard_kpi_casos_cliente_externo,
    "Cargar Excel Incidentes": vista_cargar_incidentes,
    TEXT_INCIDENTES: vista_incidentes,
    "Dashboard Incidentes": dashboard_incidentes,
    MENU_KPI_INCIDENTES: dashboard_kpi_incidentes,
    MENU_SEGUIMIENTO_INCIDENTES_ADMIN: vista_seguimiento_incidentes,
    "Clientes Clave": dashboard_clientes_clave,
    "Administrar Usuarios": vista_administrar_usuarios,
}

VIEWER_VIEWS = {
    TEXT_CASOS: dashboard_casos,
    MENU_KPI_CASOS_CLIENTE_EXTERNO: dashboard_kpi_casos_cliente_externo,
    TEXT_INCIDENTES: dashboard_incidentes,
    MENU_KPI_INCIDENTES: dashboard_kpi_incidentes,
    MENU_SEGUIMIENTO_INCIDENTES_VIEWER: vista_seguimiento_incidentes,
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
