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

UI_PALETTE = {
    "bg": "#fffafa",
    "bg_soft": "#fff3ec",
    "surface": "#fffafa",
    "surface_alt": "#fff4ef",
    "border": "#ead8d1",

    "text": "#141414",
    "muted": "#5a5151",

    "primary": "#f35b04",
    "primary_hover": "#f18701",
    "orange": "#f18701",
    "yellow": "#f7b801",
    "yellow_soft": "#ffe0a1",
    "lavender": "#9683ec",
    "purple": "#5d16a6",

    "red": "#f35b04",
    "red_soft": "#f18701",
    "green": "#f35b04",
    "green_soft": "#f18701",
    "blue": "#9683ec",
    "blue_soft": "#5d16a6",
}

CHART_COLORS = [
    UI_PALETTE["primary"],
    UI_PALETTE["yellow"],
    UI_PALETTE["orange"],
    UI_PALETTE["lavender"],
    UI_PALETTE["purple"],
    UI_PALETTE["text"],
]

SLA_CASOS_HORAS = 36
CASE_TIPIFICATION_RENAMES = {
    "8 - Instalaciones": "9 - Redireccionamiento Agenda",
    "8 - Agenda Instalaciones IVR": "9 - Redireccionamiento Agenda",
    "9 - Agenda Instalaciones IVR": "9 - Redireccionamiento Agenda",
    "9 - Redireccionamiento Agenda IVR": "9 - Redireccionamiento Agenda",
    "9 - Agenda sin evidencia": "10 - Cliente no asistio",
    "10 - Agenda sin evidencia": "10 - Cliente no asistio",
}


def normalizar_tipificaciones_casos_df(df):
    if df.empty or "tipificacion" not in df.columns:
        return df
    trabajo = df.copy()
    trabajo["tipificacion"] = trabajo["tipificacion"].replace(CASE_TIPIFICATION_RENAMES)
    return trabajo


def preparar_fechas_dashboard(df, columna="creado"):
    trabajo = df.copy()
    trabajo["_creado_dt_dashboard"] = pd.to_datetime(
        trabajo[columna].apply(normalizar_fecha),
        errors="coerce",
    )
    return trabajo


def meses_disponibles(df, columna_dt="_creado_dt_dashboard"):
    if df.empty or columna_dt not in df.columns:
        return []
    meses = df[columna_dt].dropna().dt.to_period("M").astype(str).sort_values().unique().tolist()
    return meses


def selector_mes_dashboard(df, key, columna_dt="_creado_dt_dashboard"):
    meses = meses_disponibles(df, columna_dt)
    if not meses:
        st.caption("No hay fechas validas para filtrar por mes.")
        return "Todos"
    opciones = ["Todos"] + meses
    return st.selectbox("Mes del dashboard", opciones, index=len(opciones) - 1, key=key)


def filtrar_mes_dashboard(df, mes, columna_dt="_creado_dt_dashboard"):
    if mes == "Todos" or df.empty or columna_dt not in df.columns:
        return df
    return df[df[columna_dt].dt.to_period("M").astype(str) == mes].copy()


CASE_TIPIFICATION_GUIDE = [
    {
        "Tipificacion": "1 - phishing",
        "Descripcion": "Correos sospechosos, suplantacion, enlaces fraudulentos o reportes de phishing.",
    },
    {
        "Tipificacion": "2 - Soporte Uso",
        "Descripcion": "Dudas de uso, acompanamiento funcional, configuracion, orientacion y paso a paso.",
    },
    {
        "Tipificacion": "3 - Soporte Falla",
        "Descripcion": "Errores, fallas, caidas, lentitud, indisponibilidad o afectaciones tecnicas.",
    },
    {
        "Tipificacion": "4 - solicitudes",
        "Descripcion": "Solicitudes operativas o comerciales como certificados, biometria, pagos u otros tramites.",
    },
    {
        "Tipificacion": "5 - incidente",
        "Descripcion": "Casos marcados como incidente o con afectacion operativa reportada.",
    },
    {
        "Tipificacion": "6 - Plataformas Ext",
        "Descripcion": "Problemas relacionados con plataformas externas como Adobe, Autofirma o DocuSign.",
    },
    {
        "Tipificacion": "7 - No Aplica",
        "Descripcion": "Casos sin informacion suficiente o que no encajan en las reglas definidas.",
    },
    {
        "Tipificacion": "9 - Redireccionamiento Agenda",
        "Descripcion": "Casos detectados como instalacion que deben redirigirse a agenda.",
    },
    {
        "Tipificacion": "10 - Cliente no asistio",
        "Descripcion": "Cliente no conectado, no ingreso o no se presento a la agenda.",
    },
]

AGENDA_CASE_TIPIFICATION = "9 - Redireccionamiento Agenda"
AGENDA_DIRECT_CHANNEL_HINTS = ["calendario"]
AGENDA_HELP_DESK_CHANNELS = {"web", "telefono", "correo electronico", "en persona", ""}
AGENDA_REASON_RULES = [
    (
        "Token fisico / ePass",
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

INCIDENT_TIPIFICATION_GUIDE = {
    "Alerta NOC": "Senales de monitoreo o eventos detectados por NOC que requieren validacion y seguimiento.",
    "Consulta NOC": "Revisiones o validaciones del NOC sin afectacion confirmada.",
    "Incidente Cliente Externo": "Afectaciones reportadas por clientes externos o con impacto hacia clientes externos.",
    "Incidente Interno": "Afectaciones de infraestructura, red, plataforma o procesos internos.",
    "Incidente Seguridad": "Eventos asociados a confidencialidad, integridad, disponibilidad o riesgo de seguridad.",
    "Consulta Cliente": "Consultas de cliente interno o externo sin afectacion confirmada.",
    "Caso Cliente Externo": "Registros cargados como incidente que corresponden a solicitud, tramite, descarga, guia o instalacion.",
}

INCIDENT_TIPIFICATION_ORDER = [
    "Incidente Cliente Externo",
    "Incidente Interno",
    "Incidente Seguridad",
    "Alerta NOC",
    "Consulta NOC",
    "Consulta Cliente",
    "Caso Cliente Externo",
]

CLIENTES_CLAVE = [
    "SICOV",
    "TELEFONICA",
    "TUYA",
    "SUFI BANCOLOMBIA",
    "RCI COLOMBIA S.A COMPAÑÍA DE FINANCIAMIENTO",
    "PORVENIR",
    "MIBANCO S.A.",
    "BBVA",
    "BANCOOMEVA",
    "BANCOLOMBIA",
    "BANCO POPULAR",
    "BANCO FALABELLA",
    "BANCO DE OCCIDENTE",
    "BANCO DAVIVIENDA S.A.",
    "BANCO CAJA SOCIAL",
    "AV VILLAS",
    "FALLABELLA",
    "COLPENSIONES",
    "CLARO",
    "Coopcentral",
]

CLIENTES_CLAVE_ALIASES = {
    "SICOV": ["SICOV"],
    "TELEFONICA": ["TELEFONICA", "TELEFÓNICA", "MOVISTAR"],
    "TUYA": ["TUYA", "TUYA S.A"],
    "SUFI BANCOLOMBIA": ["SUFI BANCOLOMBIA", "SUFI"],
    "RCI COLOMBIA S.A COMPAÑÍA DE FINANCIAMIENTO": [
        "RCI COLOMBIA S.A COMPAÑÍA DE FINANCIAMIENTO",
        "RCI COLOMBIA",
        "RCI",
    ],
    "PORVENIR": ["PORVENIR"],
    "MIBANCO S.A.": ["MIBANCO S.A.", "MIBANCO"],
    "BBVA": ["BBVA"],
    "BANCOOMEVA": ["BANCOOMEVA", "BANCOOMEVA S.A"],
    "BANCOLOMBIA": ["BANCOLOMBIA", "BANCOLOMBIA S.A"],
    "BANCO POPULAR": ["BANCO POPULAR"],
    "BANCO FALABELLA": ["BANCO FALABELLA"],
    "BANCO DE OCCIDENTE": ["BANCO DE OCCIDENTE"],
    "BANCO DAVIVIENDA S.A.": ["BANCO DAVIVIENDA S.A.", "BANCO DAVIVIENDA", "DAVIVIENDA"],
    "BANCO CAJA SOCIAL": ["BANCO CAJA SOCIAL", "CAJA SOCIAL"],
    "AV VILLAS": ["AV VILLAS", "BANCO AV VILLAS"],
    "FALLABELLA": ["FALLABELLA", "FALABELLA"],
    "COLPENSIONES": ["COLPENSIONES"],
    "CLARO": ["CLARO", "COMCEL"],
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
            --primary: {UI_PALETTE["primary"]};
            --primary-hover: {UI_PALETTE["primary_hover"]};
            --orange: {UI_PALETTE["orange"]};
            --yellow: {UI_PALETTE["yellow"]};
            --yellow-soft: {UI_PALETTE["yellow_soft"]};
            --lavender: {UI_PALETTE["lavender"]};
            --purple: {UI_PALETTE["purple"]};
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
            grid-template-columns: repeat(5, minmax(0, 1fr));
            gap: 1rem;
            margin: 0.35rem 0 0.25rem;
        }}

        .kpi-card {{
            background: var(--surface);
            padding: 18px;
            border-radius: 8px;
            text-align: center;
            color: var(--text);
            min-height: 112px;
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
            font-size: 13px;
            font-weight: 700;
            margin-bottom: 8px;
            color: var(--muted);
            text-transform: uppercase;
            letter-spacing: 0;
        }}

        .kpi-value {{
            font-size: 28px;
            font-weight: 800;
            color: var(--primary);
            line-height: 1.1;
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
                min-height: 88px;
                padding: 14px;
            }}

            .kpi-title {{
                font-size: 11px;
                margin-bottom: 6px;
            }}

            .kpi-value {{
                font-size: 24px;
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
        font=dict(color=UI_PALETTE["text"]),
        title_font=dict(color=UI_PALETTE["primary"], size=17),
        margin=dict(l=12, r=12, t=52, b=12),
        legend=dict(bgcolor="rgba(255,250,250,0.82)"),
    )
    fig.update_xaxes(showgrid=True, gridcolor="rgba(20, 20, 20, 0.10)", zeroline=False)
    fig.update_yaxes(showgrid=False, zeroline=False)
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
        password = st.text_input("Contrasena", type="password", key="password_login")

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
        password = st.text_input("Contrasena", type="password", key="setup_admin_password")
        confirmar = st.text_input("Confirmar contrasena", type="password", key="setup_admin_password_confirm")

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

            guardar_usuario(correo, password, role="admin", active=True)
            st.session_state.user = normalizar_email(correo)
            st.session_state.role = "admin"
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
    "Sin duracion",
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
        return f"{horas_texto} / {dias:.2f} dias"
    dias_texto = f"{int(dias)} dias" if float(dias).is_integer() else f"{dias:.2f} dias"
    return f"{horas_texto} / {dias_texto}"


def clasificar_rango_resolucion(horas):
    try:
        horas = float(horas)
    except (TypeError, ValueError):
        return "Sin duracion"
    if pd.isna(horas):
        return "Sin duracion"
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
    if df.empty or "estado" not in df.columns:
        return pd.Series(False, index=df.index)
    return df["estado"].apply(normalizar_texto).eq("cerrado")


def mascara_prioridad_alta(df):
    if df.empty or "prioridad" not in df.columns:
        return pd.Series(False, index=df.index)
    return df["prioridad"].fillna("").str.contains(
        r"^(?:1|2\s*-\s*Alta|alta|critica|critico)",
        case=False,
        regex=True,
    )


def preparar_seguimiento_operativo_incidentes(df, horas_proximas=24):
    trabajo = agregar_campos_sla_incidentes(df)
    if trabajo.empty:
        return trabajo

    ahora = pd.Timestamp.now()
    trabajo["_cerrado"] = mascara_cerrados(trabajo)
    trabajo["_abierto"] = ~trabajo["_cerrado"]
    trabajo["_creado_dt"] = pd.to_datetime(trabajo["creado"].apply(normalizar_fecha), errors="coerce")
    trabajo["_vencimiento_dt"] = pd.to_datetime(
        trabajo["fecha_vencimiento_sla"].apply(normalizar_fecha),
        errors="coerce",
    )
    trabajo["_horas_abierto"] = ((ahora - trabajo["_creado_dt"]).dt.total_seconds() / 3600).round(2)
    trabajo["_horas_para_vencer_sistema"] = ((trabajo["_vencimiento_dt"] - ahora).dt.total_seconds() / 3600).round(2)
    trabajo["_horas_para_vencer_matriz"] = (trabajo["sla_objetivo_horas"] - trabajo["_horas_abierto"]).round(2)
    trabajo["_horas_para_vencer"] = trabajo["_horas_para_vencer_matriz"].where(
        trabajo["sla_objetivo_horas"].notna(),
        trabajo["_horas_para_vencer_sistema"],
    )
    trabajo["_vencido"] = trabajo["_abierto"] & trabajo["_horas_para_vencer"].notna() & (trabajo["_horas_para_vencer"] < 0)
    trabajo["_proximo_vencer"] = (
        trabajo["_abierto"]
        & trabajo["_horas_para_vencer"].notna()
        & trabajo["_horas_para_vencer"].between(0, horas_proximas, inclusive="both")
    )
    trabajo["_prioridad_alta"] = mascara_prioridad_alta(trabajo)
    trabajo["_alerta"] = trabajo["es_alerta_auto"].fillna("No").eq("Si")
    trabajo["_cliente_externo"] = trabajo["tipificacion_auto"].fillna("").isin(
        ["Incidente Cliente Externo", "Caso Cliente Externo"]
    )
    trabajo["_requiere_seguimiento"] = (
        trabajo["_abierto"]
        & (
            trabajo["_vencido"]
            | trabajo["_proximo_vencer"]
            | trabajo["_prioridad_alta"]
            | trabajo["_alerta"]
            | trabajo["_cliente_externo"]
        )
    )
    return trabajo


def render_seguimiento_operativo_incidentes(df):
    seguimiento = preparar_seguimiento_operativo_incidentes(df)
    if seguimiento.empty:
        return

    abiertos = seguimiento[seguimiento["_abierto"]].copy()
    vencidos = seguimiento[seguimiento["_vencido"]].copy()
    proximos = seguimiento[seguimiento["_proximo_vencer"]].copy()
    prioridad_alta = seguimiento[seguimiento["_abierto"] & seguimiento["_prioridad_alta"]].copy()
    alertas_abiertas = seguimiento[seguimiento["_abierto"] & seguimiento["_alerta"]].copy()
    cliente_externo_abierto = seguimiento[seguimiento["_abierto"] & seguimiento["_cliente_externo"]].copy()

    st.subheader("Seguimiento operativo")
    st.caption(
        "Vista de control para incidentes abiertos que requieren accion: vencidos por SLA, proximos a vencer "
        "en 24 horas, prioridad alta, alertas abiertas o afectacion de cliente externo."
    )
    render_tarjetas(
        [
            ("Abiertos", len(abiertos)),
            ("Vencidos", len(vencidos)),
            ("Proximos 24h", len(proximos)),
            ("Prioridad alta", len(prioridad_alta)),
            ("Alertas abiertas", len(alertas_abiertas)),
        ]
    )
    st.caption(
        "Los vencidos y proximos a vencer se calculan primero con la matriz SLA por prioridad; "
        "si no hay objetivo configurado, se usa la fecha de vencimiento del sistema."
    )

    columnas = [
        "numero",
        "estado",
        "prioridad",
        "prioridad_normalizada",
        "sla_objetivo_horas",
        "estado_sla",
        "grupo_asignacion",
        "asignado_a",
        "empresa",
        "servicio_negocio",
        "creado",
        "fecha_vencimiento_sla",
        "_horas_abierto",
        "_horas_para_vencer",
        "tipificacion_auto",
        "es_alerta_auto",
        "causa_raiz_auto",
    ]
    etiquetas = {
        "_horas_abierto": "horas_abierto",
        "_horas_para_vencer": "horas_para_vencer",
    }

    tab_vencidos, tab_proximos, tab_prioridad, tab_alertas, tab_cliente = st.tabs(
        ["Vencidos", "Proximos 24h", "Prioridad alta", "Alertas", "Cliente externo"]
    )

    tablas = [
        (tab_vencidos, vencidos.sort_values(by="_horas_para_vencer")),
        (tab_proximos, proximos.sort_values(by="_horas_para_vencer")),
        (tab_prioridad, prioridad_alta.sort_values(by="_horas_abierto", ascending=False)),
        (tab_alertas, alertas_abiertas.sort_values(by="_horas_abierto", ascending=False)),
        (tab_cliente, cliente_externo_abierto.sort_values(by="_horas_abierto", ascending=False)),
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
    trabajo["_cerrado"] = mascara_cerrados(trabajo)
    trabajo["_abierto"] = ~trabajo["_cerrado"]
    trabajo["_creado_dt"] = pd.to_datetime(trabajo["creado"].apply(normalizar_fecha), errors="coerce")
    trabajo["_horas_abierto"] = trabajo["_creado_dt"].apply(lambda fecha: horas_habiles_entre(fecha, ahora))
    trabajo["_horas_para_vencer"] = (SLA_CASOS_HORAS - trabajo["_horas_abierto"]).round(2)
    trabajo["_vencido"] = trabajo["_abierto"] & trabajo["_creado_dt"].notna() & (trabajo["_horas_para_vencer"] < 0)
    trabajo["_proximo_vencer"] = (
        trabajo["_abierto"]
        & trabajo["_creado_dt"].notna()
        & trabajo["_horas_para_vencer"].between(0, horas_proximas, inclusive="both")
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

    abiertos = seguimiento[seguimiento["_abierto"]].copy()
    vencidos = seguimiento[seguimiento["_vencido"]].copy()
    proximos = seguimiento[seguimiento["_proximo_vencer"]].copy()

    st.divider()
    st.subheader("Control de vencimiento")
    st.caption(
        f"Calculado solo sobre casos abiertos con el mismo criterio de horas habiles del SLA de {SLA_CASOS_HORAS} horas. "
        "El SLA superior resume casos cerrados; esta tabla muestra pendientes abiertos."
    )
    render_tarjetas(
        [
            ("Abiertos", len(abiertos)),
            ("Vencidos", len(vencidos)),
            ("Proximos 12h", len(proximos)),
        ]
    )

    columnas = [
        "numero",
        "estado",
        "prioridad",
        "cuenta",
        "contacto",
        "asignado",
        "creado",
        "_horas_abierto",
        "_horas_para_vencer",
        "tipificacion",
        "descripcion",
    ]
    etiquetas = {
        "_horas_abierto": "horas_habiles_abierto",
        "_horas_para_vencer": "horas_habiles_para_vencer",
    }

    tab_vencidos, tab_proximos = st.tabs(["Vencidos", "Proximos 12h"])
    for tab, tabla in [
        (tab_vencidos, vencidos.sort_values(by="_horas_para_vencer")),
        (tab_proximos, proximos.sort_values(by="_horas_para_vencer")),
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
        trabajo["cliente_clave"] = pd.Series(dtype="object")
        trabajo["fuente_cliente"] = pd.Series(dtype="object")
        trabajo["creado_dt"] = pd.Series(dtype="datetime64[ns]")
        trabajo["cerrado_dt"] = pd.Series(dtype="datetime64[ns]")
        trabajo["tiempo_respuesta_h"] = pd.Series(dtype="float")
        return trabajo

    trabajo = normalizar_tipificaciones_casos_df(df)
    detecciones = trabajo.apply(
        lambda row: detectar_cliente_en_fila(row, ["cuenta"]),
        axis=1,
        result_type="expand",
    )
    detecciones.columns = ["cliente_clave", "fuente_cliente"]
    trabajo[["cliente_clave", "fuente_cliente"]] = detecciones
    trabajo = trabajo[trabajo["cliente_clave"] != ""].copy()
    trabajo["creado_dt"] = pd.to_datetime(trabajo["creado"].apply(normalizar_fecha), errors="coerce")
    trabajo["cerrado_dt"] = pd.to_datetime(trabajo["cerrado"].apply(normalizar_fecha), errors="coerce")
    trabajo["tiempo_respuesta_h"] = pd.to_numeric(trabajo["tiempo_respuesta"], errors="coerce")
    return trabajo


def preparar_incidentes_clientes_clave(df):
    if df.empty:
        trabajo = agregar_campos_sla_incidentes(df)
        trabajo["cliente_clave"] = pd.Series(dtype="object")
        trabajo["fuente_cliente"] = pd.Series(dtype="object")
        trabajo["creado_dt"] = pd.Series(dtype="datetime64[ns]")
        trabajo["cerrado_dt"] = pd.Series(dtype="datetime64[ns]")
        trabajo["duracion_horas_num"] = pd.Series(dtype="float")
        return trabajo

    trabajo = agregar_campos_sla_incidentes(df)
    campos = [
        "empresa",
        "solicitante",
        "breve_descripcion",
        "descripcion",
        "observaciones_trabajo",
        "observaciones_adicionales",
    ]
    detecciones = trabajo.apply(
        lambda row: detectar_cliente_en_fila(row, campos),
        axis=1,
        result_type="expand",
    )
    detecciones.columns = ["cliente_clave", "fuente_cliente"]
    trabajo[["cliente_clave", "fuente_cliente"]] = detecciones
    trabajo = trabajo[trabajo["cliente_clave"] != ""].copy()
    trabajo["creado_dt"] = pd.to_datetime(trabajo["creado"].apply(normalizar_fecha), errors="coerce")
    trabajo["cerrado_dt"] = pd.to_datetime(trabajo["cerrado"].apply(normalizar_fecha), errors="coerce")
    trabajo["duracion_horas_num"] = pd.to_numeric(trabajo["duracion_horas"], errors="coerce")
    return trabajo


def fecha_maxima_cliente(casos_cliente, incidentes_cliente):
    fechas = []
    if not casos_cliente.empty:
        fechas.append(casos_cliente["creado_dt"].max())
    if not incidentes_cliente.empty:
        fechas.append(incidentes_cliente["creado_dt"].max())
    fechas_validas = [fecha for fecha in fechas if pd.notna(fecha)]
    if not fechas_validas:
        return ""
    return max(fechas_validas).strftime("%Y-%m-%d %H:%M")


def valor_mas_frecuente(serie, default="Sin dato"):
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
        return "Amarillo", "Seguimiento", score
    return "Rojo", "Prioritario", score


def resumen_clientes_clave(casos, incidentes):
    filas = []
    for cliente in CLIENTES_CLAVE:
        casos_cliente = casos[casos["cliente_clave"] == cliente] if not casos.empty else pd.DataFrame()
        incidentes_cliente = (
            incidentes[incidentes["cliente_clave"] == cliente] if not incidentes.empty else pd.DataFrame()
        )

        total_casos = len(casos_cliente)
        total_incidentes = len(incidentes_cliente)
        if total_casos == 0 and total_incidentes == 0:
            filas.append(
                {
                    "Cliente": cliente,
                    "Nivel": "Sin actividad",
                    "Estado atencion": "Sin actividad",
                    "Score": 100,
                    "Casos": 0,
                    "Incidentes": 0,
                    "Total atenciones": 0,
                    "Abiertos": 0,
                    "SLA casos %": None,
                    "SLA incidentes %": None,
                    "Alertas incidentes": 0,
                    "Prioridad alta": 0,
                    "Casos sin causa": 0,
                    "Producto principal": "Sin dato",
                    "Causa incidente principal": "Sin dato",
                    "Ultima atencion": "",
                }
            )
            continue

        if total_casos:
            casos_cerrados = casos_cliente[mascara_cerrados(casos_cliente)]
            casos_abiertos = total_casos - len(casos_cerrados)
            tiempos_casos = casos_cerrados["tiempo_respuesta_h"].dropna()
            sla_casos = (
                porcentaje(len(tiempos_casos[tiempos_casos < SLA_CASOS_HORAS]), len(tiempos_casos))
                if len(tiempos_casos)
                else None
            )
            casos_sin_causa = len(
                casos_cliente[
                    casos_cliente["causa"].replace("", pd.NA).fillna("Sin dato").str.lower().isin(["sin dato"])
                ]
            )
        else:
            casos_abiertos = 0
            sla_casos = None
            casos_sin_causa = 0

        if total_incidentes:
            incidentes_cerrados_todos = incidentes_cliente[mascara_cerrados(incidentes_cliente)]
            incidentes_cerrados = incidentes_cerrados_todos[
                incidentes_cerrados_todos["aplica_sla_incidente"].fillna(False)
            ]
            incidentes_abiertos = total_incidentes - len(incidentes_cerrados_todos)
            sla_incidentes_base = incidentes_cerrados[
                incidentes_cerrados["estado_sla"].isin(["Cumple", "No cumple"])
            ]
            sla_incidentes = (
                porcentaje(len(sla_incidentes_base[sla_incidentes_base["estado_sla"] == "Cumple"]), len(sla_incidentes_base))
                if len(sla_incidentes_base)
                else None
            )
            alertas = len(
                incidentes_cliente[
                    (incidentes_cliente["es_alerta_auto"].fillna("No") == "Si")
                    | (incidentes_cliente["tipificacion_auto"].fillna("") == "Alerta NOC")
                ]
            )
            prioridad_alta = len(
                incidentes_cliente[
                    incidentes_cliente["prioridad"].fillna("").str.contains(
                        r"^(?:1|2\s*-\s*Alta)", case=False, regex=True
                    )
                ]
            )
        else:
            incidentes_abiertos = 0
            sla_incidentes = None
            alertas = 0
            prioridad_alta = 0
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

        filas.append(
            {
                "Cliente": cliente,
                "Nivel": nivel,
                "Estado atencion": estado_atencion,
                "Score": score,
                "Casos": total_casos,
                "Incidentes": total_incidentes,
                "Total atenciones": total_casos + total_incidentes,
                "Abiertos": abiertos,
                "SLA casos %": sla_casos,
                "SLA incidentes %": sla_incidentes,
                "Alertas incidentes": alertas,
                "Prioridad alta": prioridad_alta,
                "Casos sin causa": casos_sin_causa,
                "Producto principal": valor_mas_frecuente(casos_cliente.get("producto", pd.Series(dtype="object"))),
                "Causa incidente principal": valor_mas_frecuente(
                    incidentes_cliente.get("causa_raiz_auto", pd.Series(dtype="object"))
                ),
                "Ultima atencion": fecha_maxima_cliente(casos_cliente, incidentes_cliente),
            }
        )

    return pd.DataFrame(filas)


def tabla_atenciones_abiertas_clientes(casos, incidentes):
    tablas = []
    if not casos.empty:
        abiertos_casos = casos[~mascara_cerrados(casos)].copy()
        if not abiertos_casos.empty:
            abiertos_casos["Tipo"] = "Caso"
            abiertos_casos["Cliente"] = abiertos_casos["cliente_clave"]
            abiertos_casos["Numero"] = abiertos_casos["numero"]
            abiertos_casos["Estado"] = abiertos_casos["estado"]
            abiertos_casos["Prioridad"] = abiertos_casos["prioridad"]
            abiertos_casos["Responsable"] = abiertos_casos["asignado"]
            abiertos_casos["Creado"] = abiertos_casos["creado"]
            abiertos_casos["Vencimiento SLA"] = ""
            abiertos_casos["Clasificacion"] = abiertos_casos["tipificacion"]
            abiertos_casos["Resumen"] = abiertos_casos["descripcion"]
            tablas.append(
                abiertos_casos[
                    [
                        "Tipo",
                        "Cliente",
                        "Numero",
                        "Estado",
                        "Prioridad",
                        "Responsable",
                        "Creado",
                        "Vencimiento SLA",
                        "Clasificacion",
                        "Resumen",
                    ]
                ]
            )

    if not incidentes.empty:
        abiertos_incidentes = incidentes[~mascara_cerrados(incidentes)].copy()
        if not abiertos_incidentes.empty:
            abiertos_incidentes["Tipo"] = "Incidente"
            abiertos_incidentes["Cliente"] = abiertos_incidentes["cliente_clave"]
            abiertos_incidentes["Numero"] = abiertos_incidentes["numero"]
            abiertos_incidentes["Estado"] = abiertos_incidentes["estado"]
            abiertos_incidentes["Prioridad"] = abiertos_incidentes["prioridad"]
            abiertos_incidentes["Responsable"] = abiertos_incidentes["asignado_a"]
            abiertos_incidentes["Creado"] = abiertos_incidentes["creado"]
            abiertos_incidentes["Vencimiento SLA"] = abiertos_incidentes["fecha_vencimiento_sla"]
            abiertos_incidentes["Clasificacion"] = abiertos_incidentes["tipificacion_auto"]
            abiertos_incidentes["Resumen"] = abiertos_incidentes["breve_descripcion"].replace("", pd.NA).fillna(
                abiertos_incidentes["descripcion"]
            )
            tablas.append(
                abiertos_incidentes[
                    [
                        "Tipo",
                        "Cliente",
                        "Numero",
                        "Estado",
                        "Prioridad",
                        "Responsable",
                        "Creado",
                        "Vencimiento SLA",
                        "Clasificacion",
                        "Resumen",
                    ]
                ]
            )

    if not tablas:
        return pd.DataFrame(
            columns=[
                "Tipo",
                "Cliente",
                "Numero",
                "Estado",
                "Prioridad",
                "Responsable",
                "Creado",
                "Vencimiento SLA",
                "Clasificacion",
                "Resumen",
            ]
        )
    return pd.concat(tablas, ignore_index=True).sort_values(by=["Cliente", "Tipo", "Creado"])


def tabla_resumen_tipificaciones_casos(df):
    conteo_tipificaciones = df["tipificacion"].value_counts()
    resumen = pd.DataFrame(CASE_TIPIFICATION_GUIDE)
    resumen["Cantidad"] = resumen["Tipificacion"].map(conteo_tipificaciones).fillna(0).astype(int)

    tipificaciones_faltantes = [
        tipificacion
        for tipificacion in conteo_tipificaciones.index.tolist()
        if tipificacion not in resumen["Tipificacion"].tolist()
    ]
    if tipificaciones_faltantes:
        adicionales = pd.DataFrame(
            {
                "Tipificacion": tipificaciones_faltantes,
                "Descripcion": ["Tipificacion detectada sin descripcion configurada."] * len(tipificaciones_faltantes),
                "Cantidad": [int(conteo_tipificaciones[tip]) for tip in tipificaciones_faltantes],
            }
        )
        resumen = pd.concat([resumen, adicionales], ignore_index=True)

    return resumen.sort_values(by=["Cantidad", "Tipificacion"], ascending=[False, True]).reset_index(drop=True)


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
        "descripcion",
        "causa",
        "codigo_resolucion",
        "notas_resolucion",
        "observaciones_adicionales",
        "observaciones_trabajo",
        "producto",
    ]
    return " ".join(normalizar_texto(row.get(campo)) for campo in campos).strip()


def inferir_motivo_agendamiento(row):
    texto = texto_analisis_agendamiento(row)
    for motivo, palabras in AGENDA_REASON_RULES:
        if any(palabra in texto for palabra in palabras):
            return motivo
    return "Sin detalle suficiente"


def cliente_agendamiento(row):
    cuenta = valor_limpio(row.get("cuenta"))
    return cuenta if cuenta else "Sin cuenta"


def fecha_corta(valor):
    if pd.isna(valor):
        return ""
    return pd.Timestamp(valor).strftime("%Y-%m-%d")


def inicio_periodo_agendamiento(agenda, mes_dashboard):
    if mes_dashboard != "Todos":
        try:
            return pd.Period(mes_dashboard, freq="M").to_timestamp()
        except Exception:
            pass
    fechas = agenda.get("_creado_dt_dashboard", pd.Series(dtype="datetime64[ns]")).dropna()
    if fechas.empty:
        return None
    return fechas.min().normalize()


def preparar_analisis_agendamiento(df_periodo, df_historico, mes_dashboard):
    columnas_base = [
        "Canal agrupado",
        "Motivo inferido",
        "Cliente agenda",
        "Ciclo cliente",
        "Cliente recurrente agenda",
        "Casos historicos cliente",
        "Agendas historicas cliente",
    ]
    if df_periodo.empty:
        trabajo = df_periodo.copy()
        for columna in columnas_base:
            trabajo[columna] = pd.Series(dtype="object")
        return trabajo

    periodo = normalizar_tipificaciones_casos_df(df_periodo).copy()
    historico = normalizar_tipificaciones_casos_df(df_historico).copy()
    if "_creado_dt_dashboard" not in periodo.columns:
        periodo = preparar_fechas_dashboard(periodo)
    if "_creado_dt_dashboard" not in historico.columns:
        historico = preparar_fechas_dashboard(historico)

    agenda = periodo[periodo["tipificacion"].apply(es_tipificacion_agendamiento)].copy()
    if agenda.empty:
        for columna in columnas_base:
            agenda[columna] = pd.Series(dtype="object")
        return agenda

    for columna in ["canal", "cuenta", "numero"]:
        if columna not in agenda.columns:
            agenda[columna] = ""
        if columna not in historico.columns:
            historico[columna] = ""

    agenda["Canal agrupado"] = agenda["canal"].apply(canal_agrupado_agendamiento)
    agenda["Motivo inferido"] = agenda.apply(inferir_motivo_agendamiento, axis=1)
    agenda["Cliente agenda"] = agenda.apply(cliente_agendamiento, axis=1)
    agenda["_cliente_norm_agenda"] = agenda["cuenta"].apply(normalizar_texto)

    historico["_cliente_norm_agenda"] = historico["cuenta"].apply(normalizar_texto)
    historico_validos = historico[historico["_cliente_norm_agenda"] != ""].copy()
    if historico_validos.empty:
        agenda["Primera atencion cliente"] = pd.NaT
        agenda["Casos historicos cliente"] = 0
        agenda["Agendas historicas cliente"] = 0
    else:
        historia_cliente = historico_validos.groupby("_cliente_norm_agenda").agg(
            Primera_atencion_cliente=("_creado_dt_dashboard", "min"),
            Casos_historicos_cliente=("numero", "nunique"),
        )
        historico_agenda = historico_validos[
            historico_validos["tipificacion"].apply(es_tipificacion_agendamiento)
        ].copy()
        agendas_historicas = historico_agenda.groupby("_cliente_norm_agenda")["numero"].nunique()
        agenda["Primera atencion cliente"] = agenda["_cliente_norm_agenda"].map(
            historia_cliente["Primera_atencion_cliente"]
        )
        agenda["Casos historicos cliente"] = (
            agenda["_cliente_norm_agenda"]
            .map(historia_cliente["Casos_historicos_cliente"])
            .fillna(0)
            .astype(int)
        )
        agenda["Agendas historicas cliente"] = (
            agenda["_cliente_norm_agenda"].map(agendas_historicas).fillna(0).astype(int)
        )

    inicio_periodo = inicio_periodo_agendamiento(agenda, mes_dashboard)

    def clasificar_ciclo_cliente(row):
        if row["_cliente_norm_agenda"] == "":
            return "Sin cuenta"
        primera = row.get("Primera atencion cliente")
        if inicio_periodo is not None and pd.notna(primera) and primera < inicio_periodo:
            return "Cliente con historial previo"
        if row.get("Agendas historicas cliente", 0) > 1:
            return "Cliente recurrente en el periodo"
        return "Cliente nuevo en la base"

    agenda["Ciclo cliente"] = agenda.apply(clasificar_ciclo_cliente, axis=1)
    agenda["Cliente recurrente agenda"] = agenda["Agendas historicas cliente"].apply(
        lambda valor: "Si" if valor > 1 else "No"
    )
    return agenda


def lectura_ejecutiva_agendamiento(agenda, total_casos_periodo):
    columnas = ["Pregunta", "Lectura", "Evidencia", "Accion sugerida"]
    if agenda.empty:
        return pd.DataFrame(columns=columnas)

    total_agenda = len(agenda)
    mesa = int((agenda["Canal agrupado"] == "Mesa de ayuda").sum())
    directa = int((agenda["Canal agrupado"] == "Agenda directa").sum())
    sin_cuenta = int((agenda["Cliente agenda"] == "Sin cuenta").sum())
    clientes_identificados = agenda[agenda["_cliente_norm_agenda"] != ""]
    clientes_total = clientes_identificados["_cliente_norm_agenda"].nunique()
    clientes_previos = clientes_identificados[
        clientes_identificados["Ciclo cliente"] == "Cliente con historial previo"
    ]["_cliente_norm_agenda"].nunique()
    recurrentes = clientes_identificados[
        clientes_identificados["Cliente recurrente agenda"] == "Si"
    ]["_cliente_norm_agenda"].nunique()
    motivo_principal = valor_mas_frecuente(agenda["Motivo inferido"])
    canal_principal = valor_mas_frecuente(agenda["canal"].replace("", "Sin canal"))
    producto_principal = valor_mas_frecuente(agenda.get("producto", pd.Series(dtype="object")))

    return pd.DataFrame(
        [
            {
                "Pregunta": "Por que siguen entrando por mesa de ayuda?",
                "Lectura": (
                    "La mayoria no esta entrando por el calendario directo, sino por canales operativos "
                    "que terminan en mesa de ayuda."
                ),
                "Evidencia": (
                    f"{mesa} de {total_agenda} casos de agendamiento ({porcentaje(mesa, total_agenda)}%) "
                    f"entraron por mesa de ayuda. Canal principal: {canal_principal}. "
                    f"Agenda directa: {directa} casos."
                ),
                "Accion sugerida": (
                    "Redirigir Web/telefono/correo al enlace de agenda antes de crear caso y marcar excepciones "
                    "cuando mesa de ayuda deba tomarlo."
                ),
            },
            {
                "Pregunta": "Cual es la razon de tantos casos?",
                "Lectura": (
                    f"El motivo dominante es {motivo_principal}; esta tipologia concentra "
                    f"{porcentaje(total_agenda, total_casos_periodo)}% de los casos del periodo."
                ),
                "Evidencia": f"Producto principal asociado: {producto_principal}.",
                "Accion sugerida": (
                    "Separar en el cierre si fue instalacion, activacion/descarga, token fisico o firma; "
                    "asi la causa queda accionable y no solo como agendamiento."
                ),
            },
            {
                "Pregunta": "Son clientes viejos?",
                "Lectura": (
                    "La lectura debe hacerse sobre clientes con cuenta identificada; los casos sin cuenta "
                    "no permiten confirmar antiguedad."
                ),
                "Evidencia": (
                    f"{clientes_previos} de {clientes_total} clientes identificados tienen historial previo "
                    f"en la base. {recurrentes} clientes ya tienen mas de un caso de agendamiento historico."
                ),
                "Accion sugerida": (
                    "Revisar los clientes recurrentes y reforzar comunicacion del canal correcto de agenda "
                    "despues de cada cierre."
                ),
            },
            {
                "Pregunta": "Que puede distorsionar el analisis?",
                "Lectura": "La falta de cuenta limita saber si el caso viene de un cliente nuevo o antiguo.",
                "Evidencia": f"{sin_cuenta} casos ({porcentaje(sin_cuenta, total_agenda)}%) estan sin cuenta.",
                "Accion sugerida": "Hacer obligatorio el campo Cuenta o normalizarlo en la carga de casos.",
            },
        ],
        columns=columnas,
    )


def resumen_clientes_agendamiento(agenda):
    columnas = [
        "Cliente",
        "Casos agenda",
        "Canal principal",
        "Motivo principal",
        "Ciclo cliente",
        "Agendas historicas",
        "Casos historicos",
        "Primera agenda",
        "Ultima agenda",
    ]
    if agenda.empty:
        return pd.DataFrame(columns=columnas)

    resumen = (
        agenda.groupby("Cliente agenda", dropna=False)
        .agg(
            Casos_agenda=("numero", "count"),
            Canal_principal=("Canal agrupado", valor_mas_frecuente),
            Motivo_principal=("Motivo inferido", valor_mas_frecuente),
            Ciclo_cliente=("Ciclo cliente", valor_mas_frecuente),
            Agendas_historicas=("Agendas historicas cliente", "max"),
            Casos_historicos=("Casos historicos cliente", "max"),
            Primera_agenda=("_creado_dt_dashboard", "min"),
            Ultima_agenda=("_creado_dt_dashboard", "max"),
        )
        .reset_index()
        .rename(
            columns={
                "Cliente agenda": "Cliente",
                "Casos_agenda": "Casos agenda",
                "Canal_principal": "Canal principal",
                "Motivo_principal": "Motivo principal",
                "Ciclo_cliente": "Ciclo cliente",
                "Agendas_historicas": "Agendas historicas",
                "Casos_historicos": "Casos historicos",
                "Primera_agenda": "Primera agenda",
                "Ultima_agenda": "Ultima agenda",
            }
        )
    )
    resumen["Primera agenda"] = resumen["Primera agenda"].apply(fecha_corta)
    resumen["Ultima agenda"] = resumen["Ultima agenda"].apply(fecha_corta)
    return resumen.sort_values(by=["Casos agenda", "Cliente"], ascending=[False, True])[columnas]


def resumen_canales_agendamiento(agenda):
    columnas = ["Canal agrupado", "Canal", "Casos", "% agendamiento"]
    if agenda.empty:
        return pd.DataFrame(columns=columnas)

    resumen = (
        agenda.assign(Canal=agenda["canal"].replace("", "Sin canal"))
        .groupby(["Canal agrupado", "Canal"], dropna=False)
        .size()
        .reset_index(name="Casos")
        .sort_values(by=["Casos", "Canal"], ascending=[False, True])
    )
    total = len(agenda)
    resumen["% agendamiento"] = resumen["Casos"].apply(lambda valor: porcentaje(valor, total))
    return resumen[columnas]


def render_analisis_agendamiento_mesa(df_periodo, df_historico, mes_dashboard):
    agenda = preparar_analisis_agendamiento(df_periodo, df_historico, mes_dashboard)

    st.subheader("Analisis de agendamiento por mesa de ayuda")
    st.caption(
        "Cruce para explicar por que los casos de redireccionamiento a agenda siguen llegando por mesa de ayuda, "
        "que motivo operativo se repite y si hay clientes con historial previo."
    )

    if agenda.empty:
        st.info("No hay casos de redireccionamiento a agenda en el periodo seleccionado.")
        return

    total_agenda = len(agenda)
    mesa = int((agenda["Canal agrupado"] == "Mesa de ayuda").sum())
    directa = int((agenda["Canal agrupado"] == "Agenda directa").sum())
    clientes_identificados = agenda[agenda["_cliente_norm_agenda"] != ""]
    cuentas_identificadas = clientes_identificados["_cliente_norm_agenda"].nunique()
    cuentas_registradas = clientes_identificados[
        clientes_identificados["Ciclo cliente"] == "Cliente con historial previo"
    ]["_cliente_norm_agenda"].nunique()
    cuentas_nuevas = clientes_identificados[
        clientes_identificados["Ciclo cliente"] != "Cliente con historial previo"
    ]["_cliente_norm_agenda"].nunique()

    render_tarjetas(
        [
            ("Agendamiento", total_agenda),
            ("Mesa ayuda", f"{porcentaje(mesa, total_agenda)}%"),
            ("Agenda directa", directa),
            ("% Ctas registradas", f"{porcentaje(cuentas_registradas, cuentas_identificadas)}%"),
            ("% Ctas nuevas", f"{porcentaje(cuentas_nuevas, cuentas_identificadas)}%"),
        ]
    )

    tab_lectura, tab_motivos, tab_clientes, tab_detalle = st.tabs(
        ["Lectura", "Motivos", "Clientes", "Detalle"]
    )

    with tab_lectura:
        st.dataframe(
            lectura_ejecutiva_agendamiento(agenda, len(df_periodo)),
            use_container_width=True,
            hide_index=True,
        )

    with tab_motivos:
        motivos = (
            agenda["Motivo inferido"]
            .value_counts()
            .rename_axis("Motivo")
            .reset_index(name="Casos")
            .sort_values(by="Casos", ascending=True)
        )
        fig = px.bar(
            motivos,
            x="Casos",
            y="Motivo",
            orientation="h",
            text="Casos",
            color_discrete_sequence=[UI_PALETTE["yellow"]],
        )
        fig.update_traces(marker_color=UI_PALETTE["yellow"], textposition="outside")
        st.plotly_chart(aplicar_estilo_figura(fig, "Motivo inferido"), use_container_width=True)

    with tab_clientes:
        resumen_clientes = resumen_clientes_agendamiento(agenda)
        canales = resumen_canales_agendamiento(agenda)
        top_clientes = resumen_clientes.head(15).copy()
        if not top_clientes.empty:
            fig = px.bar(
                top_clientes.sort_values(by="Casos agenda", ascending=True),
                x="Casos agenda",
                y="Cliente",
                orientation="h",
                text="Casos agenda",
                color="Ciclo cliente",
                color_discrete_sequence=CHART_COLORS,
            )
            fig.update_traces(textposition="outside")
            st.plotly_chart(aplicar_estilo_figura(fig, "Clientes con mas agendamiento"), use_container_width=True)
        st.caption("Entrada por canal de los casos mostrados en este analisis.")
        st.dataframe(canales, use_container_width=True, hide_index=True)
        st.dataframe(resumen_clientes, use_container_width=True, hide_index=True)

    with tab_detalle:
        columnas = [
            "numero",
            "cuenta",
            "canal",
            "Canal agrupado",
            "Motivo inferido",
            "Ciclo cliente",
            "Cliente recurrente agenda",
            "Agendas historicas cliente",
            "Casos historicos cliente",
            "producto",
            "asignado",
            "creado",
            "estado",
            "descripcion",
            "causa",
        ]
        visibles = [col for col in columnas if col in agenda.columns]
        st.dataframe(
            agenda.sort_values(by="_creado_dt_dashboard", ascending=False)[visibles],
            use_container_width=True,
            hide_index=True,
        )


def tabla_resumen_tipologias_incidentes(df):
    if df.empty:
        return pd.DataFrame(
            columns=[
                "Tipologia",
                "Lectura ejecutiva",
                "Total",
                "Cerrados",
                "Abiertos",
                "SLA aplica",
                "Cumple SLA",
                "No cumple SLA",
                "SLA %",
                "Prom. horas",
                "Prom. dias",
            ]
        )

    trabajo = df.copy()
    trabajo["tipificacion_auto"] = trabajo["tipificacion_auto"].fillna("Incidente Interno")
    trabajo["_cerrado"] = mascara_cerrados(trabajo)
    trabajo["_duracion_horas_num"] = pd.to_numeric(trabajo["duracion_horas"], errors="coerce")
    tipologias = [
        tipologia
        for tipologia in INCIDENT_TIPIFICATION_ORDER
        if tipologia in set(trabajo["tipificacion_auto"].dropna())
    ]
    tipologias.extend(
        sorted(
            tipologia
            for tipologia in trabajo["tipificacion_auto"].dropna().unique().tolist()
            if tipologia not in tipologias
        )
    )

    filas = []
    for tipologia in tipologias:
        grupo = trabajo[trabajo["tipificacion_auto"] == tipologia].copy()
        cerrados = grupo[grupo["_cerrado"]].copy()
        sla_aplica = bool(grupo.get("aplica_sla_incidente", pd.Series(dtype="bool")).fillna(False).any())
        if sla_aplica:
            con_sla = cerrados[
                cerrados.get("aplica_sla_incidente", pd.Series(False, index=cerrados.index)).fillna(False)
                & cerrados["estado_sla"].isin(["Cumple", "No cumple"])
            ].copy()
        else:
            con_sla = cerrados.iloc[0:0].copy()
        cumple = int((con_sla["estado_sla"] == "Cumple").sum())
        no_cumple = int((con_sla["estado_sla"] == "No cumple").sum())
        total_sla = cumple + no_cumple
        promedio_horas = cerrados["_duracion_horas_num"].dropna().mean()
        sla_porcentaje = f"{porcentaje(cumple, total_sla)}%" if total_sla else (
            "Sin cierre" if sla_aplica else "No aplica"
        )
        filas.append(
            {
                "Tipologia": tipologia,
                "Lectura ejecutiva": INCIDENT_TIPIFICATION_GUIDE.get(
                    tipologia,
                    "Tipologia detectada en el archivo cargado.",
                ),
                "Total": len(grupo),
                "Cerrados": len(cerrados),
                "Abiertos": len(grupo) - len(cerrados),
                "SLA aplica": "Si" if sla_aplica else "No aplica",
                "Cumple SLA": cumple if sla_aplica else pd.NA,
                "No cumple SLA": no_cumple if sla_aplica else pd.NA,
                "SLA %": sla_porcentaje,
                "Prom. horas": round(promedio_horas, 2) if pd.notna(promedio_horas) else pd.NA,
                "Prom. dias": round(promedio_horas / 24, 2) if pd.notna(promedio_horas) else pd.NA,
            }
        )

    resumen = pd.DataFrame(filas)
    for columna in ["Total", "Cerrados", "Abiertos", "Cumple SLA", "No cumple SLA"]:
        resumen[columna] = resumen[columna].astype("Int64")
    for columna in ["Prom. horas", "Prom. dias"]:
        resumen[columna] = resumen[columna].astype("Float64")
    return resumen


def clasificar_causa_cliente_externo(causa):
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


def resumen_causas_cliente_externo(df):
    columnas = [
        "Causa raiz",
        "Cantidad",
        "% cliente externo",
        "Lectura ejecutiva",
        "Accion sugerida",
        "Detalle tecnico observado",
    ]
    if df.empty:
        return pd.DataFrame(columns=columnas)

    trabajo = df.copy()
    trabajo["causa_tecnica"] = trabajo["causa_raiz_auto"].replace("", pd.NA).fillna("Sin inferencia")
    clasificacion = trabajo["causa_tecnica"].apply(clasificar_causa_cliente_externo)
    trabajo[["Causa raiz", "Lectura ejecutiva", "Accion sugerida"]] = pd.DataFrame(
        clasificacion.tolist(),
        index=trabajo.index,
    )

    total = len(trabajo)

    def detalles_tecnicos(serie):
        valores = serie.dropna().astype(str).value_counts().head(2).index.tolist()
        return "; ".join(valores)

    resumen = (
        trabajo.groupby(["Causa raiz", "Lectura ejecutiva", "Accion sugerida"], dropna=False)
        .agg(
            Cantidad=("numero", "count"),
            Detalle_tecnico_observado=("causa_tecnica", detalles_tecnicos),
        )
        .reset_index()
        .sort_values(by=["Cantidad", "Causa raiz"], ascending=[False, True])
    )
    resumen["% cliente externo"] = resumen["Cantidad"].apply(lambda valor: porcentaje(valor, total))
    resumen = resumen.rename(columns={"Detalle_tecnico_observado": "Detalle tecnico observado"})
    return resumen[columnas]


def causas_relevantes_cliente_externo(causas):
    if causas.empty:
        return causas
    excluidas = [
        "Casos repetidos o registros duplicados que distorsionan la lectura operativa.",
        "Hallazgos tecnicos con bajo volumen o descripcion no estandarizada.",
        "No hay informacion suficiente para explicar la causa.",
    ]
    relevantes = causas[~causas["Lectura ejecutiva"].isin(excluidas)].copy()
    return relevantes[relevantes["Cantidad"] >= 2].copy()


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
    tiempos_cerrados = pd.to_numeric(df_cerrados["tiempo_respuesta"], errors="coerce").dropna()
    promedio = round(tiempos_cerrados.mean(), 2) if len(tiempos_cerrados) > 0 else 0

    total_cerrados = len(tiempos_cerrados)
    cumplen = len(tiempos_cerrados[tiempos_cerrados < SLA_CASOS_HORAS])
    porcentaje_sla = round((cumplen / total_cerrados) * 100, 2) if total_cerrados > 0 else 0
    incumplen = total_cerrados - cumplen

    render_tarjetas(
        [
            ("Total Casos", total),
            ("Cerrados", cerrados),
            ("Abiertos", abiertos),
            ("Promedio (h)", promedio),
            (f"SLA <{SLA_CASOS_HORAS}h (%)", f"{porcentaje_sla}%"),
        ]
    )
    st.caption(f"Periodo: {mes_dashboard} | Cumplen: {cumplen} | No cumplen: {incumplen}")

    st.divider()
    col1, col2 = st.columns(2)

    with col1:
        tip = tabla_resumen_tipificaciones_casos(df)[["Tipificacion", "Cantidad"]]
        tip = tip.sort_values(by="Cantidad", ascending=True)
        fig = px.bar(
            tip,
            x="Cantidad",
            y="Tipificacion",
            orientation="h",
            text="Cantidad",
            color="Tipificacion",
            color_discrete_sequence=CHART_COLORS,
        )
        fig.update_traces(textposition="outside")
        fig.update_layout(showlegend=False)
        st.plotly_chart(aplicar_estilo_figura(fig, "Casos por tipificacion"), use_container_width=True)

    with col2:
        serie = df.copy()
        casos_dia = serie.groupby(serie["_creado_dt_dashboard"].dt.date).size().reset_index(name="casos")
        casos_dia.columns = ["Fecha", "casos"]
        fig = px.bar(casos_dia, x="Fecha", y="casos", color_discrete_sequence=[UI_PALETTE["yellow"]])
        fig.update_traces(marker_color=UI_PALETTE["yellow"])
        st.plotly_chart(aplicar_estilo_figura(fig, "Casos por dia"), use_container_width=True)

    st.divider()
    render_analisis_agendamiento_mesa(df, df_historico, mes_dashboard)

    st.divider()
    st.subheader("Resumen de tipificaciones")
    st.caption("Descripcion breve de cada tipificacion y cantidad actual de casos clasificados en el dashboard.")
    st.dataframe(tabla_resumen_tipificaciones_casos(df), use_container_width=True, hide_index=True)

    render_seguimiento_casos(df)


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

    incidentes_reales = df[df["aplica_sla_incidente"]].copy()
    alertas_noc = df[df["tipificacion_auto"].fillna("") == "Alerta NOC"].copy()
    consultas_noc = df[df["tipificacion_auto"].fillna("") == "Consulta NOC"].copy()
    incidentes_externos = df[df["tipificacion_auto"].fillna("") == "Incidente Cliente Externo"].copy()
    incidentes_internos = df[df["tipificacion_auto"].fillna("") == "Incidente Interno"].copy()
    incidentes_seguridad = df[df["tipificacion_auto"].fillna("") == "Incidente Seguridad"].copy()

    duraciones = pd.to_numeric(incidentes_reales["duracion_horas"], errors="coerce").dropna()
    promedio = round(duraciones.mean(), 2) if len(duraciones) > 0 else 0

    df_cerrados = incidentes_reales[mascara_cerrados(incidentes_reales)].copy()
    df_cerrados = df_cerrados[df_cerrados["estado_sla"].isin(["Cumple", "No cumple"])]
    total_cerrados = len(df_cerrados)
    cumplen = len(df_cerrados[df_cerrados["estado_sla"] == "Cumple"])
    porcentaje_sla = round((cumplen / total_cerrados) * 100, 2) if total_cerrados > 0 else 0
    incumplen = total_cerrados - cumplen
    sla_base = df_cerrados.copy()
    if not sla_base.empty:
        sla_base["duracion_horas_num"] = pd.to_numeric(sla_base["duracion_horas"], errors="coerce")
        sla_base["duracion_dias_num"] = (sla_base["duracion_horas_num"] / 24).round(2)
        sla_base["sla_objetivo_dias"] = (pd.to_numeric(sla_base["sla_objetivo_horas"], errors="coerce") / 24).round(2)
        sla_base["SLA objetivo"] = sla_base["sla_objetivo_horas"].apply(formato_horas_dias)
        sla_base["Rango resolucion"] = sla_base["duracion_horas_num"].apply(clasificar_rango_resolucion)
    alertas_tipificadas = len(
        df[
            (df["es_alerta_auto"].fillna("No") == "Si")
            | (df["tipificacion_auto"].fillna("") == "Alerta NOC")
        ]
    )

    render_tarjetas(
        [
            ("Atenciones", total),
            ("Incidentes", len(incidentes_reales)),
            ("Alertas NOC", len(alertas_noc)),
            ("Consultas NOC", len(consultas_noc)),
            ("SLA (%)", f"{porcentaje_sla}%"),
        ]
    )
    st.caption(
        f"Periodo: {mes_dashboard} | Cerrados: {cerrados} | Abiertos: {abiertos} | "
        f"Promedio incidentes: {promedio}h | Cumplen: {cumplen} | No cumplen: {incumplen} | "
        f"Alertas tipificadas: {alertas_tipificadas} | Externos: {len(incidentes_externos)} | "
        f"Internos: {len(incidentes_internos)} | Seguridad: {len(incidentes_seguridad)}"
    )

    st.divider()
    st.subheader("tipologia Incidentes")
    st.caption(
        "Vista consolidada para entender que tipo de atenciones entraron al periodo, cuantas siguen abiertas "
        "y como va el cumplimiento de SLA donde aplica."
    )
    st.dataframe(
        tabla_resumen_tipologias_incidentes(df),
        use_container_width=True,
        hide_index=True,
    )

    st.divider()
    fila1_col1, fila1_col2 = st.columns(2)

    with fila1_col1:
        tip = df["tipificacion_auto"].fillna("Incidente Interno").value_counts().reset_index()
        tip.columns = ["Tipologia", "Cantidad"]
        tip = tip.sort_values(by="Cantidad", ascending=True)
        fig = px.bar(
            tip,
            x="Cantidad",
            y="Tipologia",
            orientation="h",
            text="Cantidad",
            color="Tipologia",
            color_discrete_sequence=CHART_COLORS,
        )
        fig.update_traces(textposition="outside")
        fig.update_layout(showlegend=False, title="Atenciones por tipologia")
        st.plotly_chart(aplicar_estilo_figura(fig, "Atenciones por tipologia"), use_container_width=True)

    with fila1_col2:
        if sla_base.empty:
            st.info("No hay incidentes cerrados con objetivo SLA para graficar.")
        else:
            cumplimiento_prioridad = (
                sla_base.groupby(["prioridad_normalizada", "estado_sla"])
                .size()
                .reset_index(name="Cantidad")
            )
            fig = px.bar(
                cumplimiento_prioridad,
                x="Cantidad",
                y="prioridad_normalizada",
                orientation="h",
                text="Cantidad",
                color="estado_sla",
                barmode="stack",
                category_orders={"prioridad_normalizada": ["Bajo", "Moderado", "Alto", "Critico"]},
                color_discrete_map={
                    "Cumple": UI_PALETTE["lavender"],
                    "No cumple": UI_PALETTE["primary"],
                },
                labels={
                    "prioridad_normalizada": "Prioridad",
                    "estado_sla": "Estado SLA",
                },
            )
            fig.update_traces(textposition="inside")
            st.plotly_chart(aplicar_estilo_figura(fig, "Cumplimiento SLA por prioridad"), use_container_width=True)

    fila2_col1, fila2_col2 = st.columns(2)

    with fila2_col1:
        servicios = df["servicio_negocio"].fillna("Sin servicio").value_counts().reset_index()
        servicios.columns = ["Servicio", "Cantidad"]
        servicios = servicios.sort_values(by="Cantidad", ascending=True)
        fig = px.bar(
            servicios,
            x="Cantidad",
            y="Servicio",
            orientation="h",
            text="Cantidad",
            color_discrete_sequence=[UI_PALETTE["yellow"]],
        )
        fig.update_traces(textposition="outside")
        fig.update_layout(title="Servicios afectados")
        st.plotly_chart(aplicar_estilo_figura(fig, "Servicios afectados"), use_container_width=True)

    with fila2_col2:
        serie = df.copy()
        incidentes_dia = serie.groupby(serie["_creado_dt_dashboard"].dt.date).size().reset_index(name="incidentes")
        incidentes_dia.columns = ["Fecha", "incidentes"]
        fig = px.bar(
            incidentes_dia,
            x="Fecha",
            y="incidentes",
            title="Atenciones por dia",
            color_discrete_sequence=[UI_PALETTE["purple"]],
        )
        fig.update_traces(marker_color=UI_PALETTE["purple"])
        st.plotly_chart(aplicar_estilo_figura(fig, "Atenciones por dia"), use_container_width=True)

    st.divider()
    st.subheader("SLA por matriz")
    if sla_base.empty:
        st.info("No hay incidentes cerrados con objetivo SLA para el periodo seleccionado.")
    else:
        st.caption(
            "El objetivo se muestra en horas y dias. El cumplimiento compara la duracion cargada del incidente "
            "contra el SLA de su tipificacion y prioridad."
        )

        rango_resolucion = (
            sla_base["Rango resolucion"]
            .value_counts()
            .reindex(RANGO_RESOLUCION_ORDEN, fill_value=0)
            .reset_index()
        )
        rango_resolucion.columns = ["Rango resolucion", "Cantidad"]
        rango_resolucion = rango_resolucion[rango_resolucion["Cantidad"] > 0]

        rango_col1, rango_col2 = st.columns(2)
        with rango_col1:
            fig = px.bar(
                rango_resolucion,
                x="Rango resolucion",
                y="Cantidad",
                text="Cantidad",
                color_discrete_sequence=[UI_PALETTE["yellow"]],
            )
            fig.update_traces(marker_color=UI_PALETTE["yellow"], textposition="outside")
            st.plotly_chart(aplicar_estilo_figura(fig, "Rangos de resolucion"), use_container_width=True)

        with rango_col2:
            objetivo_resumen = (
                sla_base.groupby(["SLA objetivo", "estado_sla"])
                .size()
                .reset_index(name="Cantidad")
            )
            fig = px.bar(
                objetivo_resumen,
                x="Cantidad",
                y="SLA objetivo",
                orientation="h",
                text="Cantidad",
                color="estado_sla",
                barmode="stack",
                color_discrete_map={
                    "Cumple": UI_PALETTE["lavender"],
                    "No cumple": UI_PALETTE["primary"],
                },
                labels={"estado_sla": "Estado SLA"},
            )
            fig.update_traces(textposition="inside")
            st.plotly_chart(aplicar_estilo_figura(fig, "Cumplimiento por objetivo SLA"), use_container_width=True)

        sla_resumen = (
            sla_base.groupby(
                ["tipificacion_auto", "prioridad_normalizada", "sla_objetivo_horas", "SLA objetivo"],
                dropna=False,
            )
            .agg(
                Total=("numero", "count"),
                Cumple=("estado_sla", lambda serie: int((serie == "Cumple").sum())),
                No_cumple=("estado_sla", lambda serie: int((serie == "No cumple").sum())),
                Prom_horas=("duracion_horas_num", "mean"),
                Prom_dias=("duracion_dias_num", "mean"),
                Max_horas=("duracion_horas_num", "max"),
                Max_dias=("duracion_dias_num", "max"),
            )
            .reset_index()
        )
        sla_resumen["Cumplimiento %"] = sla_resumen.apply(
            lambda row: porcentaje(row["Cumple"], row["Total"]),
            axis=1,
        )
        sla_resumen = sla_resumen.rename(
            columns={
                "tipificacion_auto": "Tipificacion",
                "prioridad_normalizada": "Prioridad",
                "sla_objetivo_horas": "SLA objetivo h",
                "No_cumple": "No cumple",
                "Prom_horas": "Prom. horas",
                "Prom_dias": "Prom. dias",
                "Max_horas": "Max. horas",
                "Max_dias": "Max. dias",
            }
        )
        columnas_numericas = ["Prom. horas", "Prom. dias", "Max. horas", "Max. dias", "SLA objetivo h"]
        for columna in columnas_numericas:
            sla_resumen[columna] = pd.to_numeric(sla_resumen[columna], errors="coerce").round(2)
        st.dataframe(
            sla_resumen[
                [
                    "Tipificacion",
                    "Prioridad",
                    "SLA objetivo",
                    "SLA objetivo h",
                    "Prom. horas",
                    "Prom. dias",
                    "Max. horas",
                    "Max. dias",
                    "Cumple",
                    "No cumple",
                    "Total",
                    "Cumplimiento %",
                ]
            ],
            use_container_width=True,
            hide_index=True,
        )

    st.divider()
    st.subheader("Cliente externo")

    df_cliente_externo = df[df["tipificacion_auto"].fillna("Incidente Interno") == "Incidente Cliente Externo"].copy()
    causas_cliente = resumen_causas_cliente_externo(df_cliente_externo)
    causas_relevantes = causas_relevantes_cliente_externo(causas_cliente)

    resumen_col1, resumen_col2 = st.columns(2)
    porcentaje_externo = round((len(df_cliente_externo) / total) * 100, 2) if total > 0 else 0

    with resumen_col1:
        st.markdown(tarjeta("Incidentes cliente externo", len(df_cliente_externo)), unsafe_allow_html=True)
        st.caption(f"Del total de atenciones del periodo: {porcentaje_externo}%")

    with resumen_col2:
        st.markdown(tarjeta("Causas raiz", len(causas_relevantes)), unsafe_allow_html=True)
        if not causas_relevantes.empty:
            causa_principal = causas_relevantes.iloc[0]
            st.caption(f"Mas frecuente: {causa_principal['Lectura ejecutiva']} ({causa_principal['Cantidad']})")
        else:
            st.caption("Sin causas raiz relevantes en el periodo.")

    if not causas_relevantes.empty:
        st.caption(
            "Base de la grafica: solo incidentes tipificados como cliente externo. Se excluyen duplicados, "
            "hallazgos sin causa clara, alertas, consultas y atenciones a reclasificar."
        )
        grafico_causas = (
            causas_relevantes.groupby("Lectura ejecutiva", as_index=False)
            .agg(Cantidad=("Cantidad", "sum"))
            .sort_values(by="Cantidad", ascending=True)
        )

        fig = px.bar(
            grafico_causas,
            x="Cantidad",
            y="Lectura ejecutiva",
            orientation="h",
            text="Cantidad",
            color_discrete_sequence=[UI_PALETTE["purple"]],
            title="Causas raiz en cliente externo",
        )
        fig.update_traces(textposition="outside")
        fig = aplicar_estilo_figura(fig, "Causas raiz en cliente externo")
        fig.update_layout(
            height=max(360, 58 * len(grafico_causas)),
            yaxis=dict(automargin=True),
        )
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(causas_relevantes, use_container_width=True, hide_index=True)
    elif not df_cliente_externo.empty:
        st.info("Hay incidentes de cliente externo, pero no hay causas raiz relevantes para graficar en el periodo.")
    else:
        st.info("No hay incidentes tipificados como cliente externo en los datos cargados.")


def dashboard_clientes_clave():
    casos = preparar_casos_clientes_clave(load_casos())
    incidentes = preparar_incidentes_clientes_clave(load_incidentes())

    st.subheader("Clientes clave")

    fechas = []
    if not casos.empty:
        fechas.append(casos["creado_dt"])
    if not incidentes.empty:
        fechas.append(incidentes["creado_dt"])
    fechas = pd.concat(fechas).dropna() if fechas else pd.Series(dtype="datetime64[ns]")

    filtro_col1, filtro_col2, filtro_col3 = st.columns([2, 1, 1])
    with filtro_col1:
        clientes_seleccionados = st.multiselect(
            "Clientes",
            CLIENTES_CLAVE,
            default=CLIENTES_CLAVE,
            key="clientes_clave_filtro",
        )
    with filtro_col2:
        base_meses = pd.DataFrame({"creado_dt": fechas})
        mes_dashboard = selector_mes_dashboard(base_meses, "clientes_clave_mes", "creado_dt")
    with filtro_col3:
        rango_fechas = None
        if not fechas.empty:
            fecha_min = fechas.min().date()
            fecha_max = fechas.max().date()
            rango_fechas = st.date_input(
                "Rango de fechas",
                value=(fecha_min, fecha_max),
                min_value=fecha_min,
                max_value=fecha_max,
                key="clientes_clave_rango",
            )

    if not clientes_seleccionados:
        st.warning("Selecciona al menos un cliente clave.")
        return

    casos = casos[casos["cliente_clave"].isin(clientes_seleccionados)].copy()
    incidentes = incidentes[incidentes["cliente_clave"].isin(clientes_seleccionados)].copy()

    if mes_dashboard != "Todos":
        if not casos.empty:
            casos = filtrar_mes_dashboard(casos, mes_dashboard, "creado_dt")
        if not incidentes.empty:
            incidentes = filtrar_mes_dashboard(incidentes, mes_dashboard, "creado_dt")

    if rango_fechas and isinstance(rango_fechas, tuple) and len(rango_fechas) == 2:
        fecha_inicio = pd.Timestamp(rango_fechas[0])
        fecha_fin = pd.Timestamp(rango_fechas[1]) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
        if not casos.empty:
            casos = casos[casos["creado_dt"].between(fecha_inicio, fecha_fin, inclusive="both")].copy()
        if not incidentes.empty:
            incidentes = incidentes[incidentes["creado_dt"].between(fecha_inicio, fecha_fin, inclusive="both")].copy()

    resumen = resumen_clientes_clave(casos, incidentes)
    resumen = resumen[resumen["Cliente"].isin(clientes_seleccionados)].copy()
    resumen_actividad = resumen[resumen["Total atenciones"] > 0].copy()

    total_casos = len(casos)
    total_incidentes = len(incidentes)
    abiertos_casos = len(casos[~mascara_cerrados(casos)]) if not casos.empty else 0
    abiertos_incidentes = len(incidentes[~mascara_cerrados(incidentes)]) if not incidentes.empty else 0
    clientes_activos = len(resumen_actividad)
    clientes_seguimiento = len(
        resumen_actividad[
            resumen_actividad["Nivel"].isin(["Amarillo", "Rojo"])
            & (resumen_actividad["Abiertos"] > 0)
        ]
    )

    casos_cerrados = casos[mascara_cerrados(casos)] if not casos.empty else pd.DataFrame()
    tiempos_casos = casos_cerrados.get("tiempo_respuesta_h", pd.Series(dtype="float")).dropna()
    sla_casos = (
        porcentaje(len(tiempos_casos[tiempos_casos < SLA_CASOS_HORAS]), len(tiempos_casos))
        if len(tiempos_casos)
        else 0
    )

    incidentes_cerrados = incidentes[mascara_cerrados(incidentes)] if not incidentes.empty else pd.DataFrame()
    if not incidentes_cerrados.empty and "aplica_sla_incidente" in incidentes_cerrados.columns:
        incidentes_cerrados = incidentes_cerrados[incidentes_cerrados["aplica_sla_incidente"].fillna(False)]
    sla_incidentes_base = (
        incidentes_cerrados[incidentes_cerrados["estado_sla"].isin(["Cumple", "No cumple"])]
        if not incidentes_cerrados.empty and "estado_sla" in incidentes_cerrados.columns
        else pd.DataFrame()
    )
    sla_incidentes = (
        porcentaje(
            len(sla_incidentes_base[sla_incidentes_base["estado_sla"] == "Cumple"]),
            len(sla_incidentes_base),
        )
        if len(sla_incidentes_base)
        else 0
    )

    render_tarjetas(
        [
            ("Clientes activos", clientes_activos),
            ("Atenciones", total_casos + total_incidentes),
            ("Abiertos", abiertos_casos + abiertos_incidentes),
            (f"SLA casos <{SLA_CASOS_HORAS}h", f"{sla_casos}%"),
            ("SLA incidentes", f"{sla_incidentes}%"),
        ]
    )
    st.caption(
        f"Periodo: {mes_dashboard} | Casos: {total_casos} | Incidentes: {total_incidentes} | "
        f"Clientes en seguimiento: {clientes_seguimiento}"
    )

    clientes_sin_actividad = resumen[resumen["Total atenciones"] == 0]["Cliente"].tolist()
    if clientes_sin_actividad:
        st.caption("Sin actividad en el periodo: " + ", ".join(clientes_sin_actividad))

    if resumen_actividad.empty:
        st.info("No hay casos o incidentes asociados a los clientes seleccionados en el periodo.")
        return

    st.divider()
    graf_col1, graf_col2 = st.columns(2)

    with graf_col1:
        grafico = resumen_actividad.sort_values(by="Total atenciones", ascending=True)
        fig = px.bar(
            grafico,
            x="Total atenciones",
            y="Cliente",
            orientation="h",
            text="Total atenciones",
            color="Nivel",
            color_discrete_map={
                "Verde": UI_PALETTE["lavender"],
                "Amarillo": UI_PALETTE["yellow"],
                "Rojo": UI_PALETTE["primary"],
            },
        )
        fig.update_traces(textposition="outside")
        st.plotly_chart(aplicar_estilo_figura(fig, "Atenciones por cliente clave"), use_container_width=True)

    with graf_col2:
        actividad = []
        if not casos.empty:
            casos_dia = casos[["creado_dt"]].dropna().copy()
            casos_dia["Fecha"] = casos_dia["creado_dt"].dt.date
            casos_dia["Tipo"] = "Casos"
            actividad.append(casos_dia.groupby(["Fecha", "Tipo"]).size().reset_index(name="Cantidad"))
        if not incidentes.empty:
            incidentes_dia = incidentes[["creado_dt"]].dropna().copy()
            incidentes_dia["Fecha"] = incidentes_dia["creado_dt"].dt.date
            incidentes_dia["Tipo"] = "Incidentes"
            actividad.append(incidentes_dia.groupby(["Fecha", "Tipo"]).size().reset_index(name="Cantidad"))
        if actividad:
            actividad_dia = pd.concat(actividad, ignore_index=True)
            fig = px.bar(
                actividad_dia,
                x="Fecha",
                y="Cantidad",
                color="Tipo",
                barmode="group",
                color_discrete_sequence=[UI_PALETTE["primary"], UI_PALETTE["purple"]],
            )
            st.plotly_chart(aplicar_estilo_figura(fig, "Actividad por dia"), use_container_width=True)
        else:
            st.info("No hay fechas validas para graficar actividad.")

    graf_col3, graf_col4 = st.columns(2)
    with graf_col3:
        if not casos.empty:
            productos = casos["producto"].replace("", pd.NA).fillna("Sin producto").value_counts().reset_index()
            productos.columns = ["Producto", "Cantidad"]
            productos = productos.head(10).sort_values(by="Cantidad", ascending=True)
            fig = px.bar(
                productos,
                x="Cantidad",
                y="Producto",
                orientation="h",
                text="Cantidad",
                color_discrete_sequence=[UI_PALETTE["yellow"]],
            )
            fig.update_traces(textposition="outside")
            st.plotly_chart(aplicar_estilo_figura(fig, "Productos con mas casos"), use_container_width=True)
        else:
            st.info("No hay casos asociados a clientes clave.")

    with graf_col4:
        if not incidentes.empty:
            causas = (
                incidentes["causa_raiz_auto"]
                .replace("", pd.NA)
                .fillna("Sin inferencia")
                .value_counts()
                .reset_index()
            )
            causas.columns = ["Causa incidente", "Cantidad"]
            causas = causas.head(10).sort_values(by="Cantidad", ascending=True)
            fig = px.bar(
                causas,
                x="Cantidad",
                y="Causa incidente",
                orientation="h",
                text="Cantidad",
                color_discrete_sequence=[UI_PALETTE["purple"]],
            )
            fig.update_traces(textposition="outside")
            st.plotly_chart(aplicar_estilo_figura(fig, "Causas en incidentes"), use_container_width=True)
        else:
            st.info("No hay incidentes asociados a clientes clave.")

    st.divider()
    tab_resumen, tab_abiertos, tab_casos, tab_incidentes, tab_seguimiento = st.tabs(
        ["Resumen", "Abiertos", "Casos", "Incidentes", "Seguimiento"]
    )

    with tab_resumen:
        columnas_resumen = [
            "Cliente",
            "Nivel",
            "Estado atencion",
            "Score",
            "Casos",
            "Incidentes",
            "Total atenciones",
            "Abiertos",
            "SLA casos %",
            "SLA incidentes %",
            "Alertas incidentes",
            "Casos sin causa",
            "Producto principal",
            "Causa incidente principal",
            "Ultima atencion",
        ]
        st.dataframe(
            resumen[columnas_resumen].sort_values(by=["Total atenciones", "Score"], ascending=[False, True]),
            use_container_width=True,
            hide_index=True,
        )

    with tab_abiertos:
        abiertos_detalle = tabla_atenciones_abiertas_clientes(casos, incidentes)
        if abiertos_detalle.empty:
            st.success("No hay casos ni incidentes abiertos para los clientes seleccionados.")
        else:
            st.caption("Detalle de atenciones abiertas. El campo Tipo indica si corresponde a caso o incidente.")
            st.dataframe(abiertos_detalle, use_container_width=True, hide_index=True)

    with tab_casos:
        if casos.empty:
            st.info("No hay casos asociados a los clientes seleccionados.")
        else:
            columnas_casos = [
                "cliente_clave",
                "numero",
                "cuenta",
                "estado",
                "prioridad",
                "producto",
                "tipificacion",
                "tiempo_respuesta",
                "creado",
                "cerrado",
                "causa",
                "codigo_resolucion",
                "fuente_cliente",
            ]
            columnas_casos = [col for col in columnas_casos if col in casos.columns]
            st.dataframe(
                casos.sort_values(by="creado_dt", ascending=False)[columnas_casos],
                use_container_width=True,
                hide_index=True,
            )

    with tab_incidentes:
        if incidentes.empty:
            st.info("No hay incidentes asociados a los clientes seleccionados.")
        else:
            columnas_incidentes = [
                "cliente_clave",
                "numero",
                "empresa",
                "solicitante",
                "estado",
                "prioridad",
                "servicio_negocio",
                "tipificacion_auto",
                "es_alerta_auto",
                "causa_raiz_auto",
                "duracion_horas",
                "creado",
                "cerrado",
                "breve_descripcion",
                "fuente_cliente",
            ]
            columnas_incidentes = [col for col in columnas_incidentes if col in incidentes.columns]
            st.dataframe(
                incidentes.sort_values(by="creado_dt", ascending=False)[columnas_incidentes],
                use_container_width=True,
                hide_index=True,
            )

    with tab_seguimiento:
        seguimiento = resumen_actividad[
            resumen_actividad["Nivel"].isin(["Amarillo", "Rojo"])
            & (resumen_actividad["Abiertos"] > 0)
        ].copy()
        if seguimiento.empty:
            st.success("Los clientes clave con actividad estan en nivel estable para el periodo seleccionado.")
        else:
            seguimiento = seguimiento.sort_values(by=["Nivel", "Score", "Abiertos"], ascending=[False, True, False])
            st.dataframe(
                seguimiento[
                    [
                        "Cliente",
                        "Nivel",
                        "Estado atencion",
                        "Score",
                        "Abiertos",
                        "Alertas incidentes",
                        "Prioridad alta",
                        "Casos sin causa",
                        "Producto principal",
                        "Causa incidente principal",
                    ]
                ],
                use_container_width=True,
                hide_index=True,
            )
            for _, row in seguimiento.iterrows():
                st.warning(
                    f"**{row['Cliente']}** | {row['Estado atencion']} | "
                    f"Abiertos: {row['Abiertos']} | Alertas: {row['Alertas incidentes']} | "
                    f"Casos sin causa: {row['Casos sin causa']}"
                )


def vista_cargar_casos():
    archivo = st.file_uploader("Sube Excel de casos", type=["xlsx"], key="casos_upload")
    if archivo:
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
            cargados, reemplazados, eliminados, meses_reemplazados, duplicados_archivo = guardar_casos(
                df,
                reemplazar_meses=reemplazar_meses,
            )
            if cargados == 0:
                st.error("No se guardaron casos. Revisa que el archivo tenga una columna de numero de caso.")
            else:
                detalle_reemplazo = ""
                if reemplazar_meses:
                    meses = ", ".join(meses_reemplazados) if meses_reemplazados else "sin fechas validas"
                    detalle_reemplazo = f" | Meses reemplazados: {meses} | Registros eliminados antes de cargar: {eliminados}"
                st.success(
                    f"Cargados: {cargados} | Registros existentes actualizados: {reemplazados} | "
                    f"Duplicados/filas sin numero omitidos del archivo: {duplicados_archivo} | "
                    "Los meses no incluidos en el archivo se conservan."
                    f"{detalle_reemplazo}"
                )


def vista_casos():
    df = normalizar_tipificaciones_casos_df(load_casos())
    if not df.empty:
        df = preparar_fechas_dashboard(df)
        df["mes"] = df["_creado_dt_dashboard"].dt.to_period("M").astype(str).replace("NaT", "Sin fecha")

        filtro_col1, filtro_col2, filtro_col3, filtro_col4 = st.columns([1, 1, 1.5, 2])
        with filtro_col1:
            filtro_mes = selector_mes_dashboard(df, "vista_casos_mes")
        if filtro_mes != "Todos":
            df = filtrar_mes_dashboard(df, filtro_mes)

        with filtro_col2:
            estados = sorted(df["estado"].dropna().unique().tolist())
            filtro_estado = st.selectbox("Estado", ["Todos"] + estados, key="estado_casos")
        with filtro_col3:
            clasificaciones = sorted(df["tipificacion"].dropna().unique().tolist())
            filtro_clasificacion = st.selectbox(
                "Clasificacion",
                ["Todos"] + clasificaciones,
                key="clasificacion_casos",
            )
        with filtro_col4:
            filtro_cuenta = st.text_input("Cuenta", key="cuenta_casos")

        if filtro_estado != "Todos":
            df = df[df["estado"] == filtro_estado]
        if filtro_clasificacion != "Todos":
            df = df[df["tipificacion"] == filtro_clasificacion]
        if filtro_cuenta:
            df = df[df["cuenta"].fillna("").str.contains(filtro_cuenta, case=False, na=False)]
        df = df.drop(columns=["_creado_dt_dashboard"], errors="ignore")
        columnas = [
            "numero",
            "estado",
            "mes",
            "cuenta",
            "contacto",
            "descripcion",
            "prioridad",
            "asignado",
            "creado",
            "cerrado",
            "producto",
            "causa",
            "tipificacion",
            "tiempo_respuesta",
            "canal",
            "creado_por",
            "actualizado",
            "codigo_resolucion",
            "notas_resolucion",
            "observaciones_adicionales",
            "observaciones_trabajo",
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
        df["mes"] = df["_creado_dt_dashboard"].dt.to_period("M").astype(str).replace("NaT", "Sin fecha")

        filtro_col1, filtro_col2, filtro_col3, filtro_col4 = st.columns(4)
        with filtro_col1:
            filtro_mes = selector_mes_dashboard(df, "vista_incidentes_mes")
        if filtro_mes != "Todos":
            df = filtrar_mes_dashboard(df, filtro_mes)

        with filtro_col2:
            estados = sorted(df["estado"].dropna().unique().tolist())
            filtro_estado = st.selectbox("Estado", ["Todos"] + estados, key="estado_inc")
        with filtro_col3:
            filtro_tipificacion = st.selectbox(
                "Tipificacion",
                ["Todos"] + sorted(df["tipificacion_auto"].dropna().unique().tolist()),
                key="tip_inc",
            )
        with filtro_col4:
            filtro_alerta = st.selectbox(
                "Es alerta",
                ["Todos"] + sorted(df["es_alerta_auto"].dropna().unique().tolist()),
                key="alerta_inc",
            )

        if filtro_estado != "Todos":
            df = df[df["estado"] == filtro_estado]
        if filtro_tipificacion != "Todos":
            df = df[df["tipificacion_auto"] == filtro_tipificacion]
        if filtro_alerta != "Todos":
            df = df[df["es_alerta_auto"] == filtro_alerta]
        columnas = [
            "numero",
            "solicitante",
            "categoria",
            "prioridad",
            "estado",
            "mes",
            "grupo_asignacion",
            "asignado_a",
            "descripcion",
            "despues_aprobacion",
            "despues_rechazo",
            "duracion_segundos",
            "fecha_vencimiento_sla",
            "tipo_falla",
            "empresa",
            "creado_por",
            "cerrado",
            "escalado_proveedor",
            "servicio_negocio",
            "creado",
            "observaciones_trabajo",
            "observaciones_adicionales",
            "actualizaciones",
            "impacto",
            "lista_notas_trabajo",
            "tipificacion_auto",
            "es_alerta_auto",
            "causa_raiz_auto",
            "prioridad_normalizada",
            "familia_sla",
            "sla_objetivo_horas",
            "estado_sla",
            "duracion_horas",
        ]
        df = df.drop(columns=["_creado_dt_dashboard"], errors="ignore")
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

    st.caption(f"Periodo: {mes_dashboard}")
    render_seguimiento_operativo_incidentes(df)


def vista_administrar_usuarios():
    st.subheader("Administrar usuarios")
    st.caption("Crea usuarios para dar acceso a los dashboards o cambia su rol y estado.")

    usuarios = listar_usuarios()
    if usuarios.empty:
        st.info("Aun no hay usuarios configurados.")
    else:
        tabla = usuarios.copy()
        tabla["active"] = tabla["active"].map({1: "Activo", 0: "Inactivo", True: "Activo", False: "Inactivo"})
        tabla["role"] = tabla["role"].map({"admin": "Admin", "viewer": "Viewer"}).fillna(tabla["role"])
        st.dataframe(tabla, use_container_width=True, hide_index=True)

    st.divider()
    st.markdown("#### Crear o actualizar usuario")
    with st.form("form_usuario"):
        email = st.text_input("Correo", key="usuario_email")
        password = st.text_input(
            "Contrasena nueva",
            type="password",
            help="Minimo 8 caracteres. Para actualizar rol/estado sin cambiar contrasena, dejala vacia.",
            key="usuario_password",
        )
        col_rol, col_estado = st.columns(2)
        with col_rol:
            role = st.selectbox("Rol", ["viewer", "admin"], format_func=lambda x: "Viewer" if x == "viewer" else "Admin")
        with col_estado:
            active = st.checkbox("Activo", value=True)
        guardar = st.form_submit_button("Guardar usuario")

    if guardar:
        if not validar_email(email):
            st.error("Escribe un correo valido.")
        else:
            try:
                guardar_usuario(email, password or None, role=role, active=active)
                st.success("Usuario guardado.")
                st.rerun()
            except ValueError as exc:
                st.error(str(exc))

    st.divider()
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
        type="password",
        key="clave_limpiar_incidentes",
    )
    puede_limpiar = confirmar_limpieza and clave_limpieza == "lina202" and total_incidentes > 0
    if clave_limpieza and clave_limpieza != "lina202":
        st.error("Clave incorrecta.")
    if st.button("Limpiar incidentes", disabled=not puede_limpiar):
        borrados = limpiar_incidentes()
        st.success(f"Incidentes eliminados: {borrados}. Los casos no fueron modificados.")
        st.rerun()

    st.divider()
    st.markdown("#### Quitar acceso")
    usuarios_actuales = listar_usuarios()
    if usuarios_actuales.empty:
        st.info("No hay usuarios para eliminar.")
        return

    email_actual = normalizar_email(st.session_state.get("user"))
    candidatos = [
        email
        for email in usuarios_actuales["email"].tolist()
        if normalizar_email(email) != email_actual
    ]
    if not candidatos:
        st.info("No puedes eliminar tu propio usuario desde aqui.")
        return

    usuario_eliminar = st.selectbox("Usuario", candidatos, key="usuario_eliminar")
    confirmar = st.checkbox("Confirmo que quiero eliminar este usuario", key="confirmar_eliminar_usuario")
    if st.button("Eliminar usuario", disabled=not confirmar):
        eliminar_usuario(usuario_eliminar)
        st.success("Usuario eliminado.")
        st.rerun()


def run_app():
    aplicar_tema_visual()
    init_db()
    if listar_usuarios().empty:
        configurar_primer_admin()
        return
    if not login():
        return

    if st.session_state.role == "admin":
        menu = st.sidebar.selectbox(
            "Menu",
            [
                "Cargar Excel Casos",
                "Casos",
                "Dashboard Casos",
                "Cargar Excel Incidentes",
                "Incidentes",
                "Dashboard Incidentes",
                "Seguimiento Incidentes",
                "Clientes Clave",
                "Administrar Usuarios",
            ],
        )
        st.sidebar.caption(f"Sesion: {st.session_state.user}")
        if st.sidebar.button("Cerrar sesion"):
            st.session_state.clear()
            st.rerun()
    else:
        if st.button("Cerrar sesion"):
            st.session_state.clear()
            st.rerun()
        vista = st.radio(
            "Vista",
            ["Casos", "Incidentes", "Seguimiento incidentes", "Clientes clave"],
            horizontal=True,
            label_visibility="collapsed",
            key="viewer_vista",
        )
        if vista == "Casos":
            ejecutar_con_carga("Casos", dashboard_casos)
        elif vista == "Incidentes":
            ejecutar_con_carga("Incidentes", dashboard_incidentes)
        elif vista == "Seguimiento incidentes":
            ejecutar_con_carga("Seguimiento incidentes", vista_seguimiento_incidentes)
        else:
            ejecutar_con_carga("Clientes clave", dashboard_clientes_clave)
        return

    if menu == "Cargar Excel Casos":
        vista_cargar_casos()
    elif menu == "Casos":
        vista_casos()
    elif menu == "Dashboard Casos":
        dashboard_casos()
    elif menu == "Cargar Excel Incidentes":
        vista_cargar_incidentes()
    elif menu == "Incidentes":
        vista_incidentes()
    elif menu == "Dashboard Incidentes":
        dashboard_incidentes()
    elif menu == "Seguimiento Incidentes":
        vista_seguimiento_incidentes()
    elif menu == "Clientes Clave":
        dashboard_clientes_clave()
    elif menu == "Administrar Usuarios":
        vista_administrar_usuarios()
