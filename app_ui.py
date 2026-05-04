import html
import re

import pandas as pd
import plotly.express as px
import streamlit as st

from app_logic import (
    autenticar_usuario,
    construir_alertas_incidentes,
    eliminar_usuario,
    guardar_casos,
    guardar_incidentes,
    guardar_usuario,
    init_db,
    listar_usuarios,
    load_casos,
    load_incidentes,
    normalizar_texto,
    normalizar_email,
)

UI_PALETTE = {
    "bg": "#f7f9fb",
    "bg_soft": "#eef4f3",
    "surface": "#ffffff",
    "surface_alt": "#f5f8f7",
    "border": "#d9e2df",

    "text": "#2b3438",
    "muted": "#667579",

   
    "yellow": "#b8a15a",
    "yellow_soft": "#d8c98b",

   
    "red": "#a84f55",
    "red_soft": "#c98489",

    "green": "#277267",
    "green_soft": "#7ba99e",
    "blue": "#4f6f73",
    "blue_soft": "#8fa6aa",
}

CHART_COLORS = [
    UI_PALETTE["green"],
    UI_PALETTE["blue"],
    UI_PALETTE["green_soft"],
    UI_PALETTE["blue_soft"],
    UI_PALETTE["yellow"],
    UI_PALETTE["red"],
]

SLA_CASOS_HORAS = 36
SLA_INCIDENTES_HORAS = 24


def preparar_fechas_dashboard(df, columna="creado"):
    trabajo = df.copy()
    trabajo["_creado_dt_dashboard"] = pd.to_datetime(trabajo[columna], errors="coerce")
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
        "Tipificacion": "8 - Instalaciones",
        "Descripcion": "Procesos de instalacion, activacion, agendamiento o citas con tecnicos de instalacion.",
    },
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
            --yellow: {UI_PALETTE["yellow"]};
            --yellow-soft: {UI_PALETTE["yellow_soft"]};
            --red: {UI_PALETTE["red"]};
            --red-soft: {UI_PALETTE["red_soft"]};
            --green: {UI_PALETTE["green"]};
            --green-soft: {UI_PALETTE["green_soft"]};
            --blue: {UI_PALETTE["blue"]};
            --blue-soft: {UI_PALETTE["blue_soft"]};
        }}

        html, body, [data-testid="stAppViewContainer"], .stApp {{
            background: linear-gradient(180deg, var(--bg-soft) 0%, var(--bg) 42%, #ffffff 100%);
            color: var(--text) !important;
            color-scheme: light !important;
        }}

        [data-testid="stHeader"], [data-testid="stToolbar"], [data-testid="stDecoration"] {{
            background: transparent !important;
        }}

        [data-testid="stSidebar"] {{
            background: linear-gradient(180deg, #ffffff 0%, #f3f6f9 100%) !important;
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
            background: rgba(255, 255, 255, 0.96) !important;
            border: 1px solid var(--border) !important;
            border-radius: 8px !important;
            box-shadow: 0 8px 20px rgba(20, 58, 90, 0.05);
            color: var(--text) !important;
        }}

        .stButton > button,
        [data-testid="baseButton-secondary"] {{
            background: var(--green) !important;
            color: white !important;
            border: none !important;
            border-radius: 8px !important;
            font-weight: 700 !important;
            box-shadow: 0 8px 18px rgba(39, 114, 103, 0.16);
        }}

        .stButton > button *,
        [data-testid="baseButton-secondary"] * {{
            color: white !important;
        }}

        .stButton > button:hover {{
            background: #1f5f56 !important;
            color: white !important;
        }}

        [data-testid="stTabs"] button[role="tab"] {{
            border-radius: 8px;
            border: 1px solid var(--border);
            background: rgba(255, 255, 255, 0.9);
            color: var(--muted) !important;
        }}

        [data-testid="stTabs"] button[aria-selected="true"] {{
            background: #edf5f3;
            border-color: var(--green-soft);
            color: var(--green) !important;
        }}

        [data-testid="stTabs"] [data-baseweb="tab-highlight"] {{
            background-color: var(--green) !important;
        }}

        [role="radiogroup"] {{
            gap: 0.5rem;
            flex-wrap: wrap;
            margin-bottom: 1rem;
        }}

        [role="radiogroup"] label {{
            background: rgba(255, 255, 255, 0.9);
            border: 1px solid var(--border);
            border-radius: 8px;
            padding: 0.35rem 0.7rem;
        }}

        [role="radiogroup"] label:has(input:checked) {{
            background: #edf5f3;
            border-color: var(--green-soft);
        }}

        [data-baseweb="tag"] {{
            background-color: #edf5f3 !important;
            border: 1px solid #c7d9d4 !important;
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
            border-color: rgba(20, 58, 90, 0.12) !important;
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
            box-shadow: 0 10px 24px rgba(20, 58, 90, 0.06);
            position: relative;
            overflow: hidden;
        }}

        .kpi-card::before {{
            content: "";
            position: absolute;
            inset: 0 auto auto 0;
            width: 100%;
            height: 4px;
            background: var(--green);
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
            color: var(--green);
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
        plot_bgcolor="rgba(255,255,255,0.94)",
        font=dict(color=UI_PALETTE["text"]),
        title_font=dict(color=UI_PALETTE["green"], size=17),
        margin=dict(l=12, r=12, t=52, b=12),
        legend=dict(bgcolor="rgba(255,255,255,0.82)"),
    )
    fig.update_xaxes(showgrid=True, gridcolor="rgba(20, 58, 90, 0.10)", zeroline=False)
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
            background: linear-gradient(180deg, #eef4f3 0%, #f7f9fb 54%, #ffffff 100%);
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
            background: rgba(255, 255, 255, 0.98);
            padding: 36px 32px;
            border-radius: 8px;
            box-shadow: 0 18px 42px rgba(20, 58, 90, 0.12);
            text-align: center;
            max-width: 420px;
            margin: auto;
            width: 100%;
            border: 1px solid #d8e0e8;
        }

        /* Título */
        .login-title {
            font-size: 28px;
            font-weight: 800;
            color: #277267;
            margin-bottom: 8px;
            text-align: left;
        }

        /* Subtítulo */
        .login-subtitle {
            font-size: 14px;
            color: #5d6b7a;
            margin-bottom: 24px;
            text-align: left;
        }

        /* Inputs */
        input {
            border-radius: 8px !important;
            border: 1px solid #d8e0e8 !important;
            padding: 10px !important;
            transition: all 0.2s ease;
        }

        input:focus {
            border: 1px solid #0f6b5f !important;
            box-shadow: 0 0 0 3px rgba(15, 107, 95, 0.12);
            outline: none;
        }

        /* Botón */
        div.stButton > button {
            width: 100%;
            border-radius: 8px;
            background: #277267;
            color: white;
            font-weight: 700;
            border: none;
            padding: 0.7rem 1rem;
            transition: all 0.25s ease;
            box-shadow: 0 8px 18px rgba(39, 114, 103, 0.18);
        }

        div.stButton > button:hover {
            background: #0f6b5f;
            box-shadow: 0 12px 24px rgba(15, 107, 95, 0.20);
        }

        /* Placeholder */
        ::placeholder {
            color: #8a97a5;
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

    trabajo = df.copy()
    detecciones = trabajo.apply(
        lambda row: detectar_cliente_en_fila(row, ["cuenta"]),
        axis=1,
        result_type="expand",
    )
    detecciones.columns = ["cliente_clave", "fuente_cliente"]
    trabajo[["cliente_clave", "fuente_cliente"]] = detecciones
    trabajo = trabajo[trabajo["cliente_clave"] != ""].copy()
    trabajo["creado_dt"] = pd.to_datetime(trabajo["creado"], errors="coerce")
    trabajo["cerrado_dt"] = pd.to_datetime(trabajo["cerrado"], errors="coerce")
    trabajo["tiempo_respuesta_h"] = pd.to_numeric(trabajo["tiempo_respuesta"], errors="coerce")
    return trabajo


def preparar_incidentes_clientes_clave(df):
    if df.empty:
        trabajo = df.copy()
        trabajo["cliente_clave"] = pd.Series(dtype="object")
        trabajo["fuente_cliente"] = pd.Series(dtype="object")
        trabajo["creado_dt"] = pd.Series(dtype="datetime64[ns]")
        trabajo["cerrado_dt"] = pd.Series(dtype="datetime64[ns]")
        trabajo["duracion_horas_num"] = pd.Series(dtype="float")
        return trabajo

    trabajo = df.copy()
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
    trabajo["creado_dt"] = pd.to_datetime(trabajo["creado"], errors="coerce")
    trabajo["cerrado_dt"] = pd.to_datetime(trabajo["cerrado"], errors="coerce")
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
            casos_cerrados = casos_cliente[casos_cliente["estado"] == "Cerrado"]
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
            incidentes_cerrados = incidentes_cliente[incidentes_cliente["estado"] == "Cerrado"]
            incidentes_abiertos = total_incidentes - len(incidentes_cerrados)
            duraciones_incidentes = incidentes_cerrados["duracion_horas_num"].dropna()
            sla_incidentes = (
                porcentaje(len(duraciones_incidentes[duraciones_incidentes < SLA_INCIDENTES_HORAS]), len(duraciones_incidentes))
                if len(duraciones_incidentes)
                else None
            )
            alertas = len(incidentes_cliente[incidentes_cliente["es_alerta_auto"].fillna("No") == "Si"])
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


def dashboard_casos():
    df = load_casos()
    if df.empty:
        st.info("No hay datos de casos cargados.")
        return

    df = preparar_fechas_dashboard(df)
    mes_dashboard = selector_mes_dashboard(df, "dashboard_casos_mes")
    df = filtrar_mes_dashboard(df, mes_dashboard)
    if df.empty:
        st.info(f"No hay casos cargados para {mes_dashboard}.")
        return

    total = len(df)
    cerrados = len(df[df.estado == "Cerrado"])
    abiertos = total - cerrados

    df_cerrados = df[df["estado"] == "Cerrado"]
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
        tip = df["tipificacion"].value_counts().reset_index()
        tip.columns = ["Tipificacion", "Cantidad"]
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
    st.subheader("Resumen de tipificaciones")
    st.caption("Descripcion breve de cada tipificacion y cantidad actual de casos clasificados en el dashboard.")
    st.dataframe(tabla_resumen_tipificaciones_casos(df), use_container_width=True, hide_index=True)


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

    total = len(df)
    cerrados = len(df[df["estado"] == "Cerrado"])
    abiertos = total - cerrados

    duraciones = pd.to_numeric(df["duracion_horas"], errors="coerce").dropna()
    promedio = round(duraciones.mean(), 2) if len(duraciones) > 0 else 0

    df_cerrados = df[df["estado"] == "Cerrado"]
    duraciones_cerrados = pd.to_numeric(df_cerrados["duracion_horas"], errors="coerce").dropna()
    total_cerrados = len(duraciones_cerrados)
    cumplen = len(duraciones_cerrados[duraciones_cerrados < SLA_INCIDENTES_HORAS])
    porcentaje_sla = round((cumplen / total_cerrados) * 100, 2) if total_cerrados > 0 else 0
    incumplen = total_cerrados - cumplen
    alertas_tipificadas = len(df[df["es_alerta_auto"].fillna("No") == "Si"])

    render_tarjetas(
        [
            ("Total Incidentes", total),
            ("Cerrados", cerrados),
            ("Abiertos", abiertos),
            ("Promedio (h)", promedio),
            (f"SLA <{SLA_INCIDENTES_HORAS}h (%)", f"{porcentaje_sla}%"),
        ]
    )
    st.caption(
        f"Periodo: {mes_dashboard} | Cumplen: {cumplen} | No cumplen: {incumplen} | "
        f"Alertas tipificadas: {alertas_tipificadas}"
    )

    st.divider()
    fila1_col1, fila1_col2 = st.columns(2)

    with fila1_col1:
        tip = df["tipificacion_auto"].fillna("Cliente Interno").value_counts().reset_index()
        tip.columns = ["Tipificacion", "Cantidad"]
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
        fig.update_layout(showlegend=False, title="Incidentes por tipificacion")
        st.plotly_chart(aplicar_estilo_figura(fig, "Incidentes por tipificacion"), use_container_width=True)

    with fila1_col2:
        causas = df["causa_raiz_auto"].replace("", pd.NA).fillna("Sin inferencia").value_counts().reset_index()
        causas.columns = ["Causa raiz inferida", "Cantidad"]
        causas = causas.sort_values(by="Cantidad", ascending=True)
        fig = px.bar(
            causas,
            x="Cantidad",
            y="Causa raiz inferida",
            orientation="h",
            text="Cantidad",
            color_discrete_sequence=[UI_PALETTE["green"]],
        )
        fig.update_traces(textposition="outside")
        fig.update_layout(title="Causa raiz inferida")
        st.plotly_chart(aplicar_estilo_figura(fig, "Causa raiz inferida"), use_container_width=True)

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
            title="Incidentes por dia",
            color_discrete_sequence=[UI_PALETTE["red"]],
        )
        fig.update_traces(marker_color=UI_PALETTE["red"])
        st.plotly_chart(aplicar_estilo_figura(fig, "Incidentes por dia"), use_container_width=True)

    st.divider()
    st.subheader("Cliente Externo")

    df_cliente_externo = df[df["tipificacion_auto"].fillna("Cliente Interno") == "Cliente Externo"].copy()
    df_caso_externo = df[df["tipificacion_auto"].fillna("") == "Caso Cliente Externo"].copy()

    resumen_col1, resumen_col2 = st.columns(2)
    porcentaje_externo = round((len(df_cliente_externo) / total) * 100, 2) if total > 0 else 0
    porcentaje_caso_externo = round((len(df_caso_externo) / total) * 100, 2) if total > 0 else 0

    with resumen_col1:
        st.markdown(tarjeta("Incidentes Cliente Externo", len(df_cliente_externo)), unsafe_allow_html=True)
        st.caption(f"Participacion sobre el total: {porcentaje_externo}%")

    with resumen_col2:
        st.markdown(tarjeta("Casos cargados como incidente", len(df_caso_externo)), unsafe_allow_html=True)
        st.caption(f"Participacion sobre el total: {porcentaje_caso_externo}%")

    if not df_cliente_externo.empty:
        causas_externo = df_cliente_externo["causa_raiz_auto"].replace("", pd.NA).fillna("Sin inferencia").value_counts().reset_index()
        causas_externo.columns = ["Afectacion", "Cantidad"]
        causas_externo = causas_externo.sort_values(by="Cantidad", ascending=True)
        principal_afectacion = causas_externo.iloc[-1]["Afectacion"]
        st.caption(f"Principal afectacion inferida en incidentes reales: {principal_afectacion}")

        fig = px.bar(
            causas_externo,
            x="Cantidad",
            y="Afectacion",
            orientation="h",
            text="Cantidad",
            color_discrete_sequence=[UI_PALETTE["red"]],
            title="Que esta afectando al Cliente Externo",
        )
        fig.update_traces(textposition="outside")
        st.plotly_chart(aplicar_estilo_figura(fig, "Que esta afectando al Cliente Externo"), use_container_width=True)
    elif not df_caso_externo.empty:
        st.info("No hay incidentes reales tipificados como Cliente Externo. Los registros detectados en este grupo fueron reclasificados como casos.")
    else:
        st.info("No hay incidentes tipificados como Cliente Externo en los datos cargados.")

    alertas = construir_alertas_incidentes(df, sla_horas=SLA_INCIDENTES_HORAS)

    if alertas:
        st.divider()
        st.subheader("Alertas")
        st.caption(
            "Estas alertas se construyen solo con los incidentes cargados y documentan recurrencia, "
            "concentracion y cumplimiento de SLA."
        )
        for alerta in alertas:
            st.warning(f"**{alerta['titulo']}**\n\n{alerta['detalle']}")
            if alerta["incidentes"]:
                relacionados = ", ".join(alerta["incidentes"])
                if alerta["incidentes_adicionales"] > 0:
                    relacionados += f" y {alerta['incidentes_adicionales']} mas"
                st.caption(f"Incidentes relacionados: {relacionados}")


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
    abiertos_casos = len(casos[casos["estado"] != "Cerrado"]) if not casos.empty else 0
    abiertos_incidentes = len(incidentes[incidentes["estado"] != "Cerrado"]) if not incidentes.empty else 0
    clientes_activos = len(resumen_actividad)
    clientes_seguimiento = len(resumen_actividad[resumen_actividad["Nivel"].isin(["Amarillo", "Rojo"])])

    casos_cerrados = casos[casos["estado"] == "Cerrado"] if not casos.empty else pd.DataFrame()
    tiempos_casos = casos_cerrados.get("tiempo_respuesta_h", pd.Series(dtype="float")).dropna()
    sla_casos = (
        porcentaje(len(tiempos_casos[tiempos_casos < SLA_CASOS_HORAS]), len(tiempos_casos))
        if len(tiempos_casos)
        else 0
    )

    incidentes_cerrados = incidentes[incidentes["estado"] == "Cerrado"] if not incidentes.empty else pd.DataFrame()
    duraciones_incidentes = incidentes_cerrados.get("duracion_horas_num", pd.Series(dtype="float")).dropna()
    sla_incidentes = (
        porcentaje(
            len(duraciones_incidentes[duraciones_incidentes < SLA_INCIDENTES_HORAS]),
            len(duraciones_incidentes),
        )
        if len(duraciones_incidentes)
        else 0
    )

    render_tarjetas(
        [
            ("Clientes activos", clientes_activos),
            ("Atenciones", total_casos + total_incidentes),
            ("Abiertos", abiertos_casos + abiertos_incidentes),
            (f"SLA casos <{SLA_CASOS_HORAS}h", f"{sla_casos}%"),
            (f"SLA inc. <{SLA_INCIDENTES_HORAS}h", f"{sla_incidentes}%"),
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
                "Verde": UI_PALETTE["green"],
                "Amarillo": UI_PALETTE["yellow"],
                "Rojo": UI_PALETTE["red"],
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
                color_discrete_sequence=[UI_PALETTE["green"], UI_PALETTE["red"]],
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
                color_discrete_sequence=[UI_PALETTE["red"]],
            )
            fig.update_traces(textposition="outside")
            st.plotly_chart(aplicar_estilo_figura(fig, "Causas en incidentes"), use_container_width=True)
        else:
            st.info("No hay incidentes asociados a clientes clave.")

    st.divider()
    tab_resumen, tab_casos, tab_incidentes, tab_seguimiento = st.tabs(
        ["Resumen", "Casos", "Incidentes", "Seguimiento"]
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
        seguimiento = resumen_actividad[resumen_actividad["Nivel"].isin(["Amarillo", "Rojo"])].copy()
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
        if st.button("Procesar casos"):
            cargados, reemplazados = guardar_casos(df)
            st.success(
                f"Cargados: {cargados} | Registros existentes actualizados: {reemplazados} | "
                "Los meses anteriores se conservan."
            )


def vista_casos():
    df = load_casos()
    if not df.empty:
        filtro_estado = st.selectbox("Estado", ["Todos"] + list(df["estado"].dropna().unique()), key="estado_casos")
        filtro_cuenta = st.text_input("Cuenta", key="cuenta_casos")
        if filtro_estado != "Todos":
            df = df[df["estado"] == filtro_estado]
        if filtro_cuenta:
            df = df[df["cuenta"].fillna("").str.contains(filtro_cuenta, case=False, na=False)]
    st.dataframe(df, use_container_width=True, hide_index=True)


def vista_cargar_incidentes():
    archivo = st.file_uploader("Sube Excel de incidentes", type=["xlsx"], key="incidentes_upload")
    if archivo:
        df = pd.read_excel(archivo)
        st.write(f"Filas detectadas: {len(df)}")
        st.dataframe(df.head(), use_container_width=True, hide_index=True)
        if st.button("Procesar incidentes"):
            cargados, reemplazados = guardar_incidentes(df)
            st.success(
                f"Cargados: {cargados} | Registros existentes actualizados: {reemplazados} | "
                "Los meses anteriores se conservan."
            )


def vista_incidentes():
    df = load_incidentes()
    if not df.empty:
        filtro_estado = st.selectbox("Estado", ["Todos"] + list(df["estado"].dropna().unique()), key="estado_inc")
        filtro_tipificacion = st.selectbox(
            "Tipificacion",
            ["Todos"] + sorted(df["tipificacion_auto"].dropna().unique().tolist()),
            key="tip_inc",
        )
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
            "duracion_horas",
        ]
        df = df[[col for col in columnas if col in df.columns]]
    st.dataframe(df, use_container_width=True, hide_index=True)


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
            ["Casos", "Incidentes", "Clientes clave"],
            horizontal=True,
            label_visibility="collapsed",
            key="viewer_vista",
        )
        if vista == "Casos":
            ejecutar_con_carga("Casos", dashboard_casos)
        elif vista == "Incidentes":
            ejecutar_con_carga("Incidentes", dashboard_incidentes)
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
    elif menu == "Clientes Clave":
        dashboard_clientes_clave()
    elif menu == "Administrar Usuarios":
        vista_administrar_usuarios()
