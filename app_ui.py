import re

import pandas as pd
import plotly.express as px
import streamlit as st

from app_logic import (
    ADMIN_EMAIL,
    construir_alertas_incidentes,
    guardar_casos,
    guardar_incidentes,
    init_db,
    load_casos,
    load_incidentes,
)

UI_PALETTE = {
    "bg": "#fff8f0",
    "bg_soft": "#fff3e0",
    "surface": "#ffffff",
    "surface_alt": "#fff3e0",
    "border": "#eeeeee",

    "text": "#2f1b0c",
    "muted": "#7a5a3a",

   
    "yellow": "#ff9800",        
    "yellow_soft": "#ffc340",   

   
    "red": "#e00000",           
    "red_soft": "#ff0000",      

    "green": "#2f6b45",
    "green_soft": "#8bb174",
}

CHART_COLORS = [
    UI_PALETTE["green"],
    UI_PALETTE["yellow"],
    UI_PALETTE["red"],
    UI_PALETTE["green_soft"],
    UI_PALETTE["red_soft"],
]


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
        }}

        html, body, [data-testid="stAppViewContainer"], .stApp {{
            background:
                radial-gradient(circle at top left, rgba(212, 160, 23, 0.16), transparent 24%),
                radial-gradient(circle at top right, rgba(47, 107, 69, 0.10), transparent 22%),
                linear-gradient(180deg, var(--bg-soft) 0%, var(--bg) 100%);
            color: var(--text) !important;
            color-scheme: light !important;
        }}

        [data-testid="stHeader"], [data-testid="stToolbar"], [data-testid="stDecoration"] {{
            background: transparent !important;
        }}

        [data-testid="stSidebar"] {{
            background: linear-gradient(180deg, #f5eed7 0%, #efe4c5 100%) !important;
            border-right: 1px solid var(--border);
        }}

        .block-container {{
            padding-top: 1.6rem;
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
            background: rgba(255, 255, 255, 0.86) !important;
            border: 1px solid var(--border) !important;
            border-radius: 14px !important;
            box-shadow: 0 10px 24px rgba(47, 42, 35, 0.05);
            color: var(--text) !important;
        }}

        .stButton > button,
        [data-testid="baseButton-secondary"] {{
            background: linear-gradient(135deg, var(--green), #3d8258) !important;
            color: white !important;
            border: none !important;
            border-radius: 12px !important;
            font-weight: 700 !important;
            box-shadow: 0 10px 24px rgba(47, 107, 69, 0.18);
        }}

        .stButton > button:hover {{
            background: linear-gradient(135deg, #285a3b, var(--green)) !important;
            color: white !important;
        }}

        [data-testid="stTabs"] button[role="tab"] {{
            border-radius: 12px;
            border: 1px solid var(--border);
            background: rgba(255, 255, 255, 0.7);
        }}

        [data-testid="stTabs"] button[aria-selected="true"] {{
            background: linear-gradient(180deg, #fff8dd, #f8efd0);
            border-color: #d6c48e;
            color: var(--green) !important;
        }}

        .stDivider {{
            border-color: rgba(212, 160, 23, 0.25) !important;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def aplicar_estilo_figura(fig, titulo=None):
    fig.update_layout(
        title=titulo,
        paper_bgcolor="rgba(255,255,255,0)",
        plot_bgcolor="rgba(255,255,255,0.78)",
        font=dict(color=UI_PALETTE["text"]),
        title_font=dict(color=UI_PALETTE["green"], size=18),
        margin=dict(l=12, r=12, t=52, b=12),
        legend=dict(bgcolor="rgba(255,255,255,0.65)"),
    )
    fig.update_xaxes(showgrid=True, gridcolor="rgba(212, 160, 23, 0.16)", zeroline=False)
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
            background: linear-gradient(135deg, #fff8f0, #fff3e0);
        }

        /* Card principal */
        .login-card {
            background: rgba(255, 255, 255, 0.9);
            backdrop-filter: blur(10px);
            padding: 40px 32px;
            border-radius: 20px;
            box-shadow: 0 20px 40px rgba(0, 0, 0, 0.08);
            text-align: center;
            max-width: 420px;
            margin: auto;
            width: 100%;
            border: 1px solid #ffd9a8;
            transition: all 0.3s ease;
        }

        .login-card:hover {
            transform: translateY(-4px);
            box-shadow: 0 28px 50px rgba(0, 0, 0, 0.12);
        }

        /* Título */
        .login-title {
            font-size: 26px;
            font-weight: 700;
            color: #2f1b0c;
            margin-bottom: 6px;
        }

        /* Subtítulo */
        .login-subtitle {
            font-size: 14px;
            color: #7a5a3a;
            margin-bottom: 24px;
        }

        /* Inputs */
        input {
            border-radius: 10px !important;
            border: 1px solid #ffd9a8 !important;
            padding: 10px !important;
            transition: all 0.2s ease;
        }

        input:focus {
            border: 1px solid #ff9800 !important;
            box-shadow: 0 0 0 2px rgba(255, 152, 0, 0.2);
            outline: none;
        }

        /* Botón */
        div.stButton > button {
            width: 100%;
            border-radius: 12px;
            background: linear-gradient(135deg, #ff9800, #e00000);
            color: white;
            font-weight: 600;
            border: none;
            padding: 0.7rem 1rem;
            transition: all 0.25s ease;
            box-shadow: 0 8px 18px rgba(224, 0, 0, 0.2);
        }

        div.stButton > button:hover {
            transform: translateY(-2px);
            background: linear-gradient(135deg, #ffc340, #ff0000);
            box-shadow: 0 12px 24px rgba(224, 0, 0, 0.3);
        }

        /* Placeholder */
        ::placeholder {
            color: #a68b6b;
            font-size: 13px;
        }

        /* Responsive */
        @media (max-width: 768px) {
            .login-card {
                padding: 24px !important;
                border-radius: 16px !important;
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
    return re.match(r"^[a-zA-Z0-9._%+-]+@certicamara\.com$", str(correo).strip())


def login():
    if "user" in st.session_state:
        return True

    estilos_login()
    _, top_space_2, _ = st.columns([1, 1, 1])
    with top_space_2:
        st.markdown("<br><br><br>", unsafe_allow_html=True)

    col2 = st.container()
    with col2:
        with st.container():
            
            st.markdown('<div class="login-title">Analitica de casos</div>', unsafe_allow_html=True)
            st.markdown('<div class="login-subtitle">Ingresa con tu correo corporativo</div>', unsafe_allow_html=True)
            correo = st.text_input("Correo corporativo", key="correo_login")

            if st.button("Ingresar", key="btn_login"):
                if not validar_email(correo):
                    st.error("El correo debe ser corporativo")
                    st.markdown("</div>", unsafe_allow_html=True)
                    return False

                st.session_state.user = correo
                st.session_state.role = "admin" if correo.lower() == ADMIN_EMAIL else "user"
                st.rerun()

            st.markdown("</div>", unsafe_allow_html=True)

    return False


def tarjeta(titulo, valor):
    return f"""
        <div style="
            background: linear-gradient(180deg, {UI_PALETTE["surface"]} 0%, {UI_PALETTE["surface_alt"]} 100%);
            padding:20px;
            border-radius:16px;
            text-align:center;
            color:{UI_PALETTE["text"]};
            height:120px;
            display:flex;
            flex-direction:column;
            justify-content:center;
            align-items:center;
            border: 1px solid {UI_PALETTE["border"]};
            box-shadow: 0 16px 28px rgba(47, 42, 35, 0.06);
            position: relative;
            overflow: hidden;
        ">
            <div style="
                position:absolute;
                inset:0 auto auto 0;
                width:100%;
                height:8px;
                background: linear-gradient(90deg, {UI_PALETTE["green"]}, {UI_PALETTE["yellow"]}, {UI_PALETTE["red"]});
            "></div>
            <div style="font-size:15px; font-weight:600; margin-bottom:8px; color:{UI_PALETTE['muted']};">{titulo}</div>
            <div style="font-size:30px; font-weight:800; color:{UI_PALETTE['green']};">{valor}</div>
        </div>
    """


def dashboard_casos():
    df = load_casos()
    if df.empty:
        st.info("No hay datos de casos cargados.")
        return

    total = len(df)
    cerrados = len(df[df.estado == "Cerrado"])
    abiertos = total - cerrados

    df_cerrados = df[df["estado"] == "Cerrado"]
    tiempos_cerrados = pd.to_numeric(df_cerrados["tiempo_respuesta"], errors="coerce").dropna()
    promedio = round(tiempos_cerrados.mean(), 2) if len(tiempos_cerrados) > 0 else 0

    total_cerrados = len(tiempos_cerrados)
    cumplen = len(tiempos_cerrados[tiempos_cerrados < 24])
    porcentaje_sla = round((cumplen / total_cerrados) * 100, 2) if total_cerrados > 0 else 0
    incumplen = total_cerrados - cumplen

    col1, col2, col3, col4, col5 = st.columns(5)
    col1.markdown(tarjeta("Total Casos", total), unsafe_allow_html=True)
    col2.markdown(tarjeta("Cerrados", cerrados), unsafe_allow_html=True)
    col3.markdown(tarjeta("Abiertos", abiertos), unsafe_allow_html=True)
    col4.markdown(tarjeta("Promedio (h)", promedio), unsafe_allow_html=True)
    col5.markdown(tarjeta("SLA <24h (%)", f"{porcentaje_sla}%"), unsafe_allow_html=True)
    st.caption(f"Cumplen: {cumplen} | No cumplen: {incumplen}")

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
        serie["creado"] = pd.to_datetime(serie["creado"], errors="coerce")
        casos_dia = serie.groupby(serie["creado"].dt.date).size().reset_index(name="casos")
        fig = px.bar(casos_dia, x="creado", y="casos", color_discrete_sequence=[UI_PALETTE["yellow"]])
        fig.update_traces(marker_color=UI_PALETTE["yellow"])
        st.plotly_chart(aplicar_estilo_figura(fig, "Casos por dia"), use_container_width=True)


def dashboard_incidentes():
    df = load_incidentes()
    if df.empty:
        st.info("No hay datos de incidentes cargados.")
        return

    total = len(df)
    cerrados = len(df[df["estado"] == "Cerrado"])
    abiertos = total - cerrados

    duraciones = pd.to_numeric(df["duracion_horas"], errors="coerce").dropna()
    promedio = round(duraciones.mean(), 2) if len(duraciones) > 0 else 0

    df_cerrados = df[df["estado"] == "Cerrado"]
    duraciones_cerrados = pd.to_numeric(df_cerrados["duracion_horas"], errors="coerce").dropna()
    total_cerrados = len(duraciones_cerrados)
    cumplen = len(duraciones_cerrados[duraciones_cerrados < 24])
    porcentaje_sla = round((cumplen / total_cerrados) * 100, 2) if total_cerrados > 0 else 0
    incumplen = total_cerrados - cumplen
    alertas_tipificadas = len(df[df["es_alerta_auto"].fillna("No") == "Si"])

    col1, col2, col3, col4, col5 = st.columns(5)
    col1.markdown(tarjeta("Total Incidentes", total), unsafe_allow_html=True)
    col2.markdown(tarjeta("Cerrados", cerrados), unsafe_allow_html=True)
    col3.markdown(tarjeta("Abiertos", abiertos), unsafe_allow_html=True)
    col4.markdown(tarjeta("Promedio (h)", promedio), unsafe_allow_html=True)
    col5.markdown(tarjeta("SLA <24h (%)", f"{porcentaje_sla}%"), unsafe_allow_html=True)
    st.caption(f"Cumplen: {cumplen} | No cumplen: {incumplen} | Alertas tipificadas: {alertas_tipificadas}")

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
        serie["creado"] = pd.to_datetime(serie["creado"], errors="coerce")
        incidentes_dia = serie.groupby(serie["creado"].dt.date).size().reset_index(name="incidentes")
        fig = px.bar(
            incidentes_dia,
            x="creado",
            y="incidentes",
            title="Incidentes por dia",
            color_discrete_sequence=[UI_PALETTE["red"]],
        )
        fig.update_traces(marker_color=UI_PALETTE["red"])
        st.plotly_chart(aplicar_estilo_figura(fig, "Incidentes por dia"), use_container_width=True)

    st.divider()
    st.subheader("Impacto en Cliente Externo")

    df_cliente_externo = df[df["tipificacion_auto"].fillna("Cliente Interno") == "Cliente Externo"].copy()
    if not df_cliente_externo.empty:
        causas_externo = df_cliente_externo["causa_raiz_auto"].replace("", pd.NA).fillna("Sin inferencia").value_counts().reset_index()
        causas_externo.columns = ["Afectacion", "Cantidad"]
        causas_externo = causas_externo.sort_values(by="Cantidad", ascending=True)
        principal_afectacion = causas_externo.iloc[-1]["Afectacion"]
        porcentaje_externo = round((len(df_cliente_externo) / total) * 100, 2) if total > 0 else 0

        externo_col1, externo_col2 = st.columns([1, 2])
        with externo_col1:
            st.markdown(tarjeta("Incidentes Cliente Externo", len(df_cliente_externo)), unsafe_allow_html=True)
            st.caption(f"Participacion sobre el total: {porcentaje_externo}%")
            st.caption(f"Principal afectacion inferida: {principal_afectacion}")

        with externo_col2:
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
    else:
        st.info("No hay incidentes tipificados como Cliente Externo en los datos cargados.")

    alertas = construir_alertas_incidentes(df, sla_horas=24)

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


def vista_cargar_casos():
    archivo = st.file_uploader("Sube Excel de casos", type=["xlsx"], key="casos_upload")
    if archivo:
        df = pd.read_excel(archivo)
        st.write(f"Filas detectadas: {len(df)}")
        st.dataframe(df.head(), use_container_width=True, hide_index=True)
        if st.button("Procesar casos"):
            nuevos, actualizados = guardar_casos(df)
            st.success(f"Nuevos: {nuevos} | Actualizados: {actualizados}")


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
            nuevos, actualizados = guardar_incidentes(df)
            st.success(f"Nuevos: {nuevos} | Actualizados: {actualizados}")


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


def run_app():
    aplicar_tema_visual()
    init_db()
    if not login():
        return

    st.markdown(
        f"""
        <div style="
            background: linear-gradient(135deg, rgba(255,255,255,0.92), rgba(251,247,234,0.92));
            border: 1px solid {UI_PALETTE["border"]};
            border-radius: 20px;
            padding: 1rem 1.2rem;
            box-shadow: 0 18px 32px rgba(47, 42, 35, 0.06);
            margin-bottom: 1rem;
        ">
            <div style="font-size: 2rem; font-weight: 800; color: {UI_PALETTE['green']};">
                Gestion casos e incidentes Yerika
            </div>
            <div style="font-size: 0.98rem; color: {UI_PALETTE['muted']}; margin-top: 0.2rem;">
                
            
        </div>
        """,
        unsafe_allow_html=True,
    )

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
            ],
        )
        if st.sidebar.button("Cerrar sesion"):
            st.session_state.clear()
            st.rerun()
    else:
        if st.button("Cerrar sesion"):
            st.session_state.clear()
            st.rerun()
        tabs = st.tabs(["Dashboard Casos", "Dashboard Incidentes"])
        with tabs[0]:
            dashboard_casos()
        with tabs[1]:
            dashboard_incidentes()
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
