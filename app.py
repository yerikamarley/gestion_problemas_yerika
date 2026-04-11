import sqlite3
import pandas as pd
import plotly.express as px
import streamlit as st
from datetime import timedelta
import re

st.set_page_config(page_title="Gestión de Problemas Yerika", layout="wide")

DB = "casos.db"

def estilos_login():
    st.markdown("""
    <style>
    header {visibility: hidden;}
    footer {visibility: hidden;}

    .login-card {
        background: white;
        padding: 36px;
        border-radius: 18px;
        box-shadow: 0 12px 30px rgba(0,0,0,0.16);
        text-align: center;
        max-width: 420px;
        margin: auto;
        width: 100%;
    }

    .login-emoji {
        font-size: 42px;
        margin-bottom: 8px;
    }

    .login-title {
        font-size: 24px;
        font-weight: 700;
        color: #1e6f5c;
        margin-bottom: 6px;
    }

    .login-subtitle {
        font-size: 14px;
        color: #4b5563;
        margin-bottom: 18px;
    }

    div.stButton > button {
        width: 100%;
        border-radius: 10px;
        background: linear-gradient(135deg, #1e6f5c, #2f8f9d);
        color: white;
        font-weight: 600;
        border: none;
        padding: 0.6rem 1rem;
    }

    div.stButton > button:hover {
        background: linear-gradient(135deg, #185a4a, #277985);
        color: white;
    }

    /* RESPONSIVE LOGIN */
    @media (max-width: 768px) {
        .login-card {
            padding: 20px !important;
            border-radius: 12px !important;
        }

        .login-title {
            font-size: 20px !important;
        }

        .login-subtitle {
            font-size: 13px !important;
        }

        div.stButton > button {
            width: 100% !important;
        }
    }
    </style>
    """, unsafe_allow_html=True)

def validar_email(correo):
    return re.match(r"^[a-zA-Z0-9._%+-]+@certicamara\.com$", correo)

def login():
    if "user" not in st.session_state:
        estilos_login()

        _, top_space_2, _ = st.columns([1, 1, 1])
        with top_space_2:
            st.markdown("<br><br><br>", unsafe_allow_html=True)

        # 🔥 CAMBIO RESPONSIVE (ANTES TENÍAS columns)
        col2 = st.container()

        with col2:
            with st.container():
                
                st.markdown('<div class="login-title">Analítica de casos</div>', unsafe_allow_html=True)
                st.markdown('<div class="login-subtitle">ingresa con tu correo corporativo</div>', unsafe_allow_html=True)

                correo = st.text_input("correo corporativo", key="correo_login")

                if st.button("ingresar", key="btn_login"):
                    if not validar_email(correo):
                        st.error("el correo debe ser corporativo")
                        st.markdown("</div>", unsafe_allow_html=True)
                        return False

                    st.session_state.user = correo
                    st.session_state.role = "admin" if correo == "yerika.basto@certicamara.com" else "user"
                    st.rerun()

                st.markdown("</div>", unsafe_allow_html=True)

        return False

    return True

def get_conn():
    return sqlite3.connect(DB, check_same_thread=False)

def init_db():
    conn = get_conn()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS cases (
            numero TEXT PRIMARY KEY,
            descripcion TEXT,
            contacto TEXT,
            cuenta TEXT,
            codigo_resolucion TEXT,
            canal TEXT,
            estado TEXT,
            prioridad TEXT,
            asignado TEXT,
            actualizado TEXT,
            creado_por TEXT,
            creado TEXT,
            producto TEXT,
            cerrado TEXT,
            causa TEXT,
            notas_resolucion TEXT,
            observaciones_adicionales TEXT,
            observaciones_trabajo TEXT,
            tipificacion TEXT,
            tiempo_respuesta TEXT
        )
    """)
    conn.commit()
    conn.close()

def tipificar(row):
    texto = " ".join([
        str(row.get("Breve descripción", "")),
        str(row.get("Causa", "")),
        str(row.get("Notas de resolución", "")),
        str(row.get("Observaciones adicionales", "")),
        str(row.get("Observaciones y notas de trabajo", ""))
    ]).lower()

    if "phishing" in texto:
        return "1 - phishing"
    elif "error" in texto or "falla" in texto:
        return "2 - Soporte Falla"
    elif "contraseña" in texto or "como usar" in texto:
        return "3 - Soporte Uso"
    elif "cert" in texto or "pagar" in texto or "biometria" in texto:
        return "4 - solicitudes"
    elif "incidente" in texto:
        return "5 - incidente"
    elif "externo" in texto:
        return "6 - Plataformas Ext"
    else:
        return "7 - No Aplica"

def tiempo(creado, cerrado):
    try:
        if pd.isna(cerrado):
            return "En proceso"

        inicio = pd.to_datetime(creado)
        fin = pd.to_datetime(cerrado)

        total = timedelta()
        current = inicio

        while current < fin:
            if current.weekday() >= 5:
                current = (current + timedelta(days=1)).replace(hour=8, minute=0, second=0)
                continue

            inicio_dia = current.replace(hour=8, minute=0, second=0)
            fin_dia = current.replace(hour=17, minute=0, second=0)

            if current < inicio_dia:
                current = inicio_dia

            if current >= fin_dia:
                current = (current + timedelta(days=1)).replace(hour=8, minute=0, second=0)
                continue

            limite = min(fin, fin_dia)
            total += (limite - current)
            current = limite

        horas = total.total_seconds() / 3600
        return round(horas, 2)

    except:
        return "Error"

def guardar(df):
    conn = get_conn()
    nuevos, act = 0, 0

    df = df.loc[:, ~df.columns.duplicated()]
    df = df.drop_duplicates(subset=["Número"], keep="last")

    for _, r in df.iterrows():
        num = str(r.get("Número"))
        if not num or num == "nan":
            continue

        data = (
            num,
            str(r.get("Breve descripción")),
            str(r.get("Contacto")),
            str(r.get("Cuenta")),
            str(r.get("Código de resolución")),
            str(r.get("Canal")),
            str(r.get("Estado")),
            str(r.get("Prioridad")),
            str(r.get("Asignado a")),
            str(r.get("Actualizado")),
            str(r.get("Creado por")),
            str(r.get("Creado")),
            str(r.get("Producto")),
            str(r.get("Cerrado")),
            str(r.get("Causa")),
            str(r.get("Notas de resolución")),
            str(r.get("Observaciones adicionales")),
            str(r.get("Observaciones y notas de trabajo")),
            tipificar(r),
            tiempo(r.get("Creado"), r.get("Cerrado"))
        )

        cur = conn.cursor()
        cur.execute("SELECT numero FROM cases WHERE numero=?", (num,))
        if cur.fetchone():
            cur.execute("""
                UPDATE cases SET 
                descripcion=?, contacto=?, cuenta=?, codigo_resolucion=?, canal=?, estado=?, prioridad=?, asignado=?, actualizado=?, creado_por=?, creado=?, producto=?, cerrado=?, causa=?, notas_resolucion=?, observaciones_adicionales=?, observaciones_trabajo=?, tipificacion=?, tiempo_respuesta=?
                WHERE numero=?
            """, data[1:] + (num,))
            act += 1
        else:
            cur.execute("""
                INSERT INTO cases VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, data)
            nuevos += 1

    conn.commit()
    conn.close()
    return nuevos, act

def load():
    conn = get_conn()
    df = pd.read_sql("SELECT * FROM cases", conn)
    conn.close()
    return df

def tarjeta(titulo, valor):
    return f"""
        <div style="
            background-color:#0d2b45;
            padding:20px;
            border-radius:12px;
            text-align:center;
            color:white;
            height:120px;
            display:flex;
            flex-direction:column;
            justify-content:center;
            align-items:center;
        ">
            <div style="font-size:16px; font-weight:600; margin-bottom:8px;">
                {titulo}
            </div>
            <div style="font-size:28px; font-weight:bold;">
                {valor}
            </div>
        </div>
    """

init_db()

if not login():
    st.stop()

st.title("👩 Gestión de Problemas Yerika")

if st.session_state.role == "admin":
    menu = st.sidebar.selectbox("Menú", ["Cargar Excel", "Casos", "Dashboard"])
    if st.sidebar.button("Cerrar sesión"):
        st.session_state.clear()
        st.rerun()
else:
    menu = "Dashboard"
    if st.button("Cerrar sesión"):
        st.session_state.clear()
        st.rerun()

if menu == "Cargar Excel":
    file = st.file_uploader("Sube Excel", type=["xlsx"])
    if file:
        df = pd.read_excel(file)
        st.write(f"Filas detectadas: {len(df)}")
        st.dataframe(df.head())

        if st.button("Procesar"):
            n, a = guardar(df)
            st.success(f"Nuevos: {n} | Actualizados: {a}")

elif menu == "Casos":
    df = load()

    if not df.empty:
        filtro_estado = st.selectbox("Estado", ["Todos"] + list(df["estado"].dropna().unique()))
        filtro_cuenta = st.text_input("Cuenta")

        if filtro_estado != "Todos":
            df = df[df["estado"] == filtro_estado]

        if filtro_cuenta:
            df = df[df["cuenta"].str.contains(filtro_cuenta, case=False, na=False)]

    st.dataframe(df, use_container_width=True)

elif menu == "Dashboard":
    df = load()

    if not df.empty:
        total = len(df)
        cerrados = len(df[df.estado == "Cerrado"])
        abiertos = total - cerrados

        df_cerrados = df[df["estado"] == "Cerrado"]
        tiempos_cerrados = pd.to_numeric(df_cerrados["tiempo_respuesta"], errors="coerce").dropna()

        promedio = round(tiempos_cerrados.mean(), 2) if len(tiempos_cerrados) > 0 else 0

        sla_objetivo = 24
        total_cerrados = len(tiempos_cerrados)
        cumplen = len(tiempos_cerrados[tiempos_cerrados < sla_objetivo])
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
            tip.columns = ["Tipificación", "Cantidad"]
            tip = tip.sort_values(by="Cantidad", ascending=True)

            fig = px.bar(
                tip,
                x="Cantidad",
                y="Tipificación",
                orientation="h",
                text="Cantidad",
                color="Tipificación"
            )

            fig.update_traces(textposition="outside")
            fig.update_layout(showlegend=False)

            st.plotly_chart(fig, use_container_width=True)

        with col2:
            df["creado"] = pd.to_datetime(df["creado"], errors="coerce")
            casos_dia = df.groupby(df["creado"].dt.date).size().reset_index(name="casos")

            fig2 = px.bar(
                casos_dia,
                x="creado",
                y="casos"
            )
            fig2.update_traces(marker_color="#f76c02")

            st.plotly_chart(fig2, use_container_width=True)