import sqlite3
import pandas as pd
import plotly.express as px
import streamlit as st
from datetime import timedelta
import re
import unicodedata

st.set_page_config(
    page_title="Control de casos e incidentes",
    page_icon=":material/fact_check:",
    layout="wide",
)

DB = "data.db"


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

        col2 = st.container()

        with col2:
            with st.container():
                st.markdown('<div class="login-card">', unsafe_allow_html=True)
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


def ensure_table_columns(conn, table_name, columns):
    existentes = {
        row[1]
        for row in conn.execute(f"PRAGMA table_info({table_name})").fetchall()
    }

    for nombre, tipo in columns.items():
        if nombre not in existentes:
            conn.execute(f"ALTER TABLE {table_name} ADD COLUMN {nombre} {tipo}")


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

    conn.execute("""
        CREATE TABLE IF NOT EXISTS incidents (
            numero TEXT PRIMARY KEY,
            solicitante TEXT,
            breve_descripcion TEXT,
            categoria TEXT,
            prioridad TEXT,
            estado TEXT,
            grupo_asignacion TEXT,
            asignado_a TEXT,
            descripcion TEXT,
            despues_aprobacion TEXT,
            despues_rechazo TEXT,
            duracion_segundos REAL,
            minutos REAL,
            fecha_vencimiento_sla TEXT,
            tipo_falla TEXT,
            empresa TEXT,
            creado_por TEXT,
            cerrado TEXT,
            escalado_proveedor TEXT,
            servicio_negocio TEXT,
            creado TEXT,
            observaciones_trabajo TEXT,
            observaciones_adicionales TEXT,
            actualizaciones TEXT,
            impacto TEXT,
            lista_notas_trabajo TEXT,
            tipificacion_original TEXT,
            causa_raiz_original TEXT,
            tipificacion_auto TEXT,
            causa_raiz_auto TEXT,
            es_alerta_auto TEXT,
            duracion_horas REAL
        )
    """)

    ensure_table_columns(conn, "incidents", {
        "solicitante": "TEXT",
        "breve_descripcion": "TEXT",
        "categoria": "TEXT",
        "prioridad": "TEXT",
        "estado": "TEXT",
        "grupo_asignacion": "TEXT",
        "asignado_a": "TEXT",
        "descripcion": "TEXT",
        "despues_aprobacion": "TEXT",
        "despues_rechazo": "TEXT",
        "duracion_segundos": "REAL",
        "minutos": "REAL",
        "fecha_vencimiento_sla": "TEXT",
        "tipo_falla": "TEXT",
        "empresa": "TEXT",
        "creado_por": "TEXT",
        "cerrado": "TEXT",
        "escalado_proveedor": "TEXT",
        "servicio_negocio": "TEXT",
        "creado": "TEXT",
        "observaciones_trabajo": "TEXT",
        "observaciones_adicionales": "TEXT",
        "actualizaciones": "TEXT",
        "impacto": "TEXT",
        "lista_notas_trabajo": "TEXT",
        "tipificacion_original": "TEXT",
        "causa_raiz_original": "TEXT",
        "tipificacion_auto": "TEXT",
        "causa_raiz_auto": "TEXT",
        "es_alerta_auto": "TEXT",
        "duracion_horas": "REAL"
    })

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

    except Exception:
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


def segundos_a_horas(valor):
    try:
        return round(float(valor) / 3600, 2)
    except Exception:
        return None


def limpiar_incidentes(df):
    df = df.copy()
    df = df.loc[:, ~df.columns.duplicated()]
    df = df.loc[:, ~df.columns.str.contains("^Unnamed", na=False)]

    columnas_necesarias = {
        "Solicitante": None,
        "Categoria": None,
        "Prioridad": None,
        "Estado": None,
        "Grupo de asignación": None,
        "Asignado a": None,
        "Descripción": None,
        "Después de la aprobación": None,
        "Después del rechazo": None,
        "Duración segundos": None,
        "Duración": None,
        "Fecha de vencimiento del SLA": None,
        "Tipo de falla": None,
        "Empresa": None,
        "Creado por": None,
        "Cerrado": None,
        "Escalado a proveedor": None,
        "Servicio de Negocio": None,
        "Creado": None,
        "Observaciones y notas de trabajo": None,
        "Observaciones adicionales": None,
        "Actualizaciones": None,
        "Impacto": None,
        "Lista de notas de trabajo": None,
        "Breve descripción": None,
        "Minutos": None
    }

    for col, val in columnas_necesarias.items():
        if col not in df.columns:
            df[col] = val

    if df["Duración segundos"].isna().all() and "Duración" in df.columns:
        df["Duración segundos"] = df["Duración"]

    df = df.replace({pd.NA: None})
    df = df.drop_duplicates(subset=["Número"], keep="last")
    return df


def tipificacion_incidente(row):
    categoria = str(row.get("Categoria", "")).strip().lower()

    texto = " ".join([
        str(row.get("Grupo de asignación", "")),
        str(row.get("Asignado a", "")),
        str(row.get("Descripción", "")),
        str(row.get("Breve descripción", "")),
        str(row.get("Observaciones y notas de trabajo", "")),
        str(row.get("Observaciones adicionales", "")),
        str(row.get("Después de la aprobación", "")),
        str(row.get("Después del rechazo", "")),
        str(row.get("Servicio de Negocio", "")),
        str(row.get("Tipo de falla", "")),
        str(row.get("Impacto", ""))
    ]).lower()

    if "consulta" in categoria:
        return "Consulta"

    if "solicitud" in categoria:
        return "Solicitud"

    if "incidente" in categoria:
        if "noc" in texto:
            return "NOC"
        if any(palabra in texto for palabra in [
            "firma", "firmar", "cliente", "usuario", "certificado",
            "certificados", "token", "portal", "ocsp", "tsa",
            "clave segura", "certitoken"
        ]):
            return "Cliente Externo"
        if any(palabra in texto for palabra in [
            "ldap", "directorio activo", "active directory", "servidor",
            "infraestructura", "red interna", "vpn", "base de datos",
            "bd", "interno", "correo interno", "aplicativo interno"
        ]):
            return "Cliente Interno"
        return "Cliente Interno"

    if "noc" in texto:
        return "NOC"

    if any(palabra in texto for palabra in [
        "firma", "firmar", "cliente", "usuario", "certificado",
        "certificados", "token", "portal", "ocsp", "tsa",
        "clave segura", "certitoken"
    ]):
        return "Cliente Externo"

    if any(palabra in texto for palabra in [
        "ldap", "directorio activo", "active directory", "servidor",
        "infraestructura", "red interna", "vpn", "base de datos",
        "bd", "interno", "correo interno", "aplicativo interno"
    ]):
        return "Cliente Interno"

    return "Cliente Interno"


def causa_raiz_incidente(row):
    texto = " ".join([
        str(row.get("Descripción", "")),
        str(row.get("Breve descripción", "")),
        str(row.get("Tipo de falla", "")),
        str(row.get("Observaciones y notas de trabajo", "")),
        str(row.get("Observaciones adicionales", "")),
        str(row.get("Después de la aprobación", "")),
        str(row.get("Después del rechazo", "")),
        str(row.get("Lista de notas de trabajo", "")),
        str(row.get("Actualizaciones", ""))
    ]).lower()

    if "ocsp" in texto:
        return "Servicios OCSP"
    elif "certitoken" in texto or "token" in texto:
        return "Servicios Certitoken"
    elif "clave segura" in texto:
        return "Clave Segura"
    elif "lentitud" in texto or "degradación" in texto or "degradacion" in texto:
        return "Degradación del servicio"
    elif "firma" in texto or "firmar" in texto:
        return "Error al firmar"
    elif "tsa" in texto:
        return "Servicios TSA"
    elif "ssl" in texto or "certificado ssl" in texto:
        return "Certificado SSL"
    elif "portal" in texto and "rpost" in texto:
        return "Portal RPOST"
    elif "timeout" in texto or "time out" in texto:
        return "Timeout"
    elif "caída" in texto or "caida" in texto:
        return "Caída del servicio"
    elif "proveedor" in texto:
        return "Dependencia de proveedor"
    elif "ldap" in texto:
        return "LDAP / Directorio Activo"
    elif "correo" in texto:
        return "Servicio de correo"
    elif "red" in texto or "conexion" in texto or "conexión" in texto:
        return "Red / Conectividad"
    else:
        return "Otros"


def valor_campo(row, *claves):
    for clave in claves:
        if clave in row:
            valor = row.get(clave)
            if valor is not None and not pd.isna(valor):
                return valor
    return ""


def normalizar_texto(valor):
    if valor is None or pd.isna(valor):
        return ""

    texto = str(valor).strip().lower()
    texto = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", texto)


def unir_campos(row, grupos_claves):
    partes = []

    for claves in grupos_claves:
        texto = normalizar_texto(valor_campo(row, *claves))
        if texto:
            partes.append(texto)

    return " ".join(partes)


def contiene_alguna(texto, palabras):
    return any(palabra in texto for palabra in palabras)


CAMPOS_TIPIFICACION_INCIDENTE = [
    ("Grupo de asignaciÃ³n", "grupo_asignacion"),
    ("Asignado a", "asignado_a"),
    ("DescripciÃ³n", "descripcion"),
    ("Breve descripciÃ³n", "breve_descripcion"),
    ("Observaciones y notas de trabajo", "observaciones_trabajo"),
    ("Observaciones adicionales", "observaciones_adicionales"),
    ("DespuÃ©s de la aprobaciÃ³n", "despues_aprobacion"),
    ("DespuÃ©s del rechazo", "despues_rechazo"),
    ("Servicio de Negocio", "servicio_negocio"),
    ("Tipo de falla", "tipo_falla"),
    ("Impacto", "impacto")
]


CAMPOS_CAUSA_INCIDENTE = [
    ("DescripciÃ³n", "descripcion"),
    ("Breve descripciÃ³n", "breve_descripcion"),
    ("Tipo de falla", "tipo_falla"),
    ("Observaciones y notas de trabajo", "observaciones_trabajo"),
    ("Observaciones adicionales", "observaciones_adicionales"),
    ("DespuÃ©s de la aprobaciÃ³n", "despues_aprobacion"),
    ("DespuÃ©s del rechazo", "despues_rechazo"),
    ("Lista de notas de trabajo", "lista_notas_trabajo"),
    ("Actualizaciones", "actualizaciones")
]


CAMPOS_ALERTA_INCIDENTE = CAMPOS_CAUSA_INCIDENTE + [
    ("Grupo de asignaciÃ³n", "grupo_asignacion"),
    ("Servicio de Negocio", "servicio_negocio"),
    ("Impacto", "impacto"),
    ("Estado", "estado")
]


PALABRAS_CLIENTE_EXTERNO = [
    "firma", "firmar", "cliente", "usuario", "certificado",
    "certificados", "token", "portal", "ocsp", "tsa",
    "clave segura", "certitoken"
]


PALABRAS_CLIENTE_INTERNO = [
    "ldap", "directorio activo", "active directory", "servidor",
    "infraestructura", "red interna", "vpn", "base de datos",
    "bd", "interno", "correo interno", "aplicativo interno"
]


PALABRAS_ALERTA = [
    "alerta", "alarma", "monitoreo", "monitor", "zabbix",
    "grafana", "prometheus", "observabilidad", "cpu", "memoria",
    "disco", "latencia", "packet loss", "perdida de paquetes",
    "indisponibilidad", "caida", "caido", "down", "critico",
    "critical", "warning"
]


def tipificacion_incidente(row):
    categoria = normalizar_texto(valor_campo(row, "Categoria", "categoria"))
    texto = unir_campos(row, CAMPOS_TIPIFICACION_INCIDENTE)

    if "consulta" in categoria:
        return "Consulta"

    if "solicitud" in categoria:
        return "Solicitud"

    if "noc" in texto:
        return "NOC"

    if "incidente" in categoria:
        if contiene_alguna(texto, PALABRAS_CLIENTE_EXTERNO):
            return "Cliente Externo"
        if contiene_alguna(texto, PALABRAS_CLIENTE_INTERNO):
            return "Cliente Interno"
        return "Cliente Interno"

    if contiene_alguna(texto, PALABRAS_CLIENTE_EXTERNO):
        return "Cliente Externo"

    if contiene_alguna(texto, PALABRAS_CLIENTE_INTERNO):
        return "Cliente Interno"

    return "Cliente Interno"


def es_alerta_incidente(row, tipificacion=None):
    if tipificacion is None:
        tipificacion = tipificacion_incidente(row)

    texto = unir_campos(row, CAMPOS_ALERTA_INCIDENTE)

    if tipificacion == "NOC":
        return "Si"

    if contiene_alguna(texto, PALABRAS_ALERTA):
        return "Si"

    return "No"


def causa_raiz_incidente(row, tipificacion=None, es_alerta=None):
    if tipificacion is None:
        tipificacion = tipificacion_incidente(row)

    if es_alerta is None:
        es_alerta = es_alerta_incidente(row, tipificacion)

    texto = unir_campos(row, CAMPOS_CAUSA_INCIDENTE)

    if "ocsp" in texto:
        detalle = "Servicios OCSP con indisponibilidad o validacion intermitente"
    elif "certitoken" in texto or "token" in texto:
        detalle = "Servicios Certitoken con falla de autenticacion, firma o uso del token"
    elif "clave segura" in texto:
        detalle = "Clave Segura con falla de acceso o disponibilidad"
    elif contiene_alguna(texto, ["ldap", "directorio activo", "active directory"]):
        detalle = "LDAP o Directorio Activo con falla de autenticacion"
    elif contiene_alguna(texto, ["base de datos", "bd", "database", "sql"]):
        detalle = "Base de datos o integracion interna con errores"
    elif contiene_alguna(texto, ["correo", "smtp", "mail"]):
        detalle = "Servicio de correo o notificaciones con fallas"
    elif contiene_alguna(texto, ["ssl", "certificado ssl", "certificado vencido", "cadena de certificados"]):
        detalle = "Certificado SSL o cadena de certificados con errores"
    elif contiene_alguna(texto, ["tsa", "timestamp", "sello de tiempo"]):
        detalle = "Servicios TSA con falla de sellado de tiempo"
    elif "portal" in texto and "rpost" in texto:
        detalle = "Portal RPOST con falla funcional o de acceso"
    elif contiene_alguna(texto, ["cpu", "memoria", "disco", "filesystem", "fs full", "espacio", "swap"]):
        detalle = "Infraestructura con consumo alto de recursos"
    elif contiene_alguna(texto, ["red", "conexion", "conectividad", "vpn", "packet loss", "perdida de paquetes", "latencia", "enlace"]):
        detalle = "Infraestructura de red o conectividad degradada"
    elif contiene_alguna(texto, ["caida", "caido", "indisponibilidad", "no responde", "fuera de servicio", "down"]):
        detalle = "Caida o indisponibilidad del servicio"
    elif contiene_alguna(texto, ["lentitud", "degradacion", "degradado", "timeout", "time out", "intermitencia", "intermitente"]):
        detalle = "Degradacion o intermitencia del servicio"
    elif contiene_alguna(texto, ["firma", "firmar", "rechazo de firma"]):
        detalle = "Proceso de firma con error o rechazo"
    elif contiene_alguna(texto, ["proveedor", "tercero", "externo"]):
        detalle = "Dependencia de proveedor o tercero"
    elif contiene_alguna(texto, ["alerta", "alarma", "monitoreo", "monitor", "zabbix", "grafana", "prometheus", "observabilidad"]):
        detalle = "Monitoreo detecta comportamiento anomalo"
    else:
        detalle = "Analisis pendiente para precisar la causa raiz"

    if tipificacion == "NOC":
        return f"Alerta NOC - {detalle}"

    if es_alerta == "Si" and detalle == "Analisis pendiente para precisar la causa raiz":
        return f"Alerta operativa - {detalle}"

    return detalle


def clasificacion_incidente(row):
    tipificacion = tipificacion_incidente(row)
    es_alerta = es_alerta_incidente(row, tipificacion)
    causa_raiz = causa_raiz_incidente(row, tipificacion, es_alerta)
    return tipificacion, es_alerta, causa_raiz


def normalizar_clave(valor):
    texto = unicodedata.normalize("NFKD", str(valor)).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", texto.strip().lower())


def valor_campo(row, *claves):
    claves_disponibles = {
        normalizar_clave(col): col
        for col in row.index
    }

    for clave in claves:
        clave_real = claves_disponibles.get(normalizar_clave(clave))

        if clave_real is not None:
            valor = row.get(clave_real)
            if valor is not None and not pd.isna(valor):
                return valor
    return ""


def normalizar_texto(valor):
    if valor is None or pd.isna(valor):
        return ""

    return normalizar_clave(valor)


CAMPOS_TIPIFICACION_INCIDENTE = [
    ("grupo de asignacion", "grupo_asignacion"),
    ("asignado a", "asignado_a"),
    ("descripcion", "descripcion"),
    ("breve descripcion", "breve_descripcion"),
    ("observaciones y notas de trabajo", "observaciones_trabajo"),
    ("observaciones adicionales", "observaciones_adicionales"),
    ("despues de la aprobacion", "despues_aprobacion"),
    ("despues del rechazo", "despues_rechazo"),
    ("servicio de negocio", "servicio_negocio"),
    ("tipo de falla", "tipo_falla"),
    ("impacto", "impacto")
]


CAMPOS_CAUSA_INCIDENTE = [
    ("descripcion", "descripcion"),
    ("breve descripcion", "breve_descripcion"),
    ("tipo de falla", "tipo_falla"),
    ("observaciones y notas de trabajo", "observaciones_trabajo"),
    ("observaciones adicionales", "observaciones_adicionales"),
    ("despues de la aprobacion", "despues_aprobacion"),
    ("despues del rechazo", "despues_rechazo"),
    ("lista de notas de trabajo", "lista_notas_trabajo"),
    ("actualizaciones", "actualizaciones")
]


CAMPOS_ALERTA_INCIDENTE = CAMPOS_CAUSA_INCIDENTE + [
    ("grupo de asignacion", "grupo_asignacion"),
    ("servicio de negocio", "servicio_negocio"),
    ("impacto", "impacto"),
    ("estado", "estado")
]


PALABRAS_CLIENTE_EXTERNO = [
    "firma", "firmar", "cliente", "usuario", "certificado",
    "certificados", "token", "portal", "ocsp", "tsa",
    "clave segura", "certitoken"
]


PALABRAS_CLIENTE_INTERNO = [
    "ldap", "directorio activo", "active directory", "servidor",
    "infraestructura", "red interna", "vpn", "base de datos",
    "bd", "interno", "correo interno", "aplicativo interno"
]


PALABRAS_ALERTA = [
    "alerta", "alarma", "monitoreo", "monitor", "zabbix",
    "grafana", "prometheus", "observabilidad", "cpu", "memoria",
    "disco", "latencia", "packet loss", "perdida de paquetes",
    "indisponibilidad", "caida", "caido", "down", "critico",
    "critical", "warning"
]


def unir_campos(row, grupos_claves):
    partes = []

    for claves in grupos_claves:
        texto = normalizar_texto(valor_campo(row, *claves))
        if texto:
            partes.append(texto)

    return " ".join(partes)


def contiene_alguna(texto, palabras):
    return any(palabra in texto for palabra in palabras)


def tipificacion_incidente(row):
    categoria = normalizar_texto(valor_campo(row, "categoria"))
    texto = unir_campos(row, CAMPOS_TIPIFICACION_INCIDENTE)

    if "consulta" in categoria:
        return "Consulta"

    if "solicitud" in categoria:
        return "Solicitud"

    if "noc" in texto:
        return "NOC"

    if "incidente" in categoria:
        if contiene_alguna(texto, PALABRAS_CLIENTE_EXTERNO):
            return "Cliente Externo"
        if contiene_alguna(texto, PALABRAS_CLIENTE_INTERNO):
            return "Cliente Interno"
        return "Cliente Interno"

    if contiene_alguna(texto, PALABRAS_CLIENTE_EXTERNO):
        return "Cliente Externo"

    if contiene_alguna(texto, PALABRAS_CLIENTE_INTERNO):
        return "Cliente Interno"

    return "Cliente Interno"


def es_alerta_incidente(row, tipificacion=None):
    if tipificacion is None:
        tipificacion = tipificacion_incidente(row)

    texto = unir_campos(row, CAMPOS_ALERTA_INCIDENTE)

    if tipificacion == "NOC":
        return "Si"

    if contiene_alguna(texto, PALABRAS_ALERTA):
        return "Si"

    return "No"


def causa_raiz_incidente(row, tipificacion=None, es_alerta=None):
    if tipificacion is None:
        tipificacion = tipificacion_incidente(row)

    if es_alerta is None:
        es_alerta = es_alerta_incidente(row, tipificacion)

    texto = unir_campos(row, CAMPOS_CAUSA_INCIDENTE)

    if "ocsp" in texto:
        detalle = "Servicios OCSP con indisponibilidad o validacion intermitente"
    elif "certitoken" in texto or "token" in texto:
        detalle = "Servicios Certitoken con falla de autenticacion, firma o uso del token"
    elif "clave segura" in texto:
        detalle = "Clave Segura con falla de acceso o disponibilidad"
    elif contiene_alguna(texto, ["ldap", "directorio activo", "active directory"]):
        detalle = "LDAP o Directorio Activo con falla de autenticacion"
    elif contiene_alguna(texto, ["base de datos", "bd", "database", "sql"]):
        detalle = "Base de datos o integracion interna con errores"
    elif contiene_alguna(texto, ["correo", "smtp", "mail"]):
        detalle = "Servicio de correo o notificaciones con fallas"
    elif contiene_alguna(texto, ["ssl", "certificado ssl", "certificado vencido", "cadena de certificados"]):
        detalle = "Certificado SSL o cadena de certificados con errores"
    elif contiene_alguna(texto, ["tsa", "timestamp", "sello de tiempo"]):
        detalle = "Servicios TSA con falla de sellado de tiempo"
    elif "portal" in texto and "rpost" in texto:
        detalle = "Portal RPOST con falla funcional o de acceso"
    elif contiene_alguna(texto, ["cpu", "memoria", "disco", "filesystem", "fs full", "espacio", "swap"]):
        detalle = "Infraestructura con consumo alto de recursos"
    elif contiene_alguna(texto, ["red", "conexion", "conectividad", "vpn", "packet loss", "perdida de paquetes", "latencia", "enlace"]):
        detalle = "Infraestructura de red o conectividad degradada"
    elif contiene_alguna(texto, ["caida", "caido", "indisponibilidad", "no responde", "fuera de servicio", "down"]):
        detalle = "Caida o indisponibilidad del servicio"
    elif contiene_alguna(texto, ["lentitud", "degradacion", "degradado", "timeout", "time out", "intermitencia", "intermitente"]):
        detalle = "Degradacion o intermitencia del servicio"
    elif contiene_alguna(texto, ["firma", "firmar", "rechazo de firma"]):
        detalle = "Proceso de firma con error o rechazo"
    elif contiene_alguna(texto, ["proveedor", "tercero", "externo"]):
        detalle = "Dependencia de proveedor o tercero"
    elif contiene_alguna(texto, ["alerta", "alarma", "monitoreo", "monitor", "zabbix", "grafana", "prometheus", "observabilidad"]):
        detalle = "Monitoreo detecta comportamiento anomalo"
    else:
        detalle = "Analisis pendiente para precisar la causa raiz"

    if tipificacion == "NOC":
        return f"Alerta NOC - {detalle}"

    if es_alerta == "Si" and detalle == "Analisis pendiente para precisar la causa raiz":
        return f"Alerta operativa - {detalle}"

    return detalle


def clasificacion_incidente(row):
    tipificacion = tipificacion_incidente(row)
    es_alerta = es_alerta_incidente(row, tipificacion)
    causa_raiz = causa_raiz_incidente(row, tipificacion, es_alerta)
    return tipificacion, es_alerta, causa_raiz


def guardar_incidentes(df):
    conn = get_conn()
    nuevos, actualizados = 0, 0

    df = limpiar_incidentes(df)

    for _, r in df.iterrows():
        numero = str(r.get("Número"))
        if not numero or numero == "nan":
            continue

        duracion_segundos = r.get("Duración segundos")
        duracion_horas = segundos_a_horas(duracion_segundos)
        tipificacion_auto, es_alerta_auto, causa_raiz_auto = clasificacion_incidente(r)

        data = (
            numero,
            str(r.get("Solicitante")),
            str(r.get("Categoria")),
            str(r.get("Prioridad")),
            str(r.get("Estado")),
            str(r.get("Grupo de asignación")),
            str(r.get("Asignado a")),
            str(r.get("Descripción")),
            str(r.get("Después de la aprobación")),
            str(r.get("Después del rechazo")),
            None if pd.isna(duracion_segundos) else float(duracion_segundos),
            None if pd.isna(r.get("Fecha de vencimiento del SLA")) else str(r.get("Fecha de vencimiento del SLA")),
            str(r.get("Tipo de falla")),
            str(r.get("Empresa")),
            str(r.get("Creado por")),
            None if pd.isna(r.get("Cerrado")) else str(r.get("Cerrado")),
            str(r.get("Escalado a proveedor")),
            str(r.get("Servicio de Negocio")),
            None if pd.isna(r.get("Creado")) else str(r.get("Creado")),
            str(r.get("Observaciones y notas de trabajo")),
            str(r.get("Observaciones adicionales")),
            str(r.get("Actualizaciones")),
            str(r.get("Impacto")),
            str(r.get("Lista de notas de trabajo")),
            str(r.get("Breve descripción")),
            None if pd.isna(r.get("Minutos")) else float(r.get("Minutos")),
            tipificacion_auto,
            causa_raiz_auto,
            es_alerta_auto,
            duracion_horas
        )
        incident_columns = (
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
            "breve_descripcion",
            "minutos",
            "tipificacion_auto",
            "causa_raiz_auto",
            "es_alerta_auto",
            "duracion_horas"
        )

        cur = conn.cursor()
        cur.execute("SELECT numero FROM incidents WHERE numero=?", (numero,))
        if cur.fetchone():
            set_clause = ", ".join(f"{col}=?" for col in incident_columns if col != "numero")
            cur.execute(
                f"UPDATE incidents SET {set_clause} WHERE numero=?",
                data[1:] + (numero,)
            )
            actualizados += 1
        else:
            columns_sql = ", ".join(incident_columns)
            placeholders = ", ".join("?" for _ in incident_columns)
            cur.execute(
                f"INSERT INTO incidents ({columns_sql}) VALUES ({placeholders})",
                data
            )
            nuevos += 1

    conn.commit()
    conn.close()
    return nuevos, actualizados


def load_incidentes():
    conn = get_conn()
    df = pd.read_sql("SELECT * FROM incidents", conn)
    conn.close()

    if not df.empty:
        clasificaciones = df.apply(clasificacion_incidente, axis=1, result_type="expand")
        clasificaciones.columns = ["tipificacion_auto", "es_alerta_auto", "causa_raiz_auto"]
        df[["tipificacion_auto", "es_alerta_auto", "causa_raiz_auto"]] = clasificaciones

    return df


def dashboard_casos():
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
    else:
        st.info("No hay datos de casos cargados.")


def dashboard_incidentes():
    df = load_incidentes()

    if not df.empty:
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
            fig.update_layout(showlegend=False, title="Incidentes por tipificación")
            st.plotly_chart(fig, use_container_width=True)

        with fila1_col2:
            causas = df["causa_raiz_auto"].fillna("Otros").value_counts().reset_index()
            causas.columns = ["Causa raíz", "Cantidad"]
            causas = causas.sort_values(by="Cantidad", ascending=True)

            fig2 = px.bar(
                causas,
                x="Cantidad",
                y="Causa raíz",
                orientation="h",
                text="Cantidad",
                color_discrete_sequence=["#1e6f5c"]
            )
            fig2.update_traces(textposition="outside")
            fig2.update_layout(title="Incidentes por causa raíz")
            st.plotly_chart(fig2, use_container_width=True)

        fila2_col1, fila2_col2 = st.columns(2)

        with fila2_col1:
            servicios = df["servicio_negocio"].fillna("Sin servicio").value_counts().reset_index()
            servicios.columns = ["Servicio", "Cantidad"]
            servicios = servicios.sort_values(by="Cantidad", ascending=True)

            fig3 = px.bar(
                servicios,
                x="Cantidad",
                y="Servicio",
                orientation="h",
                text="Cantidad",
                color_discrete_sequence=["#759d2f"]
            )
            fig3.update_traces(textposition="outside")
            fig3.update_layout(title="Servicios afectados")
            st.plotly_chart(fig3, use_container_width=True)

        with fila2_col2:
            df["creado"] = pd.to_datetime(df["creado"], errors="coerce")
            incidentes_dia = df.groupby(df["creado"].dt.date).size().reset_index(name="incidentes")

            fig4 = px.bar(
                incidentes_dia,
                x="creado",
                y="incidentes",
                title="Incidentes por día"
            )
            fig4.update_traces(marker_color="#f76c02")
            st.plotly_chart(fig4, use_container_width=True)

        st.divider()
        st.subheader("Impacto en Cliente Externo")

        df_cliente_externo = df[
            df["tipificacion_auto"].fillna("Cliente Interno") == "Cliente Externo"
        ].copy()

        if not df_cliente_externo.empty:
            causas_externo = (
                df_cliente_externo["causa_raiz_auto"]
                .fillna("Otros")
                .value_counts()
                .reset_index()
            )
            causas_externo.columns = ["Afectacion", "Cantidad"]
            causas_externo = causas_externo.sort_values(by="Cantidad", ascending=True)

            principal_afectacion = causas_externo.iloc[-1]["Afectacion"]
            porcentaje_externo = round((len(df_cliente_externo) / total) * 100, 2) if total > 0 else 0

            externo_col1, externo_col2 = st.columns([1, 2])

            with externo_col1:
                st.markdown(
                    tarjeta("Incidentes Cliente Externo", len(df_cliente_externo)),
                    unsafe_allow_html=True
                )
                st.caption(f"Participacion sobre el total: {porcentaje_externo}%")
                st.caption(f"Principal afectacion: {principal_afectacion}")

            with externo_col2:
                fig5 = px.bar(
                    causas_externo,
                    x="Cantidad",
                    y="Afectacion",
                    orientation="h",
                    text="Cantidad",
                    color_discrete_sequence=["#c2410c"],
                    title="Que esta afectando al Cliente Externo"
                )
                fig5.update_traces(textposition="outside")
                st.plotly_chart(fig5, use_container_width=True)
        else:
            st.info("No hay incidentes tipificados como Cliente Externo en los datos cargados.")

        alertas = []
        if len(df[df["causa_raiz_auto"].fillna("").str.contains("Servicios OCSP", case=False)]) >= 3:
            alertas.append("Alta recurrencia asociada a servicios OCSP.")
        if len(df[df["causa_raiz_auto"].fillna("").str.contains("Servicios Certitoken", case=False)]) >= 3:
            alertas.append("Alta recurrencia asociada a servicios Certitoken.")
        if porcentaje_sla < 90:
            alertas.append("El cumplimiento de SLA de incidentes está por debajo del 90%.")
        if len(df[df["servicio_negocio"] == "Certificación Digital"]) >= 5:
            alertas.append("El servicio de negocio Certificación Digital presenta concentración relevante de incidentes.")

        if alertas_tipificadas >= 5:
            alertas.append("Se observa un volumen alto de incidentes tipificados como alerta.")

        if alertas:
            st.divider()
            st.subheader("Alertas")
            for alerta in alertas:
                st.warning(alerta)
    else:
        st.info("No hay datos de incidentes cargados.")

APP_PALETTE = {
    "orange_dark": "#e28000",
    "orange": "#ff9800",
    "yellow": "#ffc340",
    "red_dark": "#e00000",
    "red": "#ff0000",
    "white": "#ffffff",
    "ink": "#5f2300",
}

APP_PALETTE_SEQUENCE = [
    APP_PALETTE["orange_dark"],
    APP_PALETTE["orange"],
    APP_PALETTE["yellow"],
    APP_PALETTE["red_dark"],
    APP_PALETTE["red"],
]

CASE_ALIASES_APP = {
    "numero": ["numero", "número"],
    "descripcion": ["breve descripcion", "breve descripción", "descripcion corta"],
    "contacto": ["contacto"],
    "cuenta": ["cuenta"],
    "codigo_resolucion": ["codigo de resolucion", "código de resolución"],
    "canal": ["canal"],
    "estado": ["estado"],
    "prioridad": ["prioridad"],
    "asignado": ["asignado a"],
    "actualizado": ["actualizado"],
    "creado_por": ["creado por"],
    "creado": ["creado"],
    "producto": ["producto"],
    "cerrado": ["cerrado"],
    "causa": ["causa"],
    "notas_resolucion": ["notas de resolucion", "notas de resolución"],
    "observaciones_adicionales": ["observaciones adicionales"],
    "observaciones_trabajo": ["observaciones y notas de trabajo", "observaciones de trabajo"],
}

INCIDENT_ALIASES_APP = {
    "numero": ["numero", "número"],
    "solicitante": ["solicitante"],
    "breve_descripcion": ["breve descripcion", "breve descripción"],
    "categoria": ["categoria", "categoría"],
    "prioridad": ["prioridad"],
    "estado": ["estado"],
    "grupo_asignacion": ["grupo de asignacion", "grupo de asignación"],
    "asignado_a": ["asignado a"],
    "descripcion": ["descripcion", "descripción"],
    "despues_aprobacion": ["despues de la aprobacion", "después de la aprobación"],
    "despues_rechazo": ["despues del rechazo", "después del rechazo"],
    "duracion_segundos": ["duracion segundos", "duración segundos"],
    "duracion": ["duracion", "duración"],
    "minutos": ["minutos"],
    "fecha_vencimiento_sla": ["fecha de vencimiento del sla", "vencimiento sla"],
    "tipo_falla": ["tipo de falla"],
    "empresa": ["empresa"],
    "creado_por": ["creado por"],
    "cerrado": ["cerrado"],
    "escalado_proveedor": ["escalado a proveedor"],
    "servicio_negocio": ["servicio de negocio"],
    "creado": ["creado"],
    "observaciones_trabajo": ["observaciones y notas de trabajo", "observaciones de trabajo"],
    "observaciones_adicionales": ["observaciones adicionales"],
    "actualizaciones": ["actualizaciones"],
    "impacto": ["impacto"],
    "lista_notas_trabajo": ["lista de notas de trabajo"],
}

INCIDENT_TEXT_FIELDS_APP = [
    "grupo_asignacion",
    "asignado_a",
    "descripcion",
    "breve_descripcion",
    "observaciones_trabajo",
    "observaciones_adicionales",
    "despues_aprobacion",
    "despues_rechazo",
    "servicio_negocio",
    "tipo_falla",
    "impacto",
]

INCIDENT_CAUSE_FIELDS_APP = [
    "descripcion",
    "breve_descripcion",
    "tipo_falla",
    "observaciones_trabajo",
    "observaciones_adicionales",
    "despues_aprobacion",
    "despues_rechazo",
    "lista_notas_trabajo",
    "actualizaciones",
]

EXTERNAL_KEYWORDS_APP = [
    "firma", "firmar", "cliente", "usuario", "certificado", "certificados",
    "token", "portal", "ocsp", "tsa", "clave segura", "certitoken",
]

INTERNAL_KEYWORDS_APP = [
    "ldap", "directorio activo", "active directory", "servidor", "infraestructura",
    "red interna", "vpn", "base de datos", "bd", "interno", "correo interno",
    "aplicativo interno",
]

ALERT_KEYWORDS_APP = [
    "alerta", "alarma", "monitoreo", "monitor", "zabbix", "grafana",
    "prometheus", "observabilidad", "cpu", "memoria", "disco", "latencia",
    "packet loss", "perdida de paquetes", "indisponibilidad", "caida", "caido",
    "down", "critico", "critical", "warning",
]


def normalizar_clave_app(valor):
    texto = unicodedata.normalize("NFKD", str(valor)).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", texto.strip().lower())


def safe_text_app(valor):
    if valor is None or pd.isna(valor):
        return ""
    return str(valor).strip()


def safe_float_app(valor):
    try:
        if valor is None or pd.isna(valor):
            return None
        return float(valor)
    except Exception:
        return None


def renombrar_columnas_app(df, aliases):
    df = df.copy()
    columnas = {normalizar_clave_app(col): col for col in df.columns}
    renames = {}
    for destino, opciones in aliases.items():
        for opcion in opciones:
            original = columnas.get(normalizar_clave_app(opcion))
            if original is not None:
                renames[original] = destino
                break
    df = df.rename(columns=renames)
    df = df.loc[:, ~df.columns.duplicated()]
    df = df.loc[:, ~df.columns.astype(str).str.contains("^Unnamed", na=False)]
    return df


def valor_fila_app(row, clave, default=""):
    valor = row.get(clave, default)
    if valor is None or pd.isna(valor):
        return default
    return valor


def unir_textos_app(row, campos):
    return " ".join(normalizar_texto(valor_fila_app(row, campo)) for campo in campos).strip()


def tipificar(row):
    texto = " ".join(
        [
            normalizar_texto(valor_fila_app(row, "descripcion")),
            normalizar_texto(valor_fila_app(row, "causa")),
            normalizar_texto(valor_fila_app(row, "notas_resolucion")),
            normalizar_texto(valor_fila_app(row, "observaciones_adicionales")),
            normalizar_texto(valor_fila_app(row, "observaciones_trabajo")),
        ]
    )
    if "phishing" in texto:
        return "1 - phishing"
    if "error" in texto or "falla" in texto:
        return "2 - Soporte Falla"
    if "contrasena" in texto or "como usar" in texto:
        return "3 - Soporte Uso"
    if any(p in texto for p in ["cert", "pagar", "biometria"]):
        return "4 - solicitudes"
    if "incidente" in texto:
        return "5 - incidente"
    if "externo" in texto:
        return "6 - Plataformas Ext"
    return "7 - No Aplica"


def tiempo(creado, cerrado):
    try:
        if cerrado in ("", None) or pd.isna(cerrado):
            return "En proceso"
        inicio = pd.to_datetime(creado, errors="coerce")
        fin = pd.to_datetime(cerrado, errors="coerce")
        if pd.isna(inicio) or pd.isna(fin):
            return "Error"
        total = timedelta()
        actual = inicio
        while actual < fin:
            if actual.weekday() >= 5:
                actual = (actual + timedelta(days=1)).replace(hour=8, minute=0, second=0, microsecond=0)
                continue
            inicio_dia = actual.replace(hour=8, minute=0, second=0, microsecond=0)
            fin_dia = actual.replace(hour=17, minute=0, second=0, microsecond=0)
            if actual < inicio_dia:
                actual = inicio_dia
            if actual >= fin_dia:
                actual = (actual + timedelta(days=1)).replace(hour=8, minute=0, second=0, microsecond=0)
                continue
            limite = min(fin, fin_dia)
            total += limite - actual
            actual = limite
        return round(total.total_seconds() / 3600, 2)
    except Exception:
        return "Error"


def preparar_casos_app(df):
    df = renombrar_columnas_app(df, CASE_ALIASES_APP)
    for columna in CASE_ALIASES_APP:
        if columna not in df.columns:
            df[columna] = None
    df = df.replace({pd.NA: None})
    df["numero"] = df["numero"].apply(safe_text_app)
    df = df[df["numero"] != ""]
    return df.drop_duplicates(subset=["numero"], keep="last")


def guardar(df):
    conn = get_conn()
    nuevos, actualizados = 0, 0
    df = preparar_casos_app(df)
    for _, row in df.iterrows():
        numero = safe_text_app(valor_fila_app(row, "numero"))
        data = (
            numero,
            safe_text_app(valor_fila_app(row, "descripcion")),
            safe_text_app(valor_fila_app(row, "contacto")),
            safe_text_app(valor_fila_app(row, "cuenta")),
            safe_text_app(valor_fila_app(row, "codigo_resolucion")),
            safe_text_app(valor_fila_app(row, "canal")),
            safe_text_app(valor_fila_app(row, "estado")),
            safe_text_app(valor_fila_app(row, "prioridad")),
            safe_text_app(valor_fila_app(row, "asignado")),
            safe_text_app(valor_fila_app(row, "actualizado")),
            safe_text_app(valor_fila_app(row, "creado_por")),
            safe_text_app(valor_fila_app(row, "creado")),
            safe_text_app(valor_fila_app(row, "producto")),
            safe_text_app(valor_fila_app(row, "cerrado")),
            safe_text_app(valor_fila_app(row, "causa")),
            safe_text_app(valor_fila_app(row, "notas_resolucion")),
            safe_text_app(valor_fila_app(row, "observaciones_adicionales")),
            safe_text_app(valor_fila_app(row, "observaciones_trabajo")),
            tipificar(row),
            tiempo(valor_fila_app(row, "creado"), valor_fila_app(row, "cerrado")),
        )
        cur = conn.cursor()
        cur.execute("SELECT numero FROM cases WHERE numero=?", (numero,))
        if cur.fetchone():
            cur.execute(
                """
                UPDATE cases SET descripcion=?, contacto=?, cuenta=?, codigo_resolucion=?, canal=?,
                estado=?, prioridad=?, asignado=?, actualizado=?, creado_por=?, creado=?, producto=?,
                cerrado=?, causa=?, notas_resolucion=?, observaciones_adicionales=?,
                observaciones_trabajo=?, tipificacion=?, tiempo_respuesta=? WHERE numero=?
                """,
                data[1:] + (numero,),
            )
            actualizados += 1
        else:
            cur.execute("INSERT INTO cases VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", data)
            nuevos += 1
    conn.commit()
    conn.close()
    return nuevos, actualizados


def load():
    conn = get_conn()
    df = pd.read_sql("SELECT * FROM cases", conn)
    conn.close()
    return df

def tipificacion_incidente(row):
    categoria = normalizar_texto(valor_fila_app(row, "categoria"))
    texto = unir_textos_app(row, INCIDENT_TEXT_FIELDS_APP)
    if "consulta" in categoria:
        return "Consulta"
    if "solicitud" in categoria:
        return "Solicitud"
    if "noc" in texto:
        return "NOC"
    if "incidente" in categoria:
        if any(p in texto for p in EXTERNAL_KEYWORDS_APP):
            return "Cliente Externo"
        if any(p in texto for p in INTERNAL_KEYWORDS_APP):
            return "Cliente Interno"
        return "Cliente Interno"
    if any(p in texto for p in EXTERNAL_KEYWORDS_APP):
        return "Cliente Externo"
    if any(p in texto for p in INTERNAL_KEYWORDS_APP):
        return "Cliente Interno"
    return "Cliente Interno"


def es_alerta_incidente(row, tipificacion=None):
    tipificacion = tipificacion or tipificacion_incidente(row)
    texto = unir_textos_app(row, INCIDENT_CAUSE_FIELDS_APP + ["grupo_asignacion", "servicio_negocio", "impacto", "estado"])
    if tipificacion == "NOC":
        return "Si"
    if any(p in texto for p in ALERT_KEYWORDS_APP):
        return "Si"
    return "No"


def causa_raiz_incidente(row, tipificacion=None, es_alerta=None):
    tipificacion = tipificacion or tipificacion_incidente(row)
    es_alerta = es_alerta or es_alerta_incidente(row, tipificacion)
    texto = unir_textos_app(row, INCIDENT_CAUSE_FIELDS_APP)
    if "ocsp" in texto:
        detalle = "Servicios OCSP con indisponibilidad o validacion intermitente"
    elif "certitoken" in texto or "token" in texto:
        detalle = "Servicios Certitoken con falla de autenticacion, firma o uso del token"
    elif "clave segura" in texto:
        detalle = "Clave Segura con falla de acceso o disponibilidad"
    elif any(p in texto for p in ["ldap", "directorio activo", "active directory"]):
        detalle = "LDAP o Directorio Activo con falla de autenticacion"
    elif any(p in texto for p in ["base de datos", "bd", "database", "sql"]):
        detalle = "Base de datos o integracion interna con errores"
    elif any(p in texto for p in ["correo", "smtp", "mail"]):
        detalle = "Servicio de correo o notificaciones con fallas"
    elif any(p in texto for p in ["ssl", "certificado ssl", "certificado vencido", "cadena de certificados"]):
        detalle = "Certificado SSL o cadena de certificados con errores"
    elif any(p in texto for p in ["tsa", "timestamp", "sello de tiempo"]):
        detalle = "Servicios TSA con falla de sellado de tiempo"
    elif "portal" in texto and "rpost" in texto:
        detalle = "Portal RPOST con falla funcional o de acceso"
    elif any(p in texto for p in ["cpu", "memoria", "disco", "filesystem", "espacio", "swap"]):
        detalle = "Infraestructura con consumo alto de recursos"
    elif any(p in texto for p in ["red", "conexion", "conectividad", "vpn", "packet loss", "perdida de paquetes", "latencia", "enlace"]):
        detalle = "Infraestructura de red o conectividad degradada"
    elif any(p in texto for p in ["caida", "caido", "indisponibilidad", "no responde", "fuera de servicio", "down"]):
        detalle = "Caida o indisponibilidad del servicio"
    elif any(p in texto for p in ["lentitud", "degradacion", "timeout", "time out", "intermitencia", "intermitente"]):
        detalle = "Degradacion o intermitencia del servicio"
    elif any(p in texto for p in ["firma", "firmar", "rechazo de firma"]):
        detalle = "Proceso de firma con error o rechazo"
    elif any(p in texto for p in ["proveedor", "tercero", "externo"]):
        detalle = "Dependencia de proveedor o tercero"
    elif any(p in texto for p in ["alerta", "alarma", "monitoreo", "monitor", "zabbix", "grafana", "prometheus", "observabilidad"]):
        detalle = "Monitoreo detecta comportamiento anomalo"
    else:
        detalle = "Analisis pendiente para precisar la causa raiz"
    if tipificacion == "NOC":
        return f"Alerta NOC - {detalle}"
    if es_alerta == "Si" and detalle == "Analisis pendiente para precisar la causa raiz":
        return f"Alerta operativa - {detalle}"
    return detalle


def clasificacion_incidente(row):
    tipificacion = tipificacion_incidente(row)
    alerta = es_alerta_incidente(row, tipificacion)
    causa = causa_raiz_incidente(row, tipificacion, alerta)
    return tipificacion, alerta, causa


def preparar_incidentes_app(df):
    df = renombrar_columnas_app(df, INCIDENT_ALIASES_APP)
    for columna in list(INCIDENT_ALIASES_APP.keys()) + ["tipificacion_original", "causa_raiz_original"]:
        if columna not in df.columns:
            df[columna] = None
    if df["duracion_segundos"].isna().all() and "duracion" in df.columns:
        df["duracion_segundos"] = df["duracion"]
    df = df.replace({pd.NA: None})
    df["numero"] = df["numero"].apply(safe_text_app)
    df = df[df["numero"] != ""]
    return df.drop_duplicates(subset=["numero"], keep="last")


def guardar_incidentes(df):
    conn = get_conn()
    nuevos, actualizados = 0, 0
    df = preparar_incidentes_app(df)
    for _, row in df.iterrows():
        numero = safe_text_app(valor_fila_app(row, "numero"))
        tipificacion_auto, es_alerta_auto, causa_raiz_auto = clasificacion_incidente(row)
        duracion_segundos = safe_float_app(valor_fila_app(row, "duracion_segundos"))
        duracion_horas = segundos_a_horas(duracion_segundos)
        data = (
            numero,
            safe_text_app(valor_fila_app(row, "solicitante")),
            safe_text_app(valor_fila_app(row, "breve_descripcion")),
            safe_text_app(valor_fila_app(row, "categoria")),
            safe_text_app(valor_fila_app(row, "prioridad")),
            safe_text_app(valor_fila_app(row, "estado")),
            safe_text_app(valor_fila_app(row, "grupo_asignacion")),
            safe_text_app(valor_fila_app(row, "asignado_a")),
            safe_text_app(valor_fila_app(row, "descripcion")),
            safe_text_app(valor_fila_app(row, "despues_aprobacion")),
            safe_text_app(valor_fila_app(row, "despues_rechazo")),
            duracion_segundos,
            safe_float_app(valor_fila_app(row, "minutos")),
            safe_text_app(valor_fila_app(row, "fecha_vencimiento_sla")),
            safe_text_app(valor_fila_app(row, "tipo_falla")),
            safe_text_app(valor_fila_app(row, "empresa")),
            safe_text_app(valor_fila_app(row, "creado_por")),
            safe_text_app(valor_fila_app(row, "cerrado")),
            safe_text_app(valor_fila_app(row, "escalado_proveedor")),
            safe_text_app(valor_fila_app(row, "servicio_negocio")),
            safe_text_app(valor_fila_app(row, "creado")),
            safe_text_app(valor_fila_app(row, "observaciones_trabajo")),
            safe_text_app(valor_fila_app(row, "observaciones_adicionales")),
            safe_text_app(valor_fila_app(row, "actualizaciones")),
            safe_text_app(valor_fila_app(row, "impacto")),
            safe_text_app(valor_fila_app(row, "lista_notas_trabajo")),
            safe_text_app(valor_fila_app(row, "tipificacion_original")),
            safe_text_app(valor_fila_app(row, "causa_raiz_original")),
            tipificacion_auto,
            causa_raiz_auto,
            es_alerta_auto,
            duracion_horas,
        )
        cur = conn.cursor()
        cur.execute("SELECT numero FROM incidents WHERE numero=?", (numero,))
        if cur.fetchone():
            cur.execute(
                """
                UPDATE incidents SET solicitante=?, breve_descripcion=?, categoria=?, prioridad=?, estado=?,
                grupo_asignacion=?, asignado_a=?, descripcion=?, despues_aprobacion=?, despues_rechazo=?,
                duracion_segundos=?, minutos=?, fecha_vencimiento_sla=?, tipo_falla=?, empresa=?, creado_por=?,
                cerrado=?, escalado_proveedor=?, servicio_negocio=?, creado=?, observaciones_trabajo=?,
                observaciones_adicionales=?, actualizaciones=?, impacto=?, lista_notas_trabajo=?,
                tipificacion_original=?, causa_raiz_original=?, tipificacion_auto=?, causa_raiz_auto=?,
                es_alerta_auto=?, duracion_horas=? WHERE numero=?
                """,
                data[1:] + (numero,),
            )
            actualizados += 1
        else:
            cur.execute(
                """
                INSERT INTO incidents (
                    numero, solicitante, breve_descripcion, categoria, prioridad, estado,
                    grupo_asignacion, asignado_a, descripcion, despues_aprobacion, despues_rechazo,
                    duracion_segundos, minutos, fecha_vencimiento_sla, tipo_falla, empresa, creado_por,
                    cerrado, escalado_proveedor, servicio_negocio, creado, observaciones_trabajo,
                    observaciones_adicionales, actualizaciones, impacto, lista_notas_trabajo,
                    tipificacion_original, causa_raiz_original, tipificacion_auto, causa_raiz_auto,
                    es_alerta_auto, duracion_horas
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                data,
            )
            nuevos += 1
    conn.commit()
    conn.close()
    return nuevos, actualizados


def load_incidentes():
    conn = get_conn()
    df = pd.read_sql("SELECT * FROM incidents", conn)
    conn.close()
    if not df.empty:
        clasificaciones = df.apply(clasificacion_incidente, axis=1, result_type="expand")
        clasificaciones.columns = ["tipificacion_auto", "es_alerta_auto", "causa_raiz_auto"]
        df[["tipificacion_auto", "es_alerta_auto", "causa_raiz_auto"]] = clasificaciones
    return df

from app_ui import run_app

run_app()
st.stop()

init_db()

if not login():
    st.stop()

st.title("Control de casos e incidentes")

if st.session_state.role == "admin":
    menu = st.sidebar.selectbox(
        "Menú",
        [
            "Cargar Excel Casos",
            "Casos",
            "Dashboard Casos",
            "Cargar Excel Incidentes",
            "Incidentes",
            "Dashboard Incidentes"
        ]
    )
    if st.sidebar.button("Cerrar sesión"):
        st.session_state.clear()
        st.rerun()
else:
    if st.button("Cerrar sesión"):
        st.session_state.clear()
        st.rerun()

    tabs = st.tabs(["Dashboard Casos", "Dashboard Incidentes"])
    with tabs[0]:
        dashboard_casos()
    with tabs[1]:
        dashboard_incidentes()
    st.stop()

if menu == "Cargar Excel Casos":
    file = st.file_uploader("Sube Excel de casos", type=["xlsx"], key="casos_upload")
    if file:
        df = pd.read_excel(file)
        st.write(f"Filas detectadas: {len(df)}")
        st.dataframe(df.head())

        if st.button("Procesar casos"):
            n, a = guardar(df)
            st.success(f"Nuevos: {n} | Actualizados: {a}")

elif menu == "Casos":
    df = load()

    if not df.empty:
        filtro_estado = st.selectbox("Estado", ["Todos"] + list(df["estado"].dropna().unique()), key="estado_casos")
        filtro_cuenta = st.text_input("Cuenta", key="cuenta_casos")

        if filtro_estado != "Todos":
            df = df[df["estado"] == filtro_estado]

        if filtro_cuenta:
            df = df[df["cuenta"].str.contains(filtro_cuenta, case=False, na=False)]

    st.dataframe(df, use_container_width=True)

elif menu == "Dashboard Casos":
    dashboard_casos()

elif menu == "Cargar Excel Incidentes":
    file = st.file_uploader("Sube Excel de incidentes", type=["xlsx"], key="incidentes_upload")
    if file:
        df = pd.read_excel(file)
        st.write(f"Filas detectadas: {len(df)}")
        st.dataframe(df.head())

        if st.button("Procesar incidentes"):
            n, a = guardar_incidentes(df)
            st.success(f"Nuevos: {n} | Actualizados: {a}")

elif menu == "Incidentes":
    df = load_incidentes()

    if not df.empty:
        filtro_estado = st.selectbox("Estado", ["Todos"] + list(df["estado"].dropna().unique()), key="estado_inc")
        filtro_tipificacion = st.selectbox(
            "Tipificación",
            ["Todos"] + sorted(df["tipificacion_auto"].dropna().unique().tolist()),
            key="tip_inc"
        )
        filtro_alerta = st.selectbox(
            "Es alerta",
            ["Todos"] + sorted(df["es_alerta_auto"].dropna().unique().tolist()),
            key="alerta_inc"
        )

        if filtro_estado != "Todos":
            df = df[df["estado"] == filtro_estado]

        if filtro_tipificacion != "Todos":
            df = df[df["tipificacion_auto"] == filtro_tipificacion]

        if filtro_alerta != "Todos":
            df = df[df["es_alerta_auto"] == filtro_alerta]

        columnas_visibles = [
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
            "duracion_horas"
        ]
        columnas_visibles = [c for c in columnas_visibles if c in df.columns]
        df = df[columnas_visibles]

    st.dataframe(df, use_container_width=True)

elif menu == "Dashboard Incidentes":
    dashboard_incidentes()
