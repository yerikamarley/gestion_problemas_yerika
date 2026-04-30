import re
import hashlib
import hmac
import os
import sqlite3
import unicodedata
from datetime import timedelta
from math import ceil

import pandas as pd


DB = "data.db"
ADMIN_EMAIL = "yerika.basto@certicamara.com"
INITIAL_ADMIN_PASSWORD = os.environ.get("APP_ADMIN_PASSWORD", "Yerika2026!")

CASE_ALIASES = {
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
    "observaciones_trabajo": [
        "observaciones y notas de trabajo",
        "observaciones de trabajo",
    ],
}

INCIDENT_ALIASES = {
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
    "fecha_vencimiento_sla": [
        "fecha de vencimiento del sla",
        "vencimiento sla",
    ],
    "tipo_falla": ["tipo de falla"],
    "empresa": ["empresa"],
    "creado_por": ["creado por"],
    "cerrado": ["cerrado"],
    "escalado_proveedor": ["escalado a proveedor"],
    "servicio_negocio": ["servicio de negocio"],
    "creado": ["creado"],
    "observaciones_trabajo": [
        "observaciones y notas de trabajo",
        "observaciones de trabajo",
    ],
    "observaciones_adicionales": ["observaciones adicionales"],
    "actualizaciones": ["actualizaciones"],
    "impacto": ["impacto"],
    "lista_notas_trabajo": ["lista de notas de trabajo"],
}

INCIDENT_TEXT_FIELDS = [
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

INCIDENT_CAUSE_FIELDS = [
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

EXTERNAL_KEYWORDS = [
    "firma",
    "firmar",
    "cliente",
    "usuario",
    "certificado",
    "certificados",
    "token",
    "portal",
    "ocsp",
    "tsa",
    "clave segura",
    "certitoken",
]

INTERNAL_KEYWORDS = [
    "ldap",
    "directorio activo",
    "active directory",
    "servidor",
    "infraestructura",
    "red interna",
    "vpn",
    "base de datos",
    "bd",
    "interno",
    "correo interno",
    "aplicativo interno",
]

ALERT_KEYWORDS = [
    "alerta",
    "alarma",
    "monitoreo",
    "monitor",
    "zabbix",
    "grafana",
    "prometheus",
    "observabilidad",
    "cpu",
    "memoria",
    "disco",
    "latencia",
    "packet loss",
    "perdida de paquetes",
    "indisponibilidad",
    "caida",
    "caido",
    "down",
    "critico",
    "critical",
    "warning",
]

INCIDENT_CASE_REDIRECT_PHRASES = [
    "no corresponde a un inc",
    "no corresponde a un incidente",
    "no es un inc",
    "no es un incidente",
    "canales de atencion",
    "mesa de ayuda",
    "mesadeayuda",
    "para futuras solicitudes",
    "no cuenta con la informacion suficiente",
]

INCIDENT_EXTERNAL_CASE_HINTS = [
    "descargar",
    "descarga",
    "instalacion",
    "manual",
    "activar",
    "activacion",
    "pendiente por descargar",
    "link de descarga",
    "numero de caso",
    "cambiar el nombre",
    "nombre del certificado",
    "certificado intermedio",
    "renovacion",
    "compra",
    "orden ",
    " op ",
    "pdf",
    "documento",
    "vuce",
    "firma digital",
    "token fisico",
]

INCIDENT_EXTERNAL_CASE_STRONG_HINTS = [
    "descargar firma",
    "manual",
    "instalacion",
    "numero de caso",
    "pendiente por descargar",
    "link de descarga",
    "nombre del certificado",
    "cambiar el nombre",
    "renovacion",
    "compra",
    "orden ",
    "token fisico",
    "vuce",
    "activacion de una firma",
]

INCIDENT_TRUE_INCIDENT_HINTS = [
    "caida",
    "caido",
    "indisponibilidad",
    "intermitencia",
    "lentitud",
    "timeout",
    "varios usuarios",
    "masiva",
    "interrupcion del servicio",
    "servicio caido",
    "portal de ventas",
    "no se visualizan transacciones",
]

INCIDENT_COMPONENT_RULES = [
    ("OCSP", ["ocsp"]),
    ("Certitoken", ["certitoken", "token virtual", "token"]),
    ("Clave Segura", ["clave segura"]),
    ("LDAP / Directorio Activo", ["ldap", "directorio activo", "active directory"]),
    ("Base de datos", ["base de datos", "database", "sql", "bd"]),
    ("Correo / Notificaciones", ["correo", "smtp", "mail"]),
    ("Certificado / cadena de confianza", ["ssl", "certificado ssl", "certificado", "cadena de confianza", "cadena de certificados"]),
    ("Servicios TSA", ["tsa", "timestamp", "sello de tiempo"]),
    ("Portal RPOST", ["rpost"]),
    ("Red / Conectividad", ["red", "conexion", "conectividad", "vpn", "packet loss", "perdida de paquetes", "latencia", "enlace"]),
    ("Infraestructura", ["cpu", "memoria", "disco", "filesystem", "espacio", "swap", "servidor"]),
    ("Monitoreo / NOC", ["alerta", "alarma", "monitoreo", "monitor", "zabbix", "grafana", "prometheus", "solarwinds"]),
    ("Proveedor externo", ["proveedor", "tercero", "externo"]),
    ("Firma digital", ["firma", "firmar"]),
]

INCIDENT_SYMPTOM_RULES = [
    ("falsa alarma", ["falsa alarma"]),
    ("caida o indisponibilidad", ["caida", "caido", "indisponibilidad", "fuera de servicio", "down", "no responde"]),
    ("lentitud o degradacion", ["lentitud", "degradacion", "degradado", "intermitencia", "intermitente"]),
    ("timeout", ["timeout", "time out"]),
    ("falla de autenticacion o acceso", ["autenticacion", "login", "acceso", "credenciales"]),
    ("error al firmar o validar", ["firmar", "firma", "validacion", "no valida", "invalido", "rechazo de firma"]),
    ("consumo alto de recursos", ["cpu", "memoria", "disco", "filesystem", "espacio", "swap"]),
    ("alerta de monitoreo", ["alerta", "alarma", "monitoreo", "zabbix", "grafana", "prometheus", "solarwinds"]),
]

CASE_USAGE_KEYWORDS = [
    "contrasena",
    "password",
    "como usar",
    "como se usa",
    "como firmar",
    "guia",
    "instructivo",
    "manual",
    "configuracion",
    "configurar",
    "instalacion",
    "instalar",
    "parametrizacion",
    "parametrizar",
    "capacitacion",
    "acompanamiento",
    "orientacion",
    "uso",
    "consulta",
    "asesoria",
    "soporte tecnico",
    "compatibilidad",
    "diligenciamiento",
    "solicitud de soporte",
    "validacion del proceso",
    "prueba funcional",
    "paso a paso",
]

CASE_INSTALLATION_KEYWORDS = [
    "instalacion",
    "instalaciones",
    "instalar",
    "instale",
    "instalada",
    "instalado",
    "tecnico de instalacion",
    "tecnicos de instalacion",
    "canal de agendamiento",
]

CASE_INSTALLATION_SCHEDULING_KEYWORDS = [
    "agendar",
    "agendamiento",
    "agenda",
    "programar una cita",
    "programar cita",
    "programar instalacion",
    "programar instalaciones",
    "sesion de instalacion",
    "sesion para instalacion",
]

CASE_INSTALLATION_CONTEXT_HINTS = [
    "instalacion",
    "instalaciones",
    "instalar",
    "activacion",
    "firma digital",
    "firma",
    "token",
    "certificado",
]

CASE_USAGE_REQUEST_HINTS = [
    "solicitud",
    "asesoria",
    "soporte tecnico",
    "compatibilidad",
    "acompanamiento",
]

CASE_RESOLUTION_HINTS = [
    "se informa",
    "se indica",
    "se explica",
    "se orienta",
    "se brinda soporte",
    "se comparte",
    "se envia",
    "se envian",
    "se realiza prueba",
    "se valida",
    "se ajusta",
    "se configura",
    "se instala",
    "solucionado",
    "resuelto",
    "quedo solucionado",
    "se da solucion",
]

CASE_STRONG_USAGE_HINTS = [
    "de acuerdo con su solicitud",
    "se brinda la informacion",
    "se brinda informacion",
    "se comparte el paso a paso",
    "se comparte procedimiento",
    "se remite",
    "validar nuevamente",
    "por favor realizar",
    "comunicarse con el area",
    "asesoria",
    "acompanamiento",
    "soporte",
]

CASE_FAILURE_KEYWORDS = [
    "error",
    "falla",
    "incidente",
    "caida",
    "indisponibilidad",
    "intermitencia",
    "timeout",
    "no funciona",
    "no responde",
    "bug",
    "afectacion masiva",
    "degradacion",
    "lentitud",
]

CASE_STRONG_FAILURE_KEYWORDS = [
    "caida total",
    "afectacion masiva",
    "indisponibilidad",
    "intermitencia",
    "timeout",
    "degradacion",
    "lentitud",
    "no responde",
    "no funciona",
    "bug",
]

CASE_INCIDENT_ESCALATION_KEYWORDS = [
    "escalado",
    "escalada",
    "escalamiento",
    "se escala",
    "se escalo",
    "escalado a proveedor",
    "escalado al proveedor",
    "escalado a desarrollo",
    "escalado al area de desarrollo",
    "escalado a nuestra area de desarrollo",
]

CASE_INCIDENT_CONTEXT_KEYWORDS = [
    *CASE_FAILURE_KEYWORDS,
    *CASE_STRONG_FAILURE_KEYWORDS,
    "incidente",
    "novedad",
    "inconveniente",
    "problema",
    "hallazgo",
    "causa raiz",
    "mitigar",
    "proveedor",
    "desarrollo",
    "logs",
]

CASE_ESCALATION_IGNORE_HINTS = [
    "pqrf",
    "canal oficial",
    "registre su requerimiento",
    "registre su solicitud",
    "tramite adecuado",
    "solicitudes",
    "peticion",
    "peticiones",
]

CASE_EXTERNAL_PLATFORM_KEYWORDS = [
    "adobe",
    "adove",
    "acrobat",
    "acrobat reader",
    "reader dc",
    "adobe reader",
    "adobe sign",
    "autofirma",
    "docusign",
]


def normalizar_texto(valor):
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return ""
    texto = str(valor).strip().lower()
    texto = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", texto)


def safe_text(valor):
    if valor is None or pd.isna(valor):
        return ""
    return str(valor).strip()


def safe_float(valor):
    try:
        if valor is None or pd.isna(valor):
            return None
        return float(valor)
    except Exception:
        return None


def normalizar_clave(valor):
    return normalizar_texto(valor)


def renombrar_columnas(df, aliases):
    df = df.copy()
    columnas = {normalizar_clave(col): col for col in df.columns}
    renames = {}
    for destino, opciones in aliases.items():
        for opcion in opciones:
            original = columnas.get(normalizar_clave(opcion))
            if original is not None:
                renames[original] = destino
                break
    df = df.rename(columns=renames)
    df = df.loc[:, ~df.columns.duplicated()]
    df = df.loc[:, ~df.columns.astype(str).str.contains("^Unnamed", na=False)]
    return df


def valor_fila(row, clave, default=""):
    valor = row.get(clave, default)
    if valor is None or pd.isna(valor):
        return default
    return valor


def unir_textos(row, campos):
    return " ".join(normalizar_texto(valor_fila(row, campo)) for campo in campos).strip()


def contar_coincidencias(texto, palabras):
    return sum(1 for palabra in palabras if palabra in texto)


def puntuar_texto(texto, palabras, peso=1):
    return contar_coincidencias(texto, palabras) * peso


def caso_cerrado(row):
    estado = normalizar_texto(valor_fila(row, "estado"))
    cerrado = normalizar_texto(valor_fila(row, "cerrado"))
    return "cerrado" in estado or cerrado != ""


def es_caso_instalacion(texto):
    if any(palabra in texto for palabra in CASE_INSTALLATION_KEYWORDS):
        return True

    agenda_instalacion = any(
        palabra in texto for palabra in CASE_INSTALLATION_SCHEDULING_KEYWORDS
    )
    contexto_instalacion = any(
        palabra in texto for palabra in CASE_INSTALLATION_CONTEXT_HINTS
    )
    return agenda_instalacion and contexto_instalacion


def es_caso_incidente(texto_apertura, texto_resolucion, texto_total):
    if re.search(r"\binc[-\s]?\d{3,}\b", texto_total):
        return True

    if "incidente" in texto_resolucion:
        return True

    if "incidente" in texto_apertura:
        return True

    if any(frase in texto_resolucion for frase in CASE_ESCALATION_IGNORE_HINTS):
        return False

    if not any(frase in texto_resolucion for frase in CASE_INCIDENT_ESCALATION_KEYWORDS):
        return False

    contexto = " ".join([texto_apertura, texto_resolucion])
    return any(palabra in contexto for palabra in CASE_INCIDENT_CONTEXT_KEYWORDS)


def texto_resolucion_incidente(row):
    return unir_textos(
        row,
        [
            "observaciones_trabajo",
            "observaciones_adicionales",
            "actualizaciones",
            "despues_aprobacion",
            "despues_rechazo",
        ],
    )


def es_caso_cliente_externo(row, texto=None, texto_resolucion=None):
    texto = texto or unir_textos(row, INCIDENT_TEXT_FIELDS)
    texto_resolucion = texto_resolucion or texto_resolucion_incidente(row)
    impacto = normalizar_texto(valor_fila(row, "impacto"))

    if any(frase in texto_resolucion for frase in INCIDENT_CASE_REDIRECT_PHRASES):
        return True

    if any(frase in texto for frase in INCIDENT_EXTERNAL_CASE_STRONG_HINTS):
        return not any(frase in texto for frase in INCIDENT_TRUE_INCIDENT_HINTS)

    if "4 - baja" not in impacto:
        return False

    if any(frase in texto for frase in INCIDENT_TRUE_INCIDENT_HINTS):
        return False

    return any(frase in texto for frase in INCIDENT_EXTERNAL_CASE_HINTS)


def get_conn():
    return sqlite3.connect(DB, check_same_thread=False)


def normalizar_email(email):
    return str(email or "").strip().lower()


def hash_password(password):
    salt = os.urandom(16)
    password_hash = hashlib.pbkdf2_hmac(
        "sha256",
        str(password).encode("utf-8"),
        salt,
        260000,
    )
    return f"pbkdf2_sha256$260000${salt.hex()}${password_hash.hex()}"


def verificar_password(password, password_hash):
    try:
        algoritmo, iteraciones, salt_hex, hash_hex = str(password_hash).split("$")
        if algoritmo != "pbkdf2_sha256":
            return False
        calculado = hashlib.pbkdf2_hmac(
            "sha256",
            str(password).encode("utf-8"),
            bytes.fromhex(salt_hex),
            int(iteraciones),
        ).hex()
        return hmac.compare_digest(calculado, hash_hex)
    except Exception:
        return False


def usuario_por_email(email):
    email = normalizar_email(email)
    conn = get_conn()
    row = conn.execute(
        """
        SELECT email, password_hash, role, active, created_at, last_login
        FROM app_users
        WHERE email = ?
        """,
        (email,),
    ).fetchone()
    conn.close()
    if not row:
        return None
    return {
        "email": row[0],
        "password_hash": row[1],
        "role": row[2],
        "active": bool(row[3]),
        "created_at": row[4],
        "last_login": row[5],
    }


def autenticar_usuario(email, password):
    usuario = usuario_por_email(email)
    if not usuario or not usuario["active"]:
        return None
    if not verificar_password(password, usuario["password_hash"]):
        return None

    conn = get_conn()
    conn.execute(
        "UPDATE app_users SET last_login = CURRENT_TIMESTAMP WHERE email = ?",
        (usuario["email"],),
    )
    conn.commit()
    conn.close()
    usuario["last_login"] = "CURRENT_TIMESTAMP"
    usuario.pop("password_hash", None)
    return usuario


def listar_usuarios():
    conn = get_conn()
    rows = conn.execute(
        """
        SELECT email, role, active, created_at, last_login
        FROM app_users
        ORDER BY role, email
        """
    ).fetchall()
    conn.close()
    return pd.DataFrame(
        rows,
        columns=["email", "role", "active", "created_at", "last_login"],
    )


def guardar_usuario(email, password, role="viewer", active=True):
    email = normalizar_email(email)
    role = role if role in ("admin", "viewer") else "viewer"
    if not email:
        raise ValueError("El correo es obligatorio.")
    if password is not None and len(str(password)) < 8:
        raise ValueError("La contrasena debe tener al menos 8 caracteres.")

    conn = get_conn()
    existe = conn.execute("SELECT 1 FROM app_users WHERE email = ?", (email,)).fetchone()
    if existe:
        if password:
            conn.execute(
                """
                UPDATE app_users
                SET password_hash = ?, role = ?, active = ?
                WHERE email = ?
                """,
                (hash_password(password), role, int(active), email),
            )
        else:
            conn.execute(
                "UPDATE app_users SET role = ?, active = ? WHERE email = ?",
                (role, int(active), email),
            )
    else:
        if not password:
            raise ValueError("La contrasena es obligatoria para crear el usuario.")
        conn.execute(
            """
            INSERT INTO app_users (email, password_hash, role, active, created_at)
            VALUES (?, ?, ?, ?, CURRENT_TIMESTAMP)
            """,
            (email, hash_password(password), role, int(active)),
        )
    conn.commit()
    conn.close()


def eliminar_usuario(email):
    email = normalizar_email(email)
    conn = get_conn()
    conn.execute("DELETE FROM app_users WHERE email = ?", (email,))
    conn.commit()
    conn.close()


def ensure_table_columns(conn, table_name, columns):
    existentes = {row[1] for row in conn.execute(f"PRAGMA table_info({table_name})").fetchall()}
    for nombre, tipo in columns.items():
        if nombre not in existentes:
            conn.execute(f"ALTER TABLE {table_name} ADD COLUMN {nombre} {tipo}")


def init_db():
    conn = get_conn()
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS app_users (
            email TEXT PRIMARY KEY,
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL DEFAULT 'viewer',
            active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT,
            last_login TEXT
        )
        """
    )
    admin_email = normalizar_email(ADMIN_EMAIL)
    admin_existe = conn.execute(
        "SELECT 1 FROM app_users WHERE email = ?",
        (admin_email,),
    ).fetchone()
    if not admin_existe:
        conn.execute(
            """
            INSERT INTO app_users (email, password_hash, role, active, created_at)
            VALUES (?, ?, 'admin', 1, CURRENT_TIMESTAMP)
            """,
            (admin_email, hash_password(INITIAL_ADMIN_PASSWORD)),
        )
    conn.execute(
        """
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
        """
    )
    conn.execute(
        """
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
        """
    )
    ensure_table_columns(
        conn,
        "incidents",
        {
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
            "duracion_horas": "REAL",
        },
    )
    conn.commit()
    conn.close()


def segundos_a_horas(valor):
    try:
        if valor is None or pd.isna(valor):
            return None
        return round(float(valor) / 3600, 2)
    except Exception:
        return None


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


def tipificar_caso(row):
    descripcion = normalizar_texto(valor_fila(row, "descripcion"))
    causa = normalizar_texto(valor_fila(row, "causa"))
    codigo_resolucion = normalizar_texto(valor_fila(row, "codigo_resolucion"))
    notas_resolucion = normalizar_texto(valor_fila(row, "notas_resolucion"))
    observaciones_adicionales = normalizar_texto(valor_fila(row, "observaciones_adicionales"))
    observaciones_trabajo = normalizar_texto(valor_fila(row, "observaciones_trabajo"))

    texto = " ".join(
        [
            descripcion,
            causa,
            codigo_resolucion,
            notas_resolucion,
            observaciones_adicionales,
            observaciones_trabajo,
        ]
    )

    if "phishing" in texto:
        return "1 - phishing"

    if es_caso_instalacion(texto):
        return "8 - Instalaciones"

    texto_resolucion = " ".join(
        [
            codigo_resolucion,
            notas_resolucion,
            observaciones_adicionales,
            observaciones_trabajo,
        ]
    )
    texto_apertura = " ".join([descripcion, causa])

    if es_caso_incidente(texto_apertura, texto_resolucion, texto):
        return "5 - incidente"

    if any(plataforma in texto for plataforma in CASE_EXTERNAL_PLATFORM_KEYWORDS):
        return "6 - Plataformas Ext"

    uso_score = (
        puntuar_texto(texto_apertura, CASE_USAGE_KEYWORDS, 1)
        + puntuar_texto(texto_resolucion, CASE_USAGE_KEYWORDS, 2)
        + puntuar_texto(texto_resolucion, CASE_RESOLUTION_HINTS, 3)
        + puntuar_texto(texto_resolucion, CASE_STRONG_USAGE_HINTS, 4)
    )

    falla_score = (
        puntuar_texto(texto_apertura, CASE_FAILURE_KEYWORDS, 3)
        + puntuar_texto(texto_resolucion, CASE_FAILURE_KEYWORDS, 1)
        + puntuar_texto(texto_apertura, CASE_STRONG_FAILURE_KEYWORDS, 4)
    )

    if "soporte" in codigo_resolucion or "soporte" in causa:
        uso_score += 2

    if caso_cerrado(row) and uso_score > 0:
        uso_score += 2

    if caso_cerrado(row) and any(p in texto_apertura for p in CASE_USAGE_REQUEST_HINTS):
        uso_score += 3

    if uso_score > 0 and "incidente" not in texto_apertura:
        uso_score += 1

    if uso_score > 0 and uso_score >= falla_score:
        return "2 - Soporte Uso"
    if falla_score > uso_score:
        return "3 - Soporte Falla"

    if any(palabra in texto for palabra in ["cert", "pagar", "biometria"]):
        return "4 - solicitudes"
    if "externo" in texto:
        return "6 - Plataformas Ext"
    return "7 - No Aplica"


def preparar_casos(df):
    df = renombrar_columnas(df, CASE_ALIASES)
    for columna in CASE_ALIASES:
        if columna not in df.columns:
            df[columna] = None
    df = df.replace({pd.NA: None})
    df["numero"] = df["numero"].apply(safe_text)
    df = df[df["numero"] != ""]
    return df.drop_duplicates(subset=["numero"], keep="last")


def guardar_casos(df):
    conn = get_conn()
    cargados = 0
    df = preparar_casos(df)
    if df.empty:
        conn.close()
        return 0, 0

    cur = conn.cursor()
    total_anterior = cur.execute("SELECT COUNT(*) FROM cases").fetchone()[0]
    cur.execute("DELETE FROM cases")

    for _, row in df.iterrows():
        numero = safe_text(valor_fila(row, "numero"))
        data = (
            numero,
            safe_text(valor_fila(row, "descripcion")),
            safe_text(valor_fila(row, "contacto")),
            safe_text(valor_fila(row, "cuenta")),
            safe_text(valor_fila(row, "codigo_resolucion")),
            safe_text(valor_fila(row, "canal")),
            safe_text(valor_fila(row, "estado")),
            safe_text(valor_fila(row, "prioridad")),
            safe_text(valor_fila(row, "asignado")),
            safe_text(valor_fila(row, "actualizado")),
            safe_text(valor_fila(row, "creado_por")),
            safe_text(valor_fila(row, "creado")),
            safe_text(valor_fila(row, "producto")),
            safe_text(valor_fila(row, "cerrado")),
            safe_text(valor_fila(row, "causa")),
            safe_text(valor_fila(row, "notas_resolucion")),
            safe_text(valor_fila(row, "observaciones_adicionales")),
            safe_text(valor_fila(row, "observaciones_trabajo")),
            tipificar_caso(row),
            tiempo(valor_fila(row, "creado"), valor_fila(row, "cerrado")),
        )
        cur.execute("INSERT INTO cases VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", data)
        cargados += 1
    conn.commit()
    conn.close()
    return cargados, total_anterior


def load_casos():
    conn = get_conn()
    df = pd.read_sql("SELECT * FROM cases", conn)
    conn.close()

    if not df.empty:
        df["tipificacion"] = df.apply(tipificar_caso, axis=1)

    return df


def tipificacion_incidente(row):
    categoria = normalizar_texto(valor_fila(row, "categoria"))
    texto = unir_textos(row, INCIDENT_TEXT_FIELDS)
    texto_resolucion = texto_resolucion_incidente(row)
    es_externo = any(p in texto for p in EXTERNAL_KEYWORDS)
    es_interno = any(p in texto for p in INTERNAL_KEYWORDS)

    if "consulta" in categoria:
        return "Consulta"
    if "solicitud" in categoria:
        return "Solicitud"
    if "noc" in texto:
        return "NOC"

    if es_externo and es_caso_cliente_externo(row, texto, texto_resolucion):
        return "Caso Cliente Externo"

    if "incidente" in categoria:
        if es_externo:
            return "Cliente Externo"
        if es_interno:
            return "Cliente Interno"
        return "Cliente Interno"
    if es_externo:
        return "Cliente Externo"
    if es_interno:
        return "Cliente Interno"
    return "Cliente Interno"


def es_alerta_incidente(row, tipificacion=None):
    tipificacion = tipificacion or tipificacion_incidente(row)
    texto = unir_textos(row, INCIDENT_CAUSE_FIELDS + ["grupo_asignacion", "servicio_negocio", "impacto", "estado"])
    if tipificacion == "Caso Cliente Externo":
        return "No"
    if tipificacion == "NOC":
        return "Si"
    if any(p in texto for p in ALERT_KEYWORDS):
        return "Si"
    return "No"


def inferir_componente_incidente(texto, tipificacion=None, es_alerta=None):
    for etiqueta, palabras in INCIDENT_COMPONENT_RULES:
        if any(palabra in texto for palabra in palabras):
            return etiqueta

    if tipificacion == "NOC" or es_alerta == "Si":
        return "Monitoreo / NOC"

    return ""


def inferir_sintoma_incidente(texto):
    for etiqueta, palabras in INCIDENT_SYMPTOM_RULES:
        if any(palabra in texto for palabra in palabras):
            return etiqueta

    tipo_error = [
        "error",
        "falla",
        "problema",
        "afectacion",
        "afectación",
        "incidente",
    ]
    if any(palabra in texto for palabra in tipo_error):
        return "error operativo reportado"

    return ""


def causa_raiz_manual_incidente(row):
    return safe_text(valor_fila(row, "causa_raiz_original"))


def causa_raiz_incidente(row, tipificacion=None, es_alerta=None):
    causa_original = causa_raiz_manual_incidente(row)
    if causa_original:
        return causa_original

    tipificacion = tipificacion or tipificacion_incidente(row)
    es_alerta = es_alerta or es_alerta_incidente(row, tipificacion)

    texto = unir_textos(
        row,
        INCIDENT_CAUSE_FIELDS + ["grupo_asignacion", "servicio_negocio", "impacto", "estado"],
    )

    componente = inferir_componente_incidente(texto, tipificacion, es_alerta)
    sintoma = inferir_sintoma_incidente(texto)
    tipo_falla = safe_text(valor_fila(row, "tipo_falla"))

    if componente and sintoma:
        return f"{componente} - {sintoma}"
    if componente and tipo_falla:
        return f"{componente} - {tipo_falla}"
    if componente:
        return componente
    if sintoma:
        return f"Hallazgo detectado - {sintoma}"
    if tipo_falla:
        return f"Hallazgo reportado - {tipo_falla}"

    return "Sin patron concluyente en descripcion y anotaciones"


def clasificacion_incidente(row):
    tipificacion = tipificacion_incidente(row)
    alerta = es_alerta_incidente(row, tipificacion)
    causa = causa_raiz_incidente(row, tipificacion, alerta)
    return tipificacion, alerta, causa


def preparar_incidentes(df):
    df = renombrar_columnas(df, INCIDENT_ALIASES)
    for columna in list(INCIDENT_ALIASES.keys()) + ["tipificacion_original", "causa_raiz_original"]:
        if columna not in df.columns:
            df[columna] = None
    if df["duracion_segundos"].isna().all() and "duracion" in df.columns:
        df["duracion_segundos"] = df["duracion"]
    df = df.replace({pd.NA: None})
    df["numero"] = df["numero"].apply(safe_text)
    df = df[df["numero"] != ""]
    return df.drop_duplicates(subset=["numero"], keep="last")


def guardar_incidentes(df):
    conn = get_conn()
    cargados = 0
    df = preparar_incidentes(df)
    if df.empty:
        conn.close()
        return 0, 0

    cur = conn.cursor()
    total_anterior = cur.execute("SELECT COUNT(*) FROM incidents").fetchone()[0]
    cur.execute("DELETE FROM incidents")

    for _, row in df.iterrows():
        numero = safe_text(valor_fila(row, "numero"))
        tipificacion_auto, es_alerta_auto, causa_raiz_auto = clasificacion_incidente(row)
        duracion_segundos = safe_float(valor_fila(row, "duracion_segundos"))
        duracion_horas = segundos_a_horas(duracion_segundos)
        data = (
            numero,
            safe_text(valor_fila(row, "solicitante")),
            safe_text(valor_fila(row, "breve_descripcion")),
            safe_text(valor_fila(row, "categoria")),
            safe_text(valor_fila(row, "prioridad")),
            safe_text(valor_fila(row, "estado")),
            safe_text(valor_fila(row, "grupo_asignacion")),
            safe_text(valor_fila(row, "asignado_a")),
            safe_text(valor_fila(row, "descripcion")),
            safe_text(valor_fila(row, "despues_aprobacion")),
            safe_text(valor_fila(row, "despues_rechazo")),
            duracion_segundos,
            safe_float(valor_fila(row, "minutos")),
            safe_text(valor_fila(row, "fecha_vencimiento_sla")),
            safe_text(valor_fila(row, "tipo_falla")),
            safe_text(valor_fila(row, "empresa")),
            safe_text(valor_fila(row, "creado_por")),
            safe_text(valor_fila(row, "cerrado")),
            safe_text(valor_fila(row, "escalado_proveedor")),
            safe_text(valor_fila(row, "servicio_negocio")),
            safe_text(valor_fila(row, "creado")),
            safe_text(valor_fila(row, "observaciones_trabajo")),
            safe_text(valor_fila(row, "observaciones_adicionales")),
            safe_text(valor_fila(row, "actualizaciones")),
            safe_text(valor_fila(row, "impacto")),
            safe_text(valor_fila(row, "lista_notas_trabajo")),
            safe_text(valor_fila(row, "tipificacion_original")),
            safe_text(valor_fila(row, "causa_raiz_original")),
            tipificacion_auto,
            causa_raiz_auto,
            es_alerta_auto,
            duracion_horas,
        )
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
        cargados += 1
    conn.commit()
    conn.close()
    return cargados, total_anterior


def load_incidentes():
    conn = get_conn()
    df = pd.read_sql("SELECT * FROM incidents", conn)
    conn.close()
    if not df.empty:
        clasificaciones = df.apply(clasificacion_incidente, axis=1, result_type="expand")
        clasificaciones.columns = ["tipificacion_auto", "es_alerta_auto", "causa_raiz_auto"]
        df[["tipificacion_auto", "es_alerta_auto", "causa_raiz_auto"]] = clasificaciones
    return df


def umbral_alerta_por_volumen(total, proporcion=0.08, minimo=2):
    if total <= 0:
        return minimo
    return max(minimo, ceil(total * proporcion))


def incidentes_relacionados(grupo, limite=6):
    numeros = []
    for valor in grupo.get("numero", pd.Series(dtype="object")).tolist():
        numero = safe_text(valor)
        if numero and numero not in numeros:
            numeros.append(numero)
    return numeros[:limite], max(0, len(numeros) - limite)


def resumir_top_causas(grupo, limite=3):
    causas = (
        grupo["causa_raiz_auto"]
        .replace("", pd.NA)
        .fillna("Sin inferencia")
        .value_counts()
        .head(limite)
    )
    resumen = []
    for causa, cantidad in causas.items():
        resumen.append(f"{causa} ({cantidad})")
    return ", ".join(resumen)


def construir_alertas_incidentes(df, sla_horas=24):
    if df is None or df.empty:
        return []

    trabajo = df.copy()
    trabajo = trabajo[trabajo["tipificacion_auto"].fillna("Cliente Interno") != "Caso Cliente Externo"].copy()
    if trabajo.empty:
        return []
    total = len(trabajo)
    trabajo["numero"] = trabajo["numero"].apply(safe_text)
    trabajo["causa_raiz_auto"] = trabajo["causa_raiz_auto"].replace("", pd.NA).fillna("Sin inferencia")
    trabajo["servicio_negocio"] = trabajo["servicio_negocio"].replace("", pd.NA).fillna("Sin servicio informado")
    trabajo["es_alerta_auto"] = trabajo["es_alerta_auto"].fillna("No")
    trabajo["duracion_horas"] = pd.to_numeric(trabajo["duracion_horas"], errors="coerce")

    alertas = []
    umbral_causa = umbral_alerta_por_volumen(total, proporcion=0.08, minimo=2)
    umbral_servicio = umbral_alerta_por_volumen(total, proporcion=0.1, minimo=3)

    for causa, grupo in sorted(trabajo.groupby("causa_raiz_auto"), key=lambda item: len(item[1]), reverse=True):
        cantidad = len(grupo)
        if causa == "Sin inferencia" or cantidad < umbral_causa:
            continue
        incidentes, adicionales = incidentes_relacionados(grupo)
        porcentaje = round((cantidad / total) * 100, 2)
        alertas.append(
            {
                "tipo": "causa_recurrente",
                "prioridad": cantidad,
                "titulo": f"Reincidencia por causa raiz: {causa}",
                "detalle": (
                    f"Se identificaron {cantidad} incidentes de {total} ({porcentaje}%) con la misma "
                    "causa raiz inferida a partir de la descripcion, anotaciones y observaciones."
                ),
                "incidentes": incidentes,
                "incidentes_adicionales": adicionales,
            }
        )

    for servicio, grupo in sorted(trabajo.groupby("servicio_negocio"), key=lambda item: len(item[1]), reverse=True):
        cantidad = len(grupo)
        if servicio == "Sin servicio informado" or cantidad < umbral_servicio:
            continue
        incidentes, adicionales = incidentes_relacionados(grupo)
        porcentaje = round((cantidad / total) * 100, 2)
        alertas.append(
            {
                "tipo": "servicio_concentrado",
                "prioridad": cantidad,
                "titulo": f"Concentracion de incidentes en servicio: {servicio}",
                "detalle": (
                    f"Este servicio acumula {cantidad} incidentes de {total} ({porcentaje}%) dentro del "
                    "conjunto analizado, por lo que conviene revisarlo como foco recurrente."
                ),
                "incidentes": incidentes,
                "incidentes_adicionales": adicionales,
            }
        )

    df_alertas = trabajo[trabajo["es_alerta_auto"] == "Si"].copy()
    if not df_alertas.empty:
        cantidad = len(df_alertas)
        incidentes, adicionales = incidentes_relacionados(df_alertas, limite=8)
        porcentaje = round((cantidad / total) * 100, 2)
        top_causas = resumir_top_causas(df_alertas)
        detalle = (
            f"Se tipificaron {cantidad} incidentes como alerta sobre {total} analizados ({porcentaje}%). "
            "Este consolidado sale directamente de la tipificacion de incidentes cargados."
        )
        if top_causas:
            detalle += f" Las causas mas repetidas dentro de este grupo son: {top_causas}."
        alertas.append(
            {
                "tipo": "volumen_alertas",
                "prioridad": cantidad,
                "titulo": "Volumen de incidentes tipificados como alerta",
                "detalle": detalle,
                "incidentes": incidentes,
                "incidentes_adicionales": adicionales,
            }
        )

    df_cerrados = trabajo[trabajo["estado"].fillna("") == "Cerrado"].copy()
    df_fuera_sla = df_cerrados[df_cerrados["duracion_horas"] >= sla_horas].copy()
    if not df_fuera_sla.empty:
        cantidad = len(df_fuera_sla)
        base = len(df_cerrados) if len(df_cerrados) > 0 else total
        incidentes, adicionales = incidentes_relacionados(df_fuera_sla)
        porcentaje = round((cantidad / base) * 100, 2) if base > 0 else 0
        alertas.append(
            {
                "tipo": "sla",
                "prioridad": cantidad,
                "titulo": f"Incidentes cerrados fuera de SLA de {sla_horas} horas",
                "detalle": (
                    f"Se encontraron {cantidad} incidentes cerrados con una duracion igual o superior a "
                    f"{sla_horas} horas, equivalente al {porcentaje}% de los incidentes cerrados analizados."
                ),
                "incidentes": incidentes,
                "incidentes_adicionales": adicionales,
            }
        )

    return sorted(alertas, key=lambda alerta: alerta["prioridad"], reverse=True)
