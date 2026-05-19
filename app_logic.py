import re
import hashlib
import hmac
import os
import time
import tomllib
import unicodedata
from datetime import timedelta
from math import ceil

import pandas as pd


def streamlit_secrets_path():
    project_secrets = os.path.join(os.path.dirname(__file__), ".streamlit", "secrets.toml")
    user_secrets = os.path.join(os.path.expanduser("~"), ".streamlit", "secrets.toml")
    if os.path.exists(project_secrets):
        return project_secrets
    if os.path.exists(user_secrets):
        return user_secrets
    return ""


def local_secret_value(name, default=""):
    secrets_path = streamlit_secrets_path()
    if not secrets_path:
        return default
    try:
        with open(secrets_path, "rb") as file:
            return tomllib.load(file).get(name, default)
    except Exception:
        return default


def config_value(name, default=""):
    value = os.environ.get(name)
    if value not in (None, ""):
        return value
    value = local_secret_value(name)
    if value not in (None, ""):
        return value
    try:
        from streamlit.runtime.scriptrunner import get_script_run_ctx

        if get_script_run_ctx(suppress_warning=True) is None:
            return default
        import streamlit as st

        return st.secrets.get(name, default)
    except Exception:
        return default


SUPABASE_URL = config_value("SUPABASE_URL")
SUPABASE_PUBLISHABLE_KEY = config_value("SUPABASE_PUBLISHABLE_KEY")
SUPABASE_DATABASE_URL = config_value("SUPABASE_DATABASE_URL")
ADMIN_EMAIL = config_value("APP_ADMIN_EMAIL")
INITIAL_ADMIN_PASSWORD = config_value("APP_ADMIN_PASSWORD")

DB_TRANSIENT_SQLSTATES = {"40P01", "40001"}
DB_MAX_REINTENTOS = 3
DB_REINTENTO_ESPERA_SEGUNDOS = 0.35
DB_BLOQUEO_CARGA_CASOS = "gestion_problemas_yerika:carga_casos"

# Literales reutilizados para evitar duplicidad y facilitar mantenimiento.
NUMERO_DE_CASO_TEXT = "numero de caso"
CLAVE_SEGURA_TEXT = "clave segura"
DIRECTORIO_ACTIVO_TEXT = "directorio activo"
PENDIENTE_POR_DESCARGAR_TEXT = "pendiente por descargar"
LINK_DESCARGA_TEXT = "link de descarga"
NOMBRE_CERTIFICADO_TEXT = "nombre del certificado"
ORDEN_TEXT = "orden "
FIRMA_DIGITAL_TEXT = "firma digital"
TOKEN_FISICO_TEXT = "token fisico"
FALSA_ALARMA_TEXT = "falsa alarma"
NO_RESPONDE_TEXT = "no responde"
SIN_INFERENCIA_TEXT = "Sin inferencia"
SIN_SERVICIO_INFORMADO_TEXT = "Sin servicio informado"

CASE_ALIASES = {
    "numero": ["numero", "número", NUMERO_DE_CASO_TEXT, "número de caso", "caso", "id caso"],
    "descripcion": ["breve descripcion", "breve descripción", "descripcion corta"],
    "contacto": ["contacto"],
    "cuenta": ["cuenta"],
    "codigo_resolucion": ["codigo de resolucion", "código de resolución"],
    "canal": ["canal"],
    "estado": ["estado"],
    "prioridad": ["prioridad"],
    "asignado": ["asignado a"],
    "actualizado": ["actualizado", "fecha actualizado", "actualizado el"],
    "creado_por": ["creado por"],
    "creado": ["creado", "fecha creado", "creado el", "fecha de creacion", "fecha de creación"],
    "producto": ["producto"],
    "cerrado": ["cerrado", "fecha cerrado", "cerrado el", "fecha de cierre"],
    "causa": ["causa"],
    "notas_resolucion": ["notas de resolucion", "notas de resolución"],
    "observaciones_adicionales": ["observaciones adicionales"],
    "observaciones_trabajo": [
        "observaciones y notas de trabajo",
        "observaciones de trabajo",
    ],
}

INCIDENT_ALIASES = {
    "numero": [
        "numero",
        "número",
        "numero de incidente",
        "número de incidente",
        "incidente",
        "id incidente",
    ],
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
        "fecha vencimiento sla",
    ],
    "tipo_falla": ["tipo de falla"],
    "empresa": ["empresa"],
    "creado_por": ["creado por"],
    "cerrado": ["cerrado", "fecha cerrado", "cerrado el", "fecha de cierre"],
    "escalado_proveedor": ["escalado a proveedor"],
    "servicio_negocio": ["servicio de negocio"],
    "creado": ["creado", "fecha creado", "creado el", "fecha de creacion", "fecha de creación"],
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
    CLAVE_SEGURA_TEXT,
    "certitoken",
]

INTERNAL_KEYWORDS = [
    "ldap",
    DIRECTORIO_ACTIVO_TEXT,
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
    PENDIENTE_POR_DESCARGAR_TEXT,
    LINK_DESCARGA_TEXT,
    NUMERO_DE_CASO_TEXT,
    "cambiar el nombre",
    NOMBRE_CERTIFICADO_TEXT,
    "certificado intermedio",
    "renovacion",
    "compra",
    ORDEN_TEXT,
    " op ",
    "pdf",
    "documento",
    "vuce",
    FIRMA_DIGITAL_TEXT,
    TOKEN_FISICO_TEXT,
]

INCIDENT_EXTERNAL_CASE_STRONG_HINTS = [
    "descargar firma",
    "manual",
    "instalacion",
    NUMERO_DE_CASO_TEXT,
    PENDIENTE_POR_DESCARGAR_TEXT,
    LINK_DESCARGA_TEXT,
    NOMBRE_CERTIFICADO_TEXT,
    "cambiar el nombre",
    "renovacion",
    "compra",
    ORDEN_TEXT,
    TOKEN_FISICO_TEXT,
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

INCIDENT_NO_IMPACT_HINTS = [
    FALSA_ALARMA_TEXT,
    "falso positivo",
    "falsos positivos",
    "sin afectacion",
    "sin afectacion del servicio",
    "sin indisponibilidad",
    "servicio operativo",
    "servicio estuvo operativo",
    "ping ok",
    "http 200",
    "servidor up",
    "se encuentra up",
    "no se evidencia afectacion",
    "no se evidencio afectacion",
]

EXTERNAL_SERVICE_KEYWORDS = [
    "certificacion digital",
    "certificación digital",
    "certihuella",
    "certifactura",
    "certitoken",
    CLAVE_SEGURA_TEXT,
    "firma biometrica",
    "firma biométrica",
    FIRMA_DIGITAL_TEXT,
    "ocsp",
    "pagina web",
    "página web",
    "portal",
    "rpost",
    "ssl",
    "ssps",
    "token virtual",
    "tsa",
]

INTERNAL_SERVICE_KEYWORDS = [
    "aplicaciones de escritorio",
    DIRECTORIO_ACTIVO_TEXT,
    "infraestructura",
    "ldap",
    "red interna",
    "servicios de infraestructura",
    "vpn",
]

INCIDENT_COMPONENT_RULES = [
    ("OCSP", ["ocsp"]),
    ("Certitoken", ["certitoken", "token virtual", "token"]),
    ("Clave Segura", [CLAVE_SEGURA_TEXT]),
    ("LDAP / Directorio Activo", ["ldap", DIRECTORIO_ACTIVO_TEXT, "active directory"]),
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
    (FALSA_ALARMA_TEXT, [FALSA_ALARMA_TEXT]),
    ("caida o indisponibilidad", ["caida", "caido", "indisponibilidad", "fuera de servicio", "down", NO_RESPONDE_TEXT]),
    ("lentitud o degradacion", ["lentitud", "degradacion", "degradado", "intermitencia", "intermitente"]),
    ("timeout", ["timeout", "time out"]),
    ("falla de autenticacion o acceso", ["autenticacion", "login", "acceso", "credenciales"]),
    ("error al firmar o validar", ["firmar", "firma", "validacion", "no valida", "invalido", "rechazo de firma"]),
    ("consumo alto de recursos", ["cpu", "memoria", "disco", "filesystem", "espacio", "swap"]),
    ("alerta de monitoreo", ["alerta", "alarma", "monitoreo", "zabbix", "grafana", "prometheus", "solarwinds"]),
]

TIPIFICACION_ATENCION_NOC = "Alertas y Consultas NOC"
TIPIFICACION_CASO_CLIENTE_EXTERNO = "Caso Cliente Externo"
TIPIFICACION_INCIDENTE_CLIENTE_EXTERNO = "Incidente Cliente Externo"
TIPIFICACION_INCIDENTE_INTERNO = "Incidente Interno"
TIPO_INCIDENTE_OPERATIVO = "Operativo"
TIPO_INCIDENTE_SEGURIDAD = "Seguridad"

TIPIFICACIONES_INCIDENTE_REAL = {
    TIPIFICACION_INCIDENTE_CLIENTE_EXTERNO,
    TIPIFICACION_INCIDENTE_INTERNO,
}

NOC_TIPIFICATION = TIPIFICACION_ATENCION_NOC

TIPIFICACIONES_NOC = {
    NOC_TIPIFICATION,
}

SLA_RESOLUCION_HORAS = {
    "Operativo": {
        "Critico": 4,
        "Alto": 8,
        "Moderado": 96,
        "Bajo": 192,
    },
    "Seguridad": {
        "Critico": 4,
        "Alto": 8,
        "Moderado": 24,
        "Bajo": 48,
    },
    "Consulta": {
        "Critico": 96,
        "Alto": 96,
        "Moderado": 96,
        "Bajo": 96,
    },
    "NOC": {
        "Critico": 96,
        "Alto": 96,
        "Moderado": 96,
        "Bajo": 96,
    },
    TIPIFICACION_CASO_CLIENTE_EXTERNO: {
        "Critico": 36,
        "Alto": 36,
        "Moderado": 36,
        "Bajo": 36,
    },
}

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
    FIRMA_DIGITAL_TEXT,
    "firma",
    "token",
    "certificado",
]

CASE_RULES_START_DATE = pd.Timestamp("2026-04-01")

CASE_IVR_KEYWORDS = [
    "ivr",
    "arbol telefonico",
    "menu telefonico",
    "llamada ivr",
]

CASE_NOT_CONNECTED_KEYWORDS = [
    "no se conecto",
    "no conecto",
    "no se conecta",
    "no se conectaron",
    "no conectaron",
    "cliente no conectado",
    "cliente no se conecto",
    "cliente no se conecta",
    "no asistio",
    "no ingreso",
    "no se presento",
]

CASE_AGENDA_KEYWORDS = [
    "agenda",
    "agendado",
    "agendada",
    "agendar",
    "agendamiento",
    "cita",
    "sesion",
    "programada",
    "programado",
]

CASE_EVIDENCE_KEYWORDS = [
    "evidencia",
    "evidencias",
    "adjunto",
    "adjunta",
    "adjuntos",
    "soporte adjunto",
    "captura",
    "pantallazo",
    "grabacion",
    "registro",
    "correo enviado",
    "meet.google",
    "teams.microsoft",
    "zoom.us",
    "link de la videollamada",
    "enlace de la videollamada",
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
    NO_RESPONDE_TEXT,
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
    NO_RESPONDE_TEXT,
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
    texto = normalizar_texto(valor)
    texto = texto.replace("n?", "nu").replace("n�", "nu")
    return re.sub(r"[^a-z0-9]+", " ", texto).strip()


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


def normalizar_fecha(valor):
    if valor is None or pd.isna(valor):
        return ""
    texto = safe_text(valor)
    if re.match(r"^\d{4}-\d{1,2}-\d{1,2}(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?$", texto):
        fecha = pd.to_datetime(texto, errors="coerce")
    elif isinstance(valor, str) and re.search(r"\d{1,2}[/-]\d{1,2}[/-]\d{2,4}", texto):
        fecha = pd.to_datetime(valor, errors="coerce", dayfirst=True)
    else:
        fecha = pd.to_datetime(valor, errors="coerce")
    if pd.isna(fecha):
        fecha = pd.to_datetime(valor, errors="coerce", dayfirst=True)
    if pd.isna(fecha):
        return safe_text(valor)
    return fecha.strftime("%Y-%m-%d %H:%M:%S")


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


def aplica_reglas_desde_abril(row):
    fecha = normalizar_fecha(valor_fila(row, "creado"))
    fecha = pd.to_datetime(fecha, errors="coerce")
    return not pd.isna(fecha) and fecha >= CASE_RULES_START_DATE


def contiene_alguna_palabra(texto, palabras):
    return any(palabra in texto for palabra in palabras)


def contiene_alguna_frase_completa(texto, palabras):
    return any(
        re.search(rf"(?<![a-z0-9]){re.escape(palabra)}(?![a-z0-9])", texto)
        for palabra in palabras
    )


def es_caso_ivr(texto, canal):
    return contiene_alguna_frase_completa(" ".join([texto, canal]), CASE_IVR_KEYWORDS)


def es_agenda_sin_evidencia(texto):
    if not contiene_alguna_palabra(texto, CASE_NOT_CONNECTED_KEYWORDS):
        return False
    if not contiene_alguna_palabra(texto, CASE_AGENDA_KEYWORDS):
        return False
    return (
        contiene_alguna_frase_completa(texto, CASE_NOT_CONNECTED_KEYWORDS)
        and contiene_alguna_frase_completa(texto, CASE_AGENDA_KEYWORDS)
        and not contiene_alguna_frase_completa(texto, CASE_EVIDENCE_KEYWORDS)
    )


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


def es_sin_afectacion_confirmada(texto):
    return any(frase in texto for frase in INCIDENT_NO_IMPACT_HINTS)


def contiene_noc(texto):
    return re.search(r"(?<![a-z0-9])noc(?![a-z0-9])", normalizar_texto(texto)) is not None


def es_origen_noc(row):
    campos = ["creado_por", "grupo_asignacion", "solicitante", "asignado_a"]
    return any(contiene_noc(valor_fila(row, campo)) for campo in campos)


def categoria_es_consulta(row):
    categoria = normalizar_texto(valor_fila(row, "categoria"))
    return "consulta" in categoria or "ayuda" in categoria


def categoria_es_seguridad(row):
    categoria = normalizar_texto(valor_fila(row, "categoria"))
    servicio = normalizar_texto(valor_fila(row, "servicio_negocio"))
    texto = unir_textos(row, INCIDENT_CAUSE_FIELDS + ["servicio_negocio", "grupo_asignacion"])
    palabras_seguridad = [
        "seguridad",
        "ciberseguridad",
        "malware",
        "phishing",
        "vulnerabilidad",
        "acceso no autorizado",
        "fuga",
        "indicador de compromiso",
    ]
    return (
        "seguridad" in categoria
        or "seguridad" in servicio
        or any(palabra in texto for palabra in palabras_seguridad)
    )


def es_reportante_externo(row):
    creado_por_original = safe_text(valor_fila(row, "creado_por"))
    creado_por = normalizar_texto(creado_por_original)
    if "@" not in creado_por_original:
        return False
    dominios_internos = ["certicamara.com", "certicamara.co"]
    return not any(dominio in creado_por for dominio in dominios_internos)


def es_empresa_externa(row):
    empresa = normalizar_texto(valor_fila(row, "empresa"))
    empresas_internas = {"", "sinempresa", "certicamara", "certicamara s a"}
    return empresa not in empresas_internas


def es_alerta_noc(row):
    if categoria_es_consulta(row):
        return False
    texto = unir_textos(
        row,
        INCIDENT_CAUSE_FIELDS
        + ["grupo_asignacion", "servicio_negocio", "impacto", "estado", "creado_por", "solicitante"],
    )
    return es_origen_noc(row) or any(palabra in texto for palabra in ALERT_KEYWORDS)


def es_cliente_externo_incidente(row, texto=None, texto_resolucion=None):
    texto = texto or unir_textos(row, INCIDENT_TEXT_FIELDS)
    texto_resolucion = texto_resolucion or texto_resolucion_incidente(row)
    texto_servicio = normalizar_texto(valor_fila(row, "servicio_negocio"))
    texto_completo = " ".join([texto, texto_servicio])
    return (
        es_empresa_externa(row)
        or es_reportante_externo(row)
        or any(p in texto_completo for p in EXTERNAL_KEYWORDS)
        or any(p in texto_completo for p in EXTERNAL_SERVICE_KEYWORDS)
        or es_caso_cliente_externo(row, texto, texto_resolucion)
    )


def es_cliente_interno_incidente(row, texto=None):
    texto = texto or unir_textos(row, INCIDENT_TEXT_FIELDS)
    texto_servicio = normalizar_texto(valor_fila(row, "servicio_negocio"))
    texto_completo = " ".join([texto, texto_servicio])
    return any(p in texto_completo for p in INTERNAL_KEYWORDS) or any(
        p in texto_completo for p in INTERNAL_SERVICE_KEYWORDS
    )


def origen_incidente(row, texto=None, texto_resolucion=None):
    if es_origen_noc(row):
        return "NOC"
    if es_cliente_externo_incidente(row, texto, texto_resolucion):
        return "Cliente externo"
    return "Cliente interno"


def es_incidente_real(row, texto=None, texto_resolucion=None):
    texto = texto or unir_textos(row, INCIDENT_TEXT_FIELDS)
    texto_resolucion = texto_resolucion or texto_resolucion_incidente(row)
    texto_completo = " ".join([texto, texto_resolucion])
    categoria = normalizar_texto(valor_fila(row, "categoria"))

    if es_caso_cliente_externo(row, texto, texto_resolucion):
        return False
    if "solicitud" in categoria:
        return False
    if es_sin_afectacion_confirmada(texto_completo):
        return False
    if categoria_es_seguridad(row):
        return True
    if "incidente" in categoria:
        return True
    if categoria_es_consulta(row):
        return False
    return any(frase in texto_completo for frase in INCIDENT_TRUE_INCIDENT_HINTS)


def ambito_incidente(row, texto=None, texto_resolucion=None):
    if es_cliente_externo_incidente(row, texto, texto_resolucion):
        return "Cliente externo"
    return "Cliente interno"


def db_placeholder():
    return "%s"


def db_placeholders(count):
    return ", ".join([db_placeholder()] * count)


def db_sql(sql):
    return sql.replace("?", "%s")


def db_bool(value):
    return bool(value)


def get_conn():
    if not SUPABASE_DATABASE_URL:
        raise RuntimeError("Configura SUPABASE_DATABASE_URL en Secrets para conectar la base de datos.")

    import psycopg

    return psycopg.connect(SUPABASE_DATABASE_URL, connect_timeout=30)


def db_execute(conn, sql, params=()):
    return conn.execute(db_sql(sql), params)


def es_error_db_transitorio(error):
    sqlstate = getattr(error, "sqlstate", None) or getattr(error, "pgcode", None)
    return sqlstate in DB_TRANSIENT_SQLSTATES


def ejecutar_con_reintentos_db(operacion, intentos=DB_MAX_REINTENTOS):
    for intento in range(intentos):
        try:
            return operacion()
        except Exception as error:
            if not es_error_db_transitorio(error) or intento >= intentos - 1:
                raise
            time.sleep(DB_REINTENTO_ESPERA_SEGUNDOS * (2**intento))


def bloquear_escritura_casos(conn):
    db_execute(conn, "SELECT pg_advisory_xact_lock(hashtext(?))", (DB_BLOQUEO_CARGA_CASOS,))
    db_execute(conn, "LOCK TABLE cases IN SHARE ROW EXCLUSIVE MODE")


def read_table(table_name):
    conn = get_conn()
    cursor = db_execute(conn, f"SELECT * FROM {table_name}")
    rows = cursor.fetchall()
    columns = [column[0] for column in cursor.description]
    conn.close()
    return pd.DataFrame(rows, columns=columns)


def upsert_sql(table_name, columns, conflict_column):
    column_sql = ", ".join(columns)
    placeholder_sql = db_placeholders(len(columns))
    updates = ", ".join(
        f"{column} = EXCLUDED.{column}"
        for column in columns
        if column != conflict_column
    )
    return (
        f"INSERT INTO {table_name} ({column_sql}) VALUES ({placeholder_sql}) "
        f"ON CONFLICT ({conflict_column}) DO UPDATE SET {updates}"
    )


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
    row = db_execute(
        conn,
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
    db_execute(
        conn,
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
    rows = db_execute(
        conn,
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
    existe = db_execute(conn, "SELECT 1 FROM app_users WHERE email = ?", (email,)).fetchone()
    if existe:
        if password:
            db_execute(
                conn,
                """
                UPDATE app_users
                SET password_hash = ?, role = ?, active = ?
                WHERE email = ?
                """,
                (hash_password(password), role, db_bool(active), email),
            )
        else:
            db_execute(
                conn,
                "UPDATE app_users SET role = ?, active = ? WHERE email = ?",
                (role, db_bool(active), email),
            )
    else:
        if not password:
            raise ValueError("La contrasena es obligatoria para crear el usuario.")
        db_execute(
            conn,
            """
            INSERT INTO app_users (email, password_hash, role, active, created_at)
            VALUES (?, ?, ?, ?, CURRENT_TIMESTAMP)
            """,
            (email, hash_password(password), role, db_bool(active)),
        )
    conn.commit()
    conn.close()


def eliminar_usuario(email):
    email = normalizar_email(email)
    conn = get_conn()
    db_execute(conn, "DELETE FROM app_users WHERE email = ?", (email,))
    conn.commit()
    conn.close()


def contar_incidentes():
    conn = get_conn()
    total = db_execute(conn, "SELECT COUNT(*) FROM incidents").fetchone()[0]
    conn.close()
    return total


def limpiar_incidentes():
    conn = get_conn()
    total = db_execute(conn, "SELECT COUNT(*) FROM incidents").fetchone()[0]
    db_execute(conn, "DELETE FROM incidents")
    conn.commit()
    conn.close()
    return total


def ensure_table_columns(conn, table_name, columns):
    existentes = {
        row[0]
        for row in db_execute(
            conn,
            """
            SELECT column_name
            FROM information_schema.columns
            WHERE table_schema = 'public' AND table_name = ?
            """,
            (table_name,),
        ).fetchall()
    }
    for nombre, tipo in columns.items():
        if nombre not in existentes:
            db_execute(conn, f"ALTER TABLE {table_name} ADD COLUMN {nombre} {tipo}")


def init_db():
    conn = get_conn()
    db_execute(
        conn,
        """
        CREATE TABLE IF NOT EXISTS app_users (
            email TEXT PRIMARY KEY,
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL DEFAULT 'viewer',
            active BOOLEAN NOT NULL DEFAULT TRUE,
            created_at TEXT,
            last_login TEXT
        )
        """
    )
    usuarios_existen = db_execute(conn, "SELECT 1 FROM app_users LIMIT 1").fetchone()
    if not usuarios_existen and ADMIN_EMAIL and INITIAL_ADMIN_PASSWORD:
        admin_email = normalizar_email(ADMIN_EMAIL)
        db_execute(
            conn,
            """
            INSERT INTO app_users (email, password_hash, role, active, created_at)
            VALUES (?, ?, 'admin', TRUE, CURRENT_TIMESTAMP)
            """,
            (admin_email, hash_password(INITIAL_ADMIN_PASSWORD)),
        )
    db_execute(
        conn,
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
    db_execute(
        conn,
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
            origen_auto TEXT,
            tipificacion_auto TEXT,
            tipo_incidente_auto TEXT,
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
            "origen_auto": "TEXT",
            "tipificacion_auto": "TEXT",
            "tipo_incidente_auto": "TEXT",
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


def textos_para_tipificacion_caso(row):
    descripcion = normalizar_texto(valor_fila(row, "descripcion"))
    causa = normalizar_texto(valor_fila(row, "causa"))
    codigo_resolucion = normalizar_texto(valor_fila(row, "codigo_resolucion"))
    notas_resolucion = normalizar_texto(valor_fila(row, "notas_resolucion"))
    observaciones_adicionales = normalizar_texto(valor_fila(row, "observaciones_adicionales"))
    observaciones_trabajo = normalizar_texto(valor_fila(row, "observaciones_trabajo"))

    texto_resolucion = " ".join(
        [
            codigo_resolucion,
            notas_resolucion,
            observaciones_adicionales,
            observaciones_trabajo,
        ]
    )
    texto_apertura = " ".join([descripcion, causa])
    texto = " ".join([texto_apertura, texto_resolucion])
    return texto, texto_apertura, texto_resolucion, codigo_resolucion, causa


def clasificacion_directa_caso(texto, texto_apertura, texto_resolucion):
    if "phishing" in texto:
        return "1 - phishing"
    if es_agenda_sin_evidencia(texto):
        return "10 - Cliente no asistio"
    if es_caso_instalacion(texto):
        return "9 - Redireccionamiento Agenda"
    if es_caso_incidente(texto_apertura, texto_resolucion, texto):
        return "5 - incidente"
    if any(plataforma in texto for plataforma in CASE_EXTERNAL_PLATFORM_KEYWORDS):
        return "6 - Plataformas Ext"
    return None


def calcular_scores_caso(row, texto_apertura, texto_resolucion, codigo_resolucion, causa):
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

    return uso_score, falla_score


def clasificacion_por_scores_caso(uso_score, falla_score):
    if uso_score > 0 and uso_score >= falla_score:
        return "2 - Soporte Uso"
    if falla_score > uso_score:
        return "3 - Soporte Falla"
    return None


def clasificacion_residual_caso(texto):
    if any(palabra in texto for palabra in ["cert", "pagar", "biometria"]):
        return "4 - solicitudes"
    if "externo" in texto:
        return "6 - Plataformas Ext"
    return "7 - No Aplica"


def tipificar_caso(row):
    texto, texto_apertura, texto_resolucion, codigo_resolucion, causa = textos_para_tipificacion_caso(row)

    clasificacion = clasificacion_directa_caso(texto, texto_apertura, texto_resolucion)
    if clasificacion:
        return clasificacion

    uso_score, falla_score = calcular_scores_caso(
        row,
        texto_apertura,
        texto_resolucion,
        codigo_resolucion,
        causa,
    )
    clasificacion = clasificacion_por_scores_caso(uso_score, falla_score)
    return clasificacion or clasificacion_residual_caso(texto)

def preparar_casos(df):
    df = renombrar_columnas(df, CASE_ALIASES)
    for columna in CASE_ALIASES:
        if columna not in df.columns:
            df[columna] = None
    df = df.replace({pd.NA: None})
    df["numero"] = df["numero"].apply(safe_text)
    df = df[df["numero"] != ""]
    return df.drop_duplicates(subset=["numero"], keep="last")


def meses_casos(df):
    if df.empty or "creado" not in df.columns:
        return []
    fechas = df["creado"].apply(normalizar_fecha)
    fechas = pd.to_datetime(fechas, errors="coerce")
    return fechas.dropna().dt.to_period("M").astype(str).sort_values().unique().tolist()


CASE_DB_COLUMNS = [
    "numero",
    "descripcion",
    "contacto",
    "cuenta",
    "codigo_resolucion",
    "canal",
    "estado",
    "prioridad",
    "asignado",
    "actualizado",
    "creado_por",
    "creado",
    "producto",
    "cerrado",
    "causa",
    "notas_resolucion",
    "observaciones_adicionales",
    "observaciones_trabajo",
    "tipificacion",
    "tiempo_respuesta",
]


def _guardar_casos_preparados(df, reemplazar_meses, duplicados_archivo):
    conn = get_conn()
    cargados = 0
    try:
        bloquear_escritura_casos(conn)
        meses_reemplazados = meses_casos(df) if reemplazar_meses else []
        eliminados = 0
        if meses_reemplazados:
            placeholders_meses = db_placeholders(len(meses_reemplazados))
            eliminados = db_execute(
                conn,
                f"DELETE FROM cases WHERE substr(creado, 1, 7) IN ({placeholders_meses})",
                meses_reemplazados,
            ).rowcount

        numeros = [safe_text(valor_fila(row, "numero")) for _, row in df.iterrows()]
        existentes = set()
        if numeros:
            placeholders = db_placeholders(len(numeros))
            existentes = {
                fila[0]
                for fila in db_execute(
                    conn,
                    f"SELECT numero FROM cases WHERE numero IN ({placeholders})",
                    numeros,
                ).fetchall()
            }
        reemplazados = len(existentes)

        for _, row in df.iterrows():
            numero = safe_text(valor_fila(row, "numero"))
            creado = normalizar_fecha(valor_fila(row, "creado"))
            cerrado = normalizar_fecha(valor_fila(row, "cerrado"))
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
                normalizar_fecha(valor_fila(row, "actualizado")),
                safe_text(valor_fila(row, "creado_por")),
                creado,
                safe_text(valor_fila(row, "producto")),
                cerrado,
                safe_text(valor_fila(row, "causa")),
                safe_text(valor_fila(row, "notas_resolucion")),
                safe_text(valor_fila(row, "observaciones_adicionales")),
                safe_text(valor_fila(row, "observaciones_trabajo")),
                tipificar_caso(row),
                tiempo(creado, cerrado),
            )
            db_execute(conn, upsert_sql("cases", CASE_DB_COLUMNS, "numero"), data)
            cargados += 1
        conn.commit()
        return cargados, reemplazados, eliminados, meses_reemplazados, duplicados_archivo
    except Exception:
        try:
            conn.rollback()
        except Exception:
            pass
        raise
    finally:
        conn.close()


def guardar_casos(df, reemplazar_meses=False):
    filas_recibidas = len(df)
    df = preparar_casos(df)
    duplicados_archivo = max(filas_recibidas - len(df), 0)
    if df.empty:
        return 0, 0, 0, [], duplicados_archivo

    df = df.sort_values("numero", kind="mergesort").reset_index(drop=True)
    return ejecutar_con_reintentos_db(
        lambda: _guardar_casos_preparados(df, reemplazar_meses, duplicados_archivo)
    )


def load_casos():
    df = read_table("cases")

    if not df.empty:
        tipificaciones = df.apply(tipificar_caso, axis=1)
        cambios = df["tipificacion"].fillna("") != tipificaciones.fillna("")
        if cambios.any():
            conn = get_conn()
            for numero, tipificacion in zip(df.loc[cambios, "numero"], tipificaciones.loc[cambios]):
                db_execute(
                    conn,
                    "UPDATE cases SET tipificacion=? WHERE numero=?",
                    (tipificacion, numero),
                )
            conn.commit()
            conn.close()
        df["tipificacion"] = tipificaciones

    return df


def tipificacion_incidente(row):
    texto = unir_textos(row, INCIDENT_TEXT_FIELDS)
    texto_resolucion = texto_resolucion_incidente(row)
    origen_noc = es_origen_noc(row)
    es_externo = es_cliente_externo_incidente(row, texto, texto_resolucion)
    incidente_real = es_incidente_real(row, texto, texto_resolucion)

    if origen_noc and not incidente_real:
        return TIPIFICACION_ATENCION_NOC
    if categoria_es_consulta(row):
        return TIPIFICACION_CASO_CLIENTE_EXTERNO if es_externo else TIPIFICACION_ATENCION_NOC
    if "solicitud" in normalizar_texto(valor_fila(row, "categoria")):
        return TIPIFICACION_CASO_CLIENTE_EXTERNO

    if es_externo and es_caso_cliente_externo(row, texto, texto_resolucion):
        return TIPIFICACION_CASO_CLIENTE_EXTERNO

    if origen_noc and not incidente_real:
        return TIPIFICACION_ATENCION_NOC

    if incidente_real and es_externo:
        return TIPIFICACION_INCIDENTE_CLIENTE_EXTERNO
    if incidente_real:
        return TIPIFICACION_INCIDENTE_INTERNO
    if es_externo:
        return TIPIFICACION_CASO_CLIENTE_EXTERNO
    return TIPIFICACION_INCIDENTE_INTERNO


def es_alerta_incidente(row, tipificacion=None):
    tipificacion = tipificacion or tipificacion_incidente(row)
    texto = unir_textos(row, INCIDENT_CAUSE_FIELDS + ["grupo_asignacion", "servicio_negocio", "impacto", "estado"])
    if es_origen_noc(row) and any(p in texto for p in ALERT_KEYWORDS):
        return "Si"
    return "No"


def tipo_incidente_auto(row, tipificacion=None):
    tipificacion = tipificacion or tipificacion_incidente(row)
    if tipificacion not in TIPIFICACIONES_INCIDENTE_REAL:
        return "No aplica"
    if categoria_es_seguridad(row):
        return TIPO_INCIDENTE_SEGURIDAD
    return TIPO_INCIDENTE_OPERATIVO


def motivo_caso_cliente_externo(row):
    texto = unir_textos(
        row,
        INCIDENT_CAUSE_FIELDS
        + ["grupo_asignacion", "servicio_negocio", "impacto", "estado", "empresa", "solicitante"],
    )
    reglas = [
        (
            "Instalacion / activacion",
            [
                "instalacion",
                "instalar",
                "activar",
                "activacion",
                TOKEN_FISICO_TEXT,
                "captores biometricos",
                "biometricos",
            ],
        ),
        (
            "Descarga / entrega",
            [
                "descargar",
                "descarga",
                LINK_DESCARGA_TEXT,
                PENDIENTE_POR_DESCARGAR_TEXT,
                "correo de descarga",
            ],
        ),
        (
            "Tramite / orden",
            [
                ORDEN_TEXT,
                " op ",
                "compra",
                "pago",
                "solicitud",
                "solicitudes",
                "tramite",
                "vuce",
            ],
        ),
        (
            "Guia / uso",
            [
                "manual",
                "guia",
                "instructivo",
                "como",
                "paso a paso",
                "orientacion",
            ],
        ),
        (
            "Validacion / documento",
            [
                "validar",
                "validez",
                "desconocida",
                "pdf",
                "documento",
                "certificado intermedio",
                NOMBRE_CERTIFICADO_TEXT,
            ],
        ),
    ]
    for motivo, palabras in reglas:
        if any(palabra in texto for palabra in palabras):
            return motivo
    return "Otro caso cliente externo"


def inferir_componente_incidente(texto, tipificacion=None, es_alerta=None):
    for etiqueta, palabras in INCIDENT_COMPONENT_RULES:
        if any(palabra in texto for palabra in palabras):
            return etiqueta

    if tipificacion == NOC_TIPIFICATION or es_alerta == "Si":
        return "Monitoreo / NOC"

    return ""


def normalizar_prioridad_incidente(valor):
    texto = normalizar_texto(valor)
    if re.search(r"(?<!\d)1(?!\d)", texto) or "crit" in texto:
        return "Critico"
    if re.search(r"(?<!\d)2(?!\d)", texto) or "alta" in texto or "alto" in texto:
        return "Alto"
    if re.search(r"(?<!\d)3(?!\d)", texto) or "moderad" in texto or "media" in texto:
        return "Moderado"
    if re.search(r"(?<!\d)4(?!\d)", texto) or "baja" in texto or "bajo" in texto:
        return "Bajo"
    return "Moderado"


def familia_sla_incidente(tipificacion, tipo_incidente=None):
    if tipificacion in TIPIFICACIONES_INCIDENTE_REAL and tipo_incidente == TIPO_INCIDENTE_SEGURIDAD:
        return "Seguridad"
    if tipificacion == NOC_TIPIFICATION:
        return "NOC"
    if tipificacion == TIPIFICACION_CASO_CLIENTE_EXTERNO:
        return TIPIFICACION_CASO_CLIENTE_EXTERNO
    if tipificacion in TIPIFICACIONES_INCIDENTE_REAL:
        return "Operativo"
    return "Sin SLA"


def sla_objetivo_horas_incidente(row):
    tipificacion = valor_fila(row, "tipificacion_auto") or tipificacion_incidente(row)
    tipo_incidente = valor_fila(row, "tipo_incidente_auto") or tipo_incidente_auto(row, tipificacion)
    familia = familia_sla_incidente(tipificacion, tipo_incidente)
    prioridad = normalizar_prioridad_incidente(valor_fila(row, "prioridad"))
    return SLA_RESOLUCION_HORAS.get(familia, {}).get(prioridad)


def aplica_sla_incidente(tipificacion):
    return tipificacion in TIPIFICACIONES_INCIDENTE_REAL


def aplica_sla_caso_cliente_externo(tipificacion):
    return tipificacion == TIPIFICACION_CASO_CLIENTE_EXTERNO


def aplica_sla_atencion_noc(tipificacion):
    return tipificacion == TIPIFICACION_ATENCION_NOC


def duracion_sla_horas_incidente(row):
    tipificacion = valor_fila(row, "tipificacion_auto") or tipificacion_incidente(row)
    if tipificacion == TIPIFICACION_CASO_CLIENTE_EXTERNO:
        horas = tiempo(normalizar_fecha(valor_fila(row, "creado")), normalizar_fecha(valor_fila(row, "cerrado")))
        return horas if isinstance(horas, (int, float)) else None
    duracion = safe_float(valor_fila(row, "duracion_horas"))
    if duracion is not None:
        return duracion
    return segundos_a_horas(safe_float(valor_fila(row, "duracion_segundos")))


def estado_sla_incidente(row):
    objetivo = safe_float(valor_fila(row, "sla_objetivo_horas"))
    duracion = safe_float(valor_fila(row, "duracion_sla_horas"))
    if duracion is None:
        duracion = safe_float(valor_fila(row, "duracion_horas"))
    if objetivo is None:
        return "Sin objetivo"
    if duracion is None:
        return "Sin duracion"
    return "Cumple" if duracion <= objetivo else "No cumple"


def agregar_campos_sla_incidentes(df):
    trabajo = df.copy()
    columnas_default = {
        "origen_auto": pd.Series(dtype="object"),
        "tipo_incidente_auto": pd.Series(dtype="object"),
        "prioridad_normalizada": pd.Series(dtype="object"),
        "familia_sla": pd.Series(dtype="object"),
        "sla_objetivo_horas": pd.Series(dtype="float"),
        "duracion_sla_horas": pd.Series(dtype="float"),
        "aplica_sla_incidente": pd.Series(dtype="bool"),
        "aplica_sla_caso_cliente_externo": pd.Series(dtype="bool"),
        "aplica_sla_atencion_noc": pd.Series(dtype="bool"),
        "estado_sla": pd.Series(dtype="object"),
    }
    if trabajo.empty:
        for columna, serie in columnas_default.items():
            trabajo[columna] = serie
        return trabajo

    clasificaciones = trabajo.apply(clasificacion_incidente_detallada, axis=1, result_type="expand")
    for columna in ["origen_auto", "tipificacion_auto", "tipo_incidente_auto", "es_alerta_auto", "causa_raiz_auto"]:
        trabajo[columna] = clasificaciones[columna]

    trabajo["prioridad_normalizada"] = trabajo["prioridad"].apply(normalizar_prioridad_incidente)
    trabajo["familia_sla"] = trabajo.apply(
        lambda row: familia_sla_incidente(row["tipificacion_auto"], row["tipo_incidente_auto"]),
        axis=1,
    )
    trabajo["sla_objetivo_horas"] = trabajo.apply(sla_objetivo_horas_incidente, axis=1)
    trabajo["duracion_sla_horas"] = trabajo.apply(duracion_sla_horas_incidente, axis=1)
    trabajo["aplica_sla_incidente"] = trabajo["tipificacion_auto"].apply(aplica_sla_incidente)
    trabajo["aplica_sla_caso_cliente_externo"] = trabajo["tipificacion_auto"].apply(aplica_sla_caso_cliente_externo)
    trabajo["aplica_sla_atencion_noc"] = trabajo["tipificacion_auto"].apply(aplica_sla_atencion_noc)
    trabajo["estado_sla"] = trabajo.apply(estado_sla_incidente, axis=1)
    return trabajo


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

    if tipificacion == TIPIFICACION_CASO_CLIENTE_EXTERNO:
        return motivo_caso_cliente_externo(row)

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


def clasificacion_incidente_detallada(row):
    texto = unir_textos(row, INCIDENT_TEXT_FIELDS)
    texto_resolucion = texto_resolucion_incidente(row)
    tipificacion = tipificacion_incidente(row)
    alerta = es_alerta_incidente(row, tipificacion)
    tipo_incidente = tipo_incidente_auto(row, tipificacion)
    causa = causa_raiz_incidente(row, tipificacion, alerta)
    return {
        "origen_auto": origen_incidente(row, texto, texto_resolucion),
        "tipificacion_auto": tipificacion,
        "tipo_incidente_auto": tipo_incidente,
        "es_alerta_auto": alerta,
        "causa_raiz_auto": causa,
    }


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


INCIDENT_DB_COLUMNS = [
    "numero",
    "solicitante",
    "breve_descripcion",
    "categoria",
    "prioridad",
    "estado",
    "grupo_asignacion",
    "asignado_a",
    "descripcion",
    "despues_aprobacion",
    "despues_rechazo",
    "duracion_segundos",
    "minutos",
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
    "tipificacion_original",
    "causa_raiz_original",
    "origen_auto",
    "tipificacion_auto",
    "tipo_incidente_auto",
    "causa_raiz_auto",
    "es_alerta_auto",
    "duracion_horas",
]


def guardar_incidentes(df):
    conn = get_conn()
    cargados = 0
    df = preparar_incidentes(df)
    if df.empty:
        conn.close()
        return 0, 0

    numeros = [safe_text(valor_fila(row, "numero")) for _, row in df.iterrows()]
    existentes = set()
    if numeros:
        placeholders = db_placeholders(len(numeros))
        existentes = {
            fila[0]
            for fila in db_execute(
                conn,
                f"SELECT numero FROM incidents WHERE numero IN ({placeholders})",
                numeros,
            ).fetchall()
        }
    reemplazados = len(existentes)

    for _, row in df.iterrows():
        numero = safe_text(valor_fila(row, "numero"))
        clasificacion = clasificacion_incidente_detallada(row)
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
            normalizar_fecha(valor_fila(row, "fecha_vencimiento_sla")),
            safe_text(valor_fila(row, "tipo_falla")),
            safe_text(valor_fila(row, "empresa")),
            safe_text(valor_fila(row, "creado_por")),
            normalizar_fecha(valor_fila(row, "cerrado")),
            safe_text(valor_fila(row, "escalado_proveedor")),
            safe_text(valor_fila(row, "servicio_negocio")),
            normalizar_fecha(valor_fila(row, "creado")),
            safe_text(valor_fila(row, "observaciones_trabajo")),
            safe_text(valor_fila(row, "observaciones_adicionales")),
            safe_text(valor_fila(row, "actualizaciones")),
            safe_text(valor_fila(row, "impacto")),
            safe_text(valor_fila(row, "lista_notas_trabajo")),
            safe_text(valor_fila(row, "tipificacion_original")),
            safe_text(valor_fila(row, "causa_raiz_original")),
            clasificacion["origen_auto"],
            clasificacion["tipificacion_auto"],
            clasificacion["tipo_incidente_auto"],
            clasificacion["causa_raiz_auto"],
            clasificacion["es_alerta_auto"],
            duracion_horas,
        )
        db_execute(conn, upsert_sql("incidents", INCIDENT_DB_COLUMNS, "numero"), data)
        cargados += 1
    conn.commit()
    conn.close()
    return cargados, reemplazados


def load_incidentes():
    df = read_table("incidents")
    if not df.empty:
        clasificaciones = df.apply(clasificacion_incidente_detallada, axis=1, result_type="expand")
        for columna in ["origen_auto", "tipificacion_auto", "tipo_incidente_auto", "es_alerta_auto", "causa_raiz_auto"]:
            df[columna] = clasificaciones[columna]
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
        .fillna(SIN_INFERENCIA_TEXT)
        .value_counts()
        .head(limite)
    )
    resumen = []
    for causa, cantidad in causas.items():
        resumen.append(f"{causa} ({cantidad})")
    return ", ".join(resumen)


def agregar_alertas_causas_recurrentes(alertas, base_recurrencia, total, umbral_causa):
    grupos = sorted(
        base_recurrencia.groupby("causa_raiz_auto"),
        key=lambda item: len(item[1]),
        reverse=True,
    )
    for causa, grupo in grupos:
        cantidad = len(grupo)
        if causa == SIN_INFERENCIA_TEXT or cantidad < umbral_causa:
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


def agregar_alertas_servicio_concentrado(alertas, base_recurrencia, total, umbral_servicio):
    grupos = sorted(
        base_recurrencia.groupby("servicio_negocio"),
        key=lambda item: len(item[1]),
        reverse=True,
    )
    for servicio, grupo in grupos:
        cantidad = len(grupo)
        if servicio == SIN_SERVICIO_INFORMADO_TEXT or cantidad < umbral_servicio:
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


def agregar_alerta_volumen_noc(alertas, trabajo):
    df_alertas = trabajo[
        (trabajo["es_alerta_auto"] == "Si")
        | (trabajo["tipificacion_auto"].fillna("") == NOC_TIPIFICATION)
    ].copy()
    if df_alertas.empty:
        return

    cantidad = len(df_alertas)
    incidentes, adicionales = incidentes_relacionados(df_alertas, limite=8)
    porcentaje = round((cantidad / len(trabajo)) * 100, 2)
    top_causas = resumir_top_causas(df_alertas)
    detalle = (
        f"Se separaron {cantidad} registros como alertas NOC o alertas de monitoreo sobre "
        f"{len(trabajo)} registros analizables ({porcentaje}%)."
    )
    if top_causas:
        detalle += f" Las causas mas repetidas dentro de este grupo son: {top_causas}."
    alertas.append(
        {
            "tipo": "volumen_alertas",
            "prioridad": cantidad,
            "titulo": "Volumen de alertas NOC / monitoreo",
            "detalle": detalle,
            "incidentes": incidentes,
            "incidentes_adicionales": adicionales,
        }
    )


def agregar_alerta_sla(alertas, incidentes_reales, total):
    df_cerrados = incidentes_reales[incidentes_reales["estado"].fillna("") == "Cerrado"].copy()
    df_fuera_sla = df_cerrados[
        (df_cerrados["sla_objetivo_horas"].notna())
        & (df_cerrados["duracion_horas"] > df_cerrados["sla_objetivo_horas"])
    ].copy()
    if df_fuera_sla.empty:
        return

    cantidad = len(df_fuera_sla)
    base = len(df_cerrados) if len(df_cerrados) > 0 else total
    incidentes, adicionales = incidentes_relacionados(df_fuera_sla)
    porcentaje = round((cantidad / base) * 100, 2) if base > 0 else 0
    alertas.append(
        {
            "tipo": "sla",
            "prioridad": cantidad,
            "titulo": "Incidentes cerrados fuera de SLA",
            "detalle": (
                f"Se encontraron {cantidad} incidentes cerrados por encima de su objetivo "
                f"segun prioridad y familia SLA, equivalente al {porcentaje}% de los incidentes cerrados analizados."
            ),
            "incidentes": incidentes,
            "incidentes_adicionales": adicionales,
        }
    )


def preparar_base_alertas_incidentes(df):
    trabajo = agregar_campos_sla_incidentes(df)
    trabajo = trabajo[
        ~trabajo["tipificacion_auto"].fillna("").isin([TIPIFICACION_CASO_CLIENTE_EXTERNO, NOC_TIPIFICATION])
    ].copy()
    if trabajo.empty:
        return trabajo

    trabajo["numero"] = trabajo["numero"].apply(safe_text)
    trabajo["causa_raiz_auto"] = trabajo["causa_raiz_auto"].replace("", pd.NA).fillna(SIN_INFERENCIA_TEXT)
    trabajo["servicio_negocio"] = trabajo["servicio_negocio"].replace("", pd.NA).fillna(SIN_SERVICIO_INFORMADO_TEXT)
    trabajo["es_alerta_auto"] = trabajo["es_alerta_auto"].fillna("No")
    trabajo["duracion_horas"] = pd.to_numeric(trabajo["duracion_horas"], errors="coerce")
    return trabajo


def construir_alertas_incidentes(df):
    if df is None or df.empty:
        return []

    trabajo = preparar_base_alertas_incidentes(df)
    if trabajo.empty:
        return []

    incidentes_reales = trabajo[trabajo["aplica_sla_incidente"]].copy()
    total = len(incidentes_reales) if not incidentes_reales.empty else len(trabajo)
    base_recurrencia = incidentes_reales if not incidentes_reales.empty else trabajo

    alertas = []
    agregar_alertas_causas_recurrentes(
        alertas,
        base_recurrencia,
        total,
        umbral_alerta_por_volumen(total, proporcion=0.08, minimo=2),
    )
    agregar_alertas_servicio_concentrado(
        alertas,
        base_recurrencia,
        total,
        umbral_alerta_por_volumen(total, proporcion=0.1, minimo=3),
    )
    agregar_alerta_volumen_noc(alertas, trabajo)
    agregar_alerta_sla(alertas, incidentes_reales, total)
    return sorted(alertas, key=lambda alerta: alerta["prioridad"], reverse=True)
