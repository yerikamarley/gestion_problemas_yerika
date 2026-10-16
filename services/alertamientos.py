"""Evidencia de alertamientos, sin inferir afectación del prefijo INC ni del NOC."""

import re
import unicodedata

import pandas as pd


CAMPOS_EVIDENCIA = (
    "breve_descripcion", "descripcion", "tipo_falla", "observaciones_trabajo",
    "observaciones_adicionales", "actualizaciones", "lista_notas_trabajo",
    "despues_aprobacion", "despues_rechazo", "tipificacion_original", "causa_raiz_original",
)
PATRON_ALERTA = re.compile(r"\b(?:alarmas?|alertas?|alertamientos?|fals[oa]s? positiv[oa]s?)\b")
PATRON_DESCARTE = re.compile(
    r"\b(?:fals[oa]s? alarmas?|fals[oa]s? positiv[oa]s?|sin afectacion(?: del servicio)?|"
    r"sin indisponibilidad|no se evidenci[ao](?:ron)? afectacion)\b"
)
AFECTACION = (
    r"(?:caid[oa]s?|indisponibilidad|intermitencia|degradacion|lentitud|"
    r"interrupcion(?: del servicio)?|afectacion(?: del servicio)?|"
    r"servicio no disponible|no responde|down|unreachable|timeouts?|"
    r"error(?:es)? (?:500|502|503|504)|falla(?:s)?(?: del servicio)?|"
    r"certificado (?:ssl )?(?:vencido|expirado)|ssl (?:vencido|expirado)|"
    r"no (?:genero|detecto) (?:la )?alerta|acceso no autorizado|malware|ransomware)"
)
PATRON_AFECTACION = re.compile(r"\b" + AFECTACION + r"\b")
# Se elimina únicamente la expresión negada o hipotética, no toda la nota:
# una misma nota puede descartar una alarma y confirmar otra caída.
PATRON_NO_CONFIRMADO = re.compile(
    r"\b(?:sin|no (?:hay|hubo|presenta|presento|existe)|"
    r"no se (?:evidencia|evidencio|presenta|presento|confirma|confirmo|detecta|detecto)|"
    r"se descarta|se descarto|posible|posibles|riesgo de|prevencion de)\s+"
    r"(?:(?:una?|la|el|ninguna?|evidencia de|signos de)\s+)*" + AFECTACION + r"\b"
)


def normalizar(valor):
    if valor is None or (not isinstance(valor, (dict, list, tuple)) and pd.isna(valor)):
        return ""
    return " ".join("".join(c for c in unicodedata.normalize("NFKD", str(valor).casefold())
                            if not unicodedata.combining(c)).split())


def analizar_alertamiento(row):
    campos = [(campo, normalizar(row.get(campo))) for campo in CAMPOS_EVIDENCIA]
    alerta = any(PATRON_ALERTA.search(texto) for _, texto in campos)
    descartes = []
    afectaciones = []
    for campo, texto in campos:
        descarte = PATRON_DESCARTE.search(texto)
        if descarte:
            descartes.append(f"{campo}: {descarte.group()}")
        evidencia = PATRON_AFECTACION.search(PATRON_NO_CONFIRMADO.sub("", texto))
        if evidencia:
            afectaciones.append(f"{campo}: {evidencia.group()}")
    return {
        "es_alertamiento": alerta,
        "evidencia_afectacion": "; ".join(afectaciones),
        "evidencia_descarte": "; ".join(descartes),
    }
