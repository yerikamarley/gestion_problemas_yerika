import json
import math
import re
import sys
import warnings
from datetime import datetime

import pandas as pd


CATEGORIES = [
    {
        "tipificacion": "8 - Instalaciones",
        "descripcion": "Procesos de instalacion, activacion, agendamiento o citas con tecnicos de instalacion.",
    },
    {
        "tipificacion": "2 - Soporte Uso",
        "descripcion": "Dudas de uso, acompanamiento funcional, configuracion, orientacion y paso a paso.",
    },
    {
        "tipificacion": "5 - incidente",
        "descripcion": "Casos marcados como incidente o con afectacion operativa reportada.",
    },
    {
        "tipificacion": "7 - No Aplica",
        "descripcion": "Casos sin informacion suficiente o que no encajan en las reglas definidas.",
    },
    {
        "tipificacion": "3 - Soporte Falla",
        "descripcion": "Errores, fallas, caidas, lentitud, indisponibilidad o afectaciones tecnicas.",
    },
    {
        "tipificacion": "4 - solicitudes",
        "descripcion": "Solicitudes operativas o comerciales como certificados, biometria, pagos u otros tramites.",
    },
    {
        "tipificacion": "6 - Plataformas Ext",
        "descripcion": "Problemas relacionados con plataformas externas como Adobe, Autofirma o DocuSign.",
    },
    {
        "tipificacion": "1 - phishing",
        "descripcion": "Correos sospechosos, suplantacion, enlaces fraudulentos o reportes de phishing.",
    },
]


DESCRIPTION_BY_CATEGORY = {item["tipificacion"]: item["descripcion"] for item in CATEGORIES}

MANUAL_OVERRIDES = {
    "CS0180913": ("3 - Soporte Falla", "Inconveniente FVC asociado a bloqueo/rechazo en validacion."),
    "CS0180986": ("3 - Soporte Falla", "Inconveniente FVC con bloqueos por intentos fallidos."),
    "CS0181629": ("4 - solicitudes", "Solicitud de desbloqueo FVC."),
    "CS0182271": ("4 - solicitudes", "Solicitud de desbloqueo."),
    "CS0182608": ("3 - Soporte Falla", "Error de biometria y preguntas sociodemograficas."),
    "CS0182090": ("8 - Instalaciones", "Validacion de acceso y condiciones para instalacion/uso de token virtual."),
    "CS0182037": ("3 - Soporte Falla", "Error en recepcion de codigos OTP."),
    "CS0184362": ("7 - No Aplica", "Caso anulado por inactividad y solicitud de informacion insuficiente."),
    "CS0184391": ("2 - Soporte Uso", "Solicitud de sesion y acompanamiento para validar aplicativo."),
    "CS0185953": ("3 - Soporte Falla", "Bloqueo reportado y ajuste interno posterior."),
    "CS0188636": ("3 - Soporte Falla", "Error biometrico reportado, cerrado por falta de respuesta."),
}


KEYWORD_RULES = [
    ("1 - phishing", r"phishing|suplant|fraud|correo sospechoso|enlace fraud"),
    ("6 - Plataformas Ext", r"adobe|autofirma|docusign"),
    ("5 - incidente", r"incidente|caida|indispon|afectacion operativa|masivo"),
    ("8 - Instalaciones", r"instal|activaci[oó]n|agenda|cita|tecnico|token virtual"),
    ("3 - Soporte Falla", r"error|falla|inconveniente|bloqueo|desbloqueo|otp|biometri|fvc|lentitud"),
    ("4 - solicitudes", r"solicitud|certificado|biometria|pago|tramite|descargar|formato"),
    ("2 - Soporte Uso", r"duda|acompan|configur|orient|paso a paso|como|uso|acceso|validacion"),
]


def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value)


def normalize_text(value):
    text = clean_text(value).lower()
    replacements = str.maketrans("áéíóúüñ", "aeiouun")
    return text.translate(replacements)


def classify(row):
    number = clean_text(row.get("Número"))
    if number in MANUAL_OVERRIDES:
        category, reason = MANUAL_OVERRIDES[number]
        return category, DESCRIPTION_BY_CATEGORY[category], reason

    text_fields = [
        "Breve descripción",
        "Código de resolución",
        "Producto",
        "Causa",
        "Notas de resolución",
        "Observaciones adicionales",
        "Observaciones y notas de trabajo",
    ]
    combined = " ".join(normalize_text(row.get(field)) for field in text_fields)

    if "anulado" in normalize_text(row.get("Código de resolución")) and len(combined.strip()) < 120:
        category = "7 - No Aplica"
        return category, DESCRIPTION_BY_CATEGORY[category], "Caso anulado con informacion limitada."

    for category, pattern in KEYWORD_RULES:
        if re.search(pattern, combined):
            return category, DESCRIPTION_BY_CATEGORY[category], f"Regla por palabras clave: {pattern}"

    category = "7 - No Aplica"
    return category, DESCRIPTION_BY_CATEGORY[category], "No se identifico una regla aplicable."


def parse_datetime(value):
    if pd.isna(value):
        return None
    parsed = pd.to_datetime(value, errors="coerce")
    if pd.isna(parsed):
        return None
    return parsed.to_pydatetime()


def resolution_bucket(hours):
    if hours is None:
        return "Sin cierre"
    if hours <= 24:
        return "<=24h"
    if hours <= 48:
        return "24-48h"
    if hours <= 72:
        return "48-72h"
    if hours <= 120:
        return "3-5 dias"
    return ">5 dias"


def json_value(value):
    if value is None:
        return None
    if isinstance(value, float) and math.isnan(value):
        return None
    if pd.isna(value):
        return None
    if isinstance(value, (pd.Timestamp, datetime)):
        return value.isoformat()
    return value


def main():
    if len(sys.argv) != 3:
        raise SystemExit("Usage: prepare_case_data.py <input.xlsx> <output.json>")

    input_path = sys.argv[1]
    output_path = sys.argv[2]

    warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
    df = pd.read_excel(input_path, sheet_name="Page 1")

    rows = []
    for _, row in df.iterrows():
        record = {col: json_value(row.get(col)) for col in df.columns}
        category, description, reason = classify(row)

        created = parse_datetime(row.get("Creado"))
        closed = parse_datetime(row.get("Cerrado"))
        hours = None
        days = None
        if created and closed:
            hours = round((closed - created).total_seconds() / 3600, 2)
            days = round(hours / 24, 2)

        result_code = normalize_text(row.get("Código de resolución"))
        result = "Anulado" if "anulado" in result_code else "Resuelto"

        record["_tipificacion"] = category
        record["_tipificacion_descripcion"] = description
        record["_criterio_tipificacion"] = reason
        record["_tiempo_resolucion_horas"] = hours
        record["_tiempo_resolucion_dias"] = days
        record["_rango_tiempo"] = resolution_bucket(hours)
        record["_mes_creacion"] = created.strftime("%Y-%m") if created else "Sin fecha"
        record["_resultado"] = result
        record["_cumplimiento_48h"] = "Sin fecha" if hours is None else ("<=48h" if hours <= 48 else ">48h")
        rows.append(record)

    payload = {
        "source": input_path,
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "columns": list(df.columns),
        "categories": CATEGORIES,
        "rows": rows,
    }

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    print(f"Wrote {len(rows)} rows to {output_path}")


if __name__ == "__main__":
    main()
