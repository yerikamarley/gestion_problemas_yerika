from __future__ import annotations

import json
import os
import re
import statistics
import unicodedata
from collections import Counter, defaultdict
from pathlib import Path

import pandas as pd


PROJECT_DIR = Path(__file__).resolve().parents[1]
DOWNLOADS_DIR = Path.home() / "Downloads"
INCIDENT_SOURCE = Path(os.environ.get("RPOST_INCIDENT_SOURCE", DOWNLOADS_DIR / "incident (63).xlsx"))
ANS_SOURCE = Path(os.environ.get("ANS_RPOST_SOURCE", DOWNLOADS_DIR / "ANS_RPOST_con_Dashboard.xlsx"))
OUT_DIR = PROJECT_DIR / "outputs" / "ans_rpost"
OUT_JSON = OUT_DIR / "ans_rpost_analysis.json"

ACUSE_CASE = {
    "Caso": "CS0196170",
    "PQRS": "PQRF-2026-1902",
    "Cliente": "Empresa de Energia de Pereira S.A. E.S.P.",
    "Servicio": "Certimail",
    "Tipo": "ELECTRONICO",
    "Fecha reporte": "2026-05-07 15:12:00",
    "Fecha asignacion": "2026-05-07 15:18:00",
    "Fecha escalamiento proveedor": "2026-05-07 22:44:00",
    "Fechas afectadas": "2026-05-05 y 2026-05-06",
    "Mes ANS": "2026-05",
    "Cantidad acuses afectados": 217,
    "Rango ANS": "101 a 250 en el mes",
    "Compensacion acuses": "5%",
    "Aplica compensacion": "SI",
    "Base ANS": "Acuses no generados oportunamente entre 101 y 250 en el mes.",
    "Observacion": "No aplica doble descuento si el evento ya fue penalizado por disponibilidad. En mayo, la disponibilidad calculada cumple, por lo que el reclamo de acuses queda como frente independiente salvo evidencia contraria.",
}


CLASSIFICATION = {
    "INC0017032": (
        "Alerta / falso positivo",
        "Monitoreo alerto, pero se validaron consumos y el servicio estuvo operativo.",
    ),
    "INC0017040": (
        "Indisponibilidad probable",
        "Monitoreo reporto portal.rpost.com / PORTAL.RPOST.COM - WBLM en estado Down.",
    ),
    "INC0017112": (
        "Indisponibilidad probable",
        "Alarma de caida del portal RPOST; se escalo al proveedor con ticket 60054.",
    ),
    "INC0017041": (
        "Incidencia funcional / correo",
        "Rechazo 550 5.7.515 hacia Hotmail; evidencia apunta a implementacion DKIM diferente a la llave provista por RPost.",
    ),
    "INC0017088": (
        "Indisponibilidad probable",
        "Aplicacion portal.rpost.com en Agrio reportada Down; escalado proveedor caso 59743.",
    ),
    "INC0017124": (
        "Indisponibilidad probable",
        "PORTAL.RPOST.COM - WBLM en nodo Agrio reportado Down.",
    ),
    "INC0017093": (
        "Indisponibilidad probable",
        "Portal RPOST.COM / WBLM reportado Down; escalado proveedor caso 59820.",
    ),
    "INC0017125": (
        "Indisponibilidad probable",
        "Portal RPOST reportado en alarma/estado WBLM; evidencia de caida en monitoreo.",
    ),
    "INC0017129": (
        "Indisponibilidad probable",
        "PORTAL.RPOST.COM - WBLM en estado unknown; escalado proveedor caso 60183.",
    ),
    "INC0017150": (
        "Indisponibilidad probable",
        "Servicios portal.rpost.com / WBLM reportados Down; escalado proveedor caso 60305.",
    ),
    "INC0017151": (
        "Indisponibilidad probable",
        "Alarma de caida del portal RPOST; escalado proveedor ticket 60306.",
    ),
    "INC0017163": (
        "Alerta / sin indisponibilidad confirmada",
        "Ping OK, portal cargo login, HTTP 200 y servidor UP.",
    ),
    "INC0017162": (
        "Alerta transitoria / intermitencia",
        "Monitoreo alerto caida, pruebas IP sin falla y retorno a normal; sugiere evento intermitente/transitorio.",
    ),
    "INC0017069": (
        "Indisponibilidad momentanea",
        "Alarma de servicio Down y escalamiento para validar indisponibilidad momentanea, caso 59483.",
    ),
    "INC0017207": (
        "Indisponibilidad reportada",
        "Caida del 1 de mayo del portal/WBLM; escalado proveedor ticket 60836.",
    ),
}


EVIDENCE = {
    "INC0017032": "Servicio operativo pese a alerta inicial; validado con consumos.",
    "INC0017040": "portal.rpost.com y PORTAL.RPOST.COM - WBLM en Agrio reportados Down.",
    "INC0017112": "Alarma de caida sobre portal RPOST en Agrio; ticket proveedor 60054.",
    "INC0017041": "Error 550 5.7.515 hacia Hotmail; revision DKIM con proveedor.",
    "INC0017088": "portal.rpost.com en Agrio reportado Down; caso proveedor 59743.",
    "INC0017124": "PORTAL.RPOST.COM - WBLM en nodo Agrio reportado Down.",
    "INC0017093": "portal.rpost.com / WBLM en nodo Agrio reportado Down; caso 59820.",
    "INC0017125": "Alarma de caida sobre portal RPOST / WBLM; cierre sin RCA tecnica detallada.",
    "INC0017129": "PORTAL.RPOST.COM - WBLM y componente asociado en estado unknown; caso 60183.",
    "INC0017150": "Servicios portal.rpost.com / WBLM en Agrio reportados Down; caso 60305.",
    "INC0017151": "Alarma de caida sobre portal RPOST; ticket proveedor 60306.",
    "INC0017163": "Ping, navegador, HTTP 200 y monitoreo UP confirmaron disponibilidad.",
    "INC0017162": "Prueba ICMP estable y retorno a normal; evento intermitente/transitorio.",
    "INC0017069": "Servicio reportado Down; se escalo para validar indisponibilidad momentanea, caso 59483.",
    "INC0017207": "Caida reportada el 1 de mayo sobre PORTAL.RPOST.COM - WBLM; ticket 60836.",
}


def norm(value: object) -> str:
    text = unicodedata.normalize("NFKD", str(value))
    return "".join(ch for ch in text if not unicodedata.combining(ch)).lower().strip()


def seconds_to_human(value: float | int | None) -> str:
    if value is None:
        return ""
    seconds = int(round(float(value)))
    days, rem = divmod(seconds, 86400)
    hours, rem = divmod(rem, 3600)
    minutes, secs = divmod(rem, 60)
    parts: list[str] = []
    if days:
        parts.append(f"{days}d")
    if hours:
        parts.append(f"{hours}h")
    if minutes:
        parts.append(f"{minutes}m")
    if secs or not parts:
        parts.append(f"{secs}s")
    return " ".join(parts)


def get_value(row: pd.Series, col_map: dict[str, str], key: str) -> object:
    col = col_map.get(key)
    if col is None:
        return ""
    value = row.get(col, "")
    if pd.isna(value):
        return ""
    return value


def parse_duration(value: object) -> int | None:
    if value is None or pd.isna(value) or str(value).strip() == "":
        return None
    try:
        return int(float(value))
    except ValueError:
        return None


def maintenance_status(full_text: str) -> tuple[str, str]:
    pattern = re.compile(r"mantenimiento|ventana|programad|scheduled|maintenance|cambio|change|RFC|crq", re.I)
    if pattern.search(full_text):
        return "Mencionada", "Hay mencion a mantenimiento/cambio; revisar detalle del ticket."
    return "No evidenciada", "No aparece ventana de mantenimiento documentada en el ticket."


def event_group(event_type: str) -> str:
    lowered = event_type.lower()
    if "alerta" in lowered or "falso" in lowered:
        return "Alerta/no confirmada"
    if "funcional" in lowered:
        return "Funcional/correo"
    return "Indisponibilidad/probable"


def solution_target(priority: str) -> tuple[str, int | None]:
    lowered = priority.lower()
    if lowered.startswith("1") or "critica" in lowered or "alta" in lowered:
        return "Alta", 8 * 3600
    if lowered.startswith("2"):
        return "Alta", 8 * 3600
    if lowered.startswith("3") or "moderada" in lowered or "media" in lowered:
        return "Media", 16 * 3600
    if lowered.startswith("4") or "baja" in lowered:
        return "Baja", 32 * 3600
    return "No definida", None


def month_seconds(month: str) -> int:
    start = pd.Timestamp(f"{month}-01")
    end = start + pd.offsets.MonthBegin(1)
    return int((end - start).total_seconds())


def availability(total_seconds: int, downtime_seconds: int) -> float:
    return (total_seconds - downtime_seconds) / total_seconds * 100


def main() -> None:
    ans_df = pd.read_excel(ANS_SOURCE, sheet_name="ANS RPOST", dtype=object).fillna("")
    incidents_df = pd.read_excel(INCIDENT_SOURCE, sheet_name="Page 1", dtype=object)
    col_map = {norm(col): col for col in incidents_df.columns}

    rpost_mask = pd.Series(False, index=incidents_df.index)
    for col in incidents_df.columns:
        rpost_mask |= incidents_df[col].astype(str).str.contains("rpost", case=False, na=False)
    rpost = incidents_df[rpost_mask].copy()

    detail: list[dict[str, object]] = []
    for _, row in rpost.iterrows():
        numero = str(get_value(row, col_map, "numero"))
        creado = pd.to_datetime(get_value(row, col_map, "creado"), errors="coerce")
        cerrado = pd.to_datetime(get_value(row, col_map, "cerrado"), errors="coerce")
        dur_secs = parse_duration(get_value(row, col_map, "duracion"))
        elapsed_secs = None
        if pd.notna(creado) and pd.notna(cerrado):
            elapsed_secs = int((cerrado - creado).total_seconds())

        text_fields = [
            str(get_value(row, col_map, "breve descripcion")),
            str(get_value(row, col_map, "descripcion")),
            str(get_value(row, col_map, "observaciones y notas de trabajo")),
            str(get_value(row, col_map, "observaciones adicionales")),
            str(get_value(row, col_map, "lista de notas de trabajo")),
        ]
        full_text = "\n".join(text_fields)
        cases = sorted(set(re.findall(r"(?:caso|ticket|#)\s*#?\s*(\d{4,6})", full_text, flags=re.I)))
        maintenance, maintenance_detail = maintenance_status(full_text)
        event_type, root_cause = CLASSIFICATION.get(numero, ("Sin clasificar", "Pendiente de revision."))
        group = event_group(event_type)
        priority = str(get_value(row, col_map, "prioridad"))
        ans_priority, target_secs = solution_target(priority)

        counts_availability = group == "Indisponibilidad/probable" and maintenance == "No evidenciada"
        outside_solution = "No evaluable"
        solution_notes = "No hay duracion o prioridad suficiente para calcular."
        if dur_secs is not None and target_secs is not None:
            if dur_secs > target_secs:
                outside_solution = "Si"
                solution_notes = "Supera el umbral de solucion del ANS para la prioridad."
            else:
                outside_solution = "No"
                solution_notes = "Dentro del umbral de solucion segun duracion oficial."
        if group == "Alerta/no confirmada":
            outside_solution = "No aplica"
            solution_notes = "Alerta/no indisponibilidad confirmada; no deberia penalizar sin evidencia adicional."

        responsibility = "Proveedor probable" if counts_availability else "No concluyente"
        if numero == "INC0017041":
            responsibility = "Sujeto a validacion: evidencia sugiere implementacion DKIM del cliente"

        discrepancy = ""
        if dur_secs is not None and elapsed_secs is not None and abs(dur_secs - elapsed_secs) > 3600:
            discrepancy = "Revisar diferencia entre duracion oficial y Creado/Cerrado."

        detail.append(
            {
                "Numero": numero,
                "Mes": "" if pd.isna(creado) else creado.strftime("%Y-%m"),
                "Creado": "" if pd.isna(creado) else creado.strftime("%Y-%m-%d %H:%M:%S"),
                "Cerrado": "" if pd.isna(cerrado) else cerrado.strftime("%Y-%m-%d %H:%M:%S"),
                "Servicio": str(get_value(row, col_map, "servicio de negocio")),
                "Categoria": str(get_value(row, col_map, "categoria")),
                "Prioridad ServiceNow": priority,
                "Prioridad ANS": ans_priority,
                "Target solucion ANS": seconds_to_human(target_secs),
                "Duracion oficial seg": dur_secs,
                "Duracion oficial": seconds_to_human(dur_secs),
                "Tiempo calendario seg": elapsed_secs,
                "Tiempo calendario": seconds_to_human(elapsed_secs),
                "Diferencia duracion": discrepancy,
                "Caso proveedor": ", ".join(cases) if cases else "No registrado",
                "Tipo evento": event_type,
                "Grupo evento": group,
                "Cuenta disponibilidad ANS": "Si" if counts_availability else "No",
                "Fuera ANS solucion": outside_solution,
                "Notas solucion ANS": solution_notes,
                "Causa raiz / probable": root_cause,
                "Evidencia": EVIDENCE.get(numero, ""),
                "Ventana mantenimiento": maintenance,
                "Detalle mantenimiento": maintenance_detail,
                "Responsabilidad preliminar": responsibility,
                "Breve descripcion": str(get_value(row, col_map, "breve descripcion")),
            }
        )

    monthly: list[dict[str, object]] = []
    for month in sorted({row["Mes"] for row in detail if row["Mes"]}):
        total = month_seconds(month)
        allowed = int(total * 0.001)
        official_down = sum(
            row["Duracion oficial seg"] or 0
            for row in detail
            if row["Mes"] == month and row["Cuenta disponibilidad ANS"] == "Si"
        )
        calendar_down = sum(
            row["Tiempo calendario seg"] or 0
            for row in detail
            if row["Mes"] == month and row["Cuenta disponibilidad ANS"] == "Si"
        )
        availability_official = availability(total, official_down)
        availability_calendar = availability(total, calendar_down)
        monthly.append(
            {
                "Mes": month,
                "Segundos mes": total,
                "Minutos caida permitidos 99.9": round(allowed / 60, 2),
                "Indisponibilidad oficial": seconds_to_human(official_down),
                "Indisponibilidad oficial seg": official_down,
                "Disponibilidad calculada": round(availability_official, 5),
                "ANS disponibilidad >=99.9": "Cumple" if availability_official >= 99.9 else "Afectado",
                "Indisponibilidad calendario": seconds_to_human(calendar_down),
                "Disponibilidad calendario": round(availability_calendar, 5),
                "ANS calendario >=99.9": "Cumple" if availability_calendar >= 99.9 else "Afectado",
                "Exceso sobre permitido": seconds_to_human(max(official_down - allowed, 0)),
                "Casos que cuentan disponibilidad": sum(
                    1 for row in detail if row["Mes"] == month and row["Cuenta disponibilidad ANS"] == "Si"
                ),
                "Ventanas mantenimiento evidenciadas": sum(
                    1 for row in detail if row["Mes"] == month and row["Ventana mantenimiento"] == "Mencionada"
                ),
            }
        )

    solution_by_month: list[dict[str, object]] = []
    for month in sorted({row["Mes"] for row in detail if row["Mes"]}):
        outside = [
            row
            for row in detail
            if row["Mes"] == month
            and row["Fuera ANS solucion"] == "Si"
            and row["Responsabilidad preliminar"] != "Sujeto a validacion: evidencia sugiere implementacion DKIM del cliente"
        ]
        conditional = [
            row
            for row in detail
            if row["Mes"] == month
            and row["Fuera ANS solucion"] == "Si"
            and row["Responsabilidad preliminar"] == "Sujeto a validacion: evidencia sugiere implementacion DKIM del cliente"
        ]
        comp_base = max(len(outside) - 1, 0) if len(outside) >= 2 else 0
        solution_by_month.append(
            {
                "Mes": month,
                "Incidentes fuera ANS solucion": len(outside),
                "Casos fuera ANS solucion": ", ".join(row["Numero"] for row in outside) or "Ninguno",
                "Casos condicionales": ", ".join(row["Numero"] for row in conditional) or "Ninguno",
                "Riesgo compensacion incidentes": f"{min(comp_base, 10)}%" if comp_base else "Sin compensacion automatica",
                "Nota": "El ANS indica 1% por incidente adicional, maximo 10%, cuando hay dos o mas incidentes fuera de ANS en el mes.",
            }
        )

    maintenance_rows = [
        {
            "Numero": row["Numero"],
            "Mes": row["Mes"],
            "Caso proveedor": row["Caso proveedor"],
            "Ventana mantenimiento": row["Ventana mantenimiento"],
            "Conclusion": row["Detalle mantenimiento"],
            "Impacto en ANS": "No se excluye del calculo" if row["Ventana mantenimiento"] == "No evidenciada" else "Validar exclusion",
            "Solicitud al proveedor": "Confirmar por escrito si existio mantenimiento, cambio o ventana acordada con 15 dias de anticipacion.",
        }
        for row in detail
    ]

    conclusions = [
        {
            "Tema": "Disponibilidad",
            "Conclusion": "Con la duracion oficial registrada, marzo y abril quedan por debajo del 99.9% de disponibilidad.",
            "Evidencia": "Los eventos RPOST/WBLM sin ventana de mantenimiento superan ampliamente el tiempo mensual permitido.",
            "Accion": "Solicitar RCA y confirmacion de no existencia de mantenimiento para marzo y abril.",
        },
        {
            "Tema": "Mayo",
            "Conclusion": "El evento de mayo no afecta el ANS de disponibilidad con la duracion registrada, pero si hay afectacion por acuses.",
            "Evidencia": "La indisponibilidad oficial fue 7m 31s, menor al umbral mensual de 99.9%; adicionalmente se reportan 217 acuses afectados.",
            "Accion": "Reclamar el frente de acuses como compensacion independiente, salvo que el proveedor pruebe doble penalizacion por disponibilidad.",
        },
        {
            "Tema": "Solucion de incidentes",
            "Conclusion": "Hay incidentes con duracion superior al target de solucion ANS de prioridad media.",
            "Evidencia": "Especialmente eventos largos de abril: INC0017124, INC0017129, INC0017150 e INC0017151.",
            "Accion": "Pedir al proveedor hora real de solucion y no solo fecha de cierre del caso.",
        },
        {
            "Tema": "Mantenimiento",
            "Conclusion": "No se evidencian ventanas de mantenimiento notificadas/acordadas en los tickets analizados.",
            "Evidencia": "Busqueda textual sin coincidencias de mantenimiento, ventana, scheduled, change, RFC o CRQ.",
            "Accion": "No excluir indisponibilidades del calculo hasta que el proveedor entregue soporte formal.",
        },
        {
            "Tema": "Limitaciones",
            "Conclusion": "No es posible evaluar acuse de envio/recibo por marca de tiempo individual, pero si se puede evaluar cantidad mensual de acuses no generados oportunamente.",
            "Evidencia": "El caso CS0196170 reporta 217 acuses afectados en mayo.",
            "Accion": "Solicitar al proveedor trazabilidad individual de generacion/certificacion de cada acuse.",
        },
        {
            "Tema": "Acuses",
            "Conclusion": "Con 217 acuses afectados en mayo, aplica compensacion del 5%.",
            "Evidencia": "El rango contractual de 101 a 250 acuses no generados oportunamente en el mes corresponde a 5%.",
            "Accion": "Incluir CS0196170 y PQRF-2026-1902 en el reclamo formal al proveedor.",
        },
    ]

    summary = {
        "ans_source": str(ANS_SOURCE),
        "incident_source": str(INCIDENT_SOURCE),
        "generated": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"),
        "total_incidents": len(detail),
        "months_affected_availability": ", ".join(
            row["Mes"] for row in monthly if row["ANS disponibilidad >=99.9"] == "Afectado"
        )
        or "Ninguno",
        "maintenance_mentions": sum(1 for row in detail if row["Ventana mantenimiento"] == "Mencionada"),
        "outside_solution_cases": sum(
            1
            for row in detail
            if row["Fuera ANS solucion"] == "Si"
            and row["Responsabilidad preliminar"] != "Sujeto a validacion: evidencia sugiere implementacion DKIM del cliente"
        ),
        "conditional_solution_cases": sum(
            1
            for row in detail
            if row["Fuera ANS solucion"] == "Si"
            and row["Responsabilidad preliminar"] == "Sujeto a validacion: evidencia sugiere implementacion DKIM del cliente"
        ),
        "acuses_affected": ACUSE_CASE["Cantidad acuses afectados"],
        "acuse_case": ACUSE_CASE["Caso"],
        "acuse_compensation": ACUSE_CASE["Compensacion acuses"],
        "availability_sla": ">= 99.9%",
        "solution_targets": "Alta 8h, Media 16h, Baja 32h",
    }

    output = {
        "summary": summary,
        "ans_rules": ans_df.astype(str).to_dict(orient="records"),
        "detail": detail,
        "monthly_availability": monthly,
        "solution_by_month": solution_by_month,
        "maintenance": maintenance_rows,
        "acuse_cases": [ACUSE_CASE],
        "conclusions": conclusions,
        "metrics": {
            "by_group": [{"Categoria": k, "Cantidad": v} for k, v in Counter(row["Grupo evento"] for row in detail).items()],
            "by_solution_status": [
                {"Estado": k, "Cantidad": v} for k, v in Counter(row["Fuera ANS solucion"] for row in detail).items()
            ],
        },
    }

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    OUT_JSON.write_text(json.dumps(output, ensure_ascii=False, indent=2), encoding="utf-8")
    print(OUT_JSON)


if __name__ == "__main__":
    main()
