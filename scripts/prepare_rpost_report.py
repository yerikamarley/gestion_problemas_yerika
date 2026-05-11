from __future__ import annotations

import json
import re
import statistics
import unicodedata
from collections import Counter
from pathlib import Path

import pandas as pd


SOURCE = Path(r"C:\Users\yerik\Downloads\incident (63).xlsx")
OUT_DIR = Path(r"C:\Users\yerik\OneDrive\Desktop\gestion_problemas_yerika\outputs\rpost_incidentes")
OUT_JSON = OUT_DIR / "rpost_report_data.json"


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
    "INC0017032": "Se confirmo que el servicio estuvo operativo pese a la alerta inicial.",
    "INC0017040": "Alarma de caida sobre portal.rpost.com y WBLM en Agrio reportado Down.",
    "INC0017112": "Alarma de caida sobre portal RPOST en Agrio; escalamiento proveedor ticket 60054.",
    "INC0017041": "Error 550 5.7.515 Access denied hacia dominios Hotmail; revision DKIM con proveedor.",
    "INC0017088": "Aplicacion portal.rpost.com en Agrio reportada Down; caso proveedor 59743.",
    "INC0017124": "PORTAL.RPOST.COM - WBLM en nodo Agrio reportado Down.",
    "INC0017093": "Application portal.rpost.com / WBLM en nodo Agrio reportada Down; caso proveedor 59820.",
    "INC0017125": "Alarma de caida sobre portal RPOST / WBLM; cierre sin RCA tecnica detallada.",
    "INC0017129": "PORTAL.RPOST.COM - WBLM y componente asociado en estado unknown; caso proveedor 60183.",
    "INC0017150": "Servicios portal.rpost.com / WBLM en Agrio reportados Down; caso proveedor 60305.",
    "INC0017151": "Alarma de caida sobre portal RPOST; ticket proveedor 60306.",
    "INC0017163": "Pruebas de ping, navegador, HTTP 200 y monitoreo UP confirmaron disponibilidad.",
    "INC0017162": "Prueba ICMP estable y retorno a normal; evento intermitente/transitorio.",
    "INC0017069": "Servicio reportado Down; se escalo para validar indisponibilidad momentanea, caso 59483.",
    "INC0017207": "Caida reportada el 1 de mayo sobre PORTAL.RPOST.COM - WBLM; ticket proveedor 60836.",
}


OUR_OBS_BY_CASE = {
    "INC0017032": "Cerrar como falso positivo y conservar evidencias de consumo exitoso para depurar monitoreo.",
    "INC0017040": "Validar discrepancia entre duracion oficial y tiempo calendario; documentar hora real de recuperacion.",
    "INC0017112": "Asociar el ticket 60054 al cierre y pedir RCA formal del proveedor.",
    "INC0017041": "Separar este caso de disponibilidad; tratarlo como configuracion/autenticacion de correo DKIM.",
    "INC0017088": "Relacionar caso 59743 y registrar evidencia de afectacion real vs alarma.",
    "INC0017124": "Solicitar confirmacion tecnica porque no hay ticket proveedor visible en el registro.",
    "INC0017093": "Mantener trazabilidad con caso 59820 y exigir causa de la caida WBLM.",
    "INC0017125": "Corregir/validar duracion, porque el campo oficial no coincide con Creado/Cerrado.",
    "INC0017129": "Registrar si el estado unknown tuvo impacto real de usuario o fue falla de monitoreo.",
    "INC0017150": "Caso de larga duracion; requiere RCA y acciones preventivas documentadas.",
    "INC0017151": "Validar discrepancia de duracion y obtener detalle del ticket 60306.",
    "INC0017163": "Clasificar como alerta sin indisponibilidad confirmada; ajustar umbrales o salud del monitor.",
    "INC0017162": "Clasificar como intermitencia; anexar pruebas de conectividad y retorno a normalidad.",
    "INC0017069": "Documentar la ventana exacta de indisponibilidad momentanea y evidencia de recuperacion.",
    "INC0017207": "Pedir explicacion del evento del 1 de mayo y confirmar si hubo mantenimiento programado.",
}


PROVIDER_OBS_BY_CASE = {
    "INC0017032": "Confirmar si el endpoint monitoreado tuvo degradacion real o si fue falso positivo de monitoreo.",
    "INC0017040": "Entregar RCA, linea de tiempo y confirmacion de si existio mantenimiento o cambio programado.",
    "INC0017112": "Para ticket 60054, informar causa, hora de inicio/fin, impacto y accion correctiva.",
    "INC0017041": "Confirmar llave DKIM correcta, evidenciar valor esperado y guia de implementacion para el cliente.",
    "INC0017088": "Para caso 59743, entregar causa del Down en Agrio/WBLM y acciones preventivas.",
    "INC0017124": "Confirmar causa del Down y si hubo evento programado no informado.",
    "INC0017093": "Para caso 59820, entregar RCA del Down en portal RPOST/WBLM.",
    "INC0017125": "Confirmar si el evento fue indisponibilidad real o condicion temporal de monitoreo.",
    "INC0017129": "Para caso 60183, explicar estado unknown y su impacto sobre el portal.",
    "INC0017150": "Para caso 60305, entregar RCA detallado por la larga duracion registrada.",
    "INC0017151": "Para ticket 60306, confirmar causa y si coincide con otros eventos de abril.",
    "INC0017163": "Revisar salud del endpoint monitoreado; las pruebas locales mostraron servicio disponible.",
    "INC0017162": "Confirmar causa de intermitencia y si hubo degradacion breve del portal.",
    "INC0017069": "Para caso 59483, confirmar indisponibilidad momentanea, causa y duracion real.",
    "INC0017207": "Para ticket 60836, confirmar evento del 1 de mayo, causa y existencia/no existencia de mantenimiento.",
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


def get_col_map(df: pd.DataFrame) -> dict[str, str]:
    return {norm(col): col for col in df.columns}


def get_value(row: pd.Series, col_map: dict[str, str], key: str) -> object:
    col = col_map.get(key)
    if col is None:
        return ""
    value = row.get(col, "")
    if pd.isna(value):
        return ""
    return value


def clean_snippet(text: str, limit: int = 260) -> str:
    text = re.sub(r"https?://\S+", "[url]", text)
    text = re.sub(r"\s+", " ", text).strip()
    if len(text) > limit:
        return text[: limit - 3] + "..."
    return text


def event_group(event_type: str) -> str:
    lowered = event_type.lower()
    if "alerta" in lowered or "falso" in lowered:
        return "Alerta/no confirmada"
    if "funcional" in lowered:
        return "Funcional/correo"
    return "Indisponibilidad/probable"


def maint_status(full_text: str) -> tuple[str, str]:
    pattern = re.compile(r"mantenimiento|ventana|programad|scheduled|maintenance|cambio|change|RFC|crq", re.I)
    if pattern.search(full_text):
        return "Mencionada", "Revisar texto del ticket; hay mencion a mantenimiento/cambio."
    return "No evidenciada", "No aparece ventana de mantenimiento documentada en el ticket."


def main() -> None:
    df = pd.read_excel(SOURCE, sheet_name="Page 1", dtype=object)
    col_map = get_col_map(df)
    mask = pd.Series(False, index=df.index)
    for col in df.columns:
        mask |= df[col].astype(str).str.contains("rpost", case=False, na=False)
    rpost = df[mask].copy()

    detail_rows: list[dict[str, object]] = []
    for _, row in rpost.iterrows():
        numero = str(get_value(row, col_map, "numero"))
        creado = pd.to_datetime(get_value(row, col_map, "creado"), errors="coerce")
        cerrado = pd.to_datetime(get_value(row, col_map, "cerrado"), errors="coerce")
        dur_raw = get_value(row, col_map, "duracion")
        dur_secs = int(float(dur_raw)) if str(dur_raw).strip() else None
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
        provider_cases = sorted(
            set(re.findall(r"(?:caso|ticket|#)\s*#?\s*(\d{4,6})", full_text, flags=re.I))
        )
        maintenance, maintenance_detail = maint_status(full_text)
        event_type, root_cause = CLASSIFICATION.get(numero, ("Sin clasificar", "Pendiente de revision."))
        discrepancy = ""
        if dur_secs is not None and elapsed_secs is not None and abs(dur_secs - elapsed_secs) > 3600:
            discrepancy = "Revisar: duracion oficial difiere de Creado/Cerrado."

        detail_rows.append(
            {
                "Numero": numero,
                "Servicio": str(get_value(row, col_map, "servicio de negocio")),
                "Categoria": str(get_value(row, col_map, "categoria")),
                "Prioridad": str(get_value(row, col_map, "prioridad")),
                "Impacto": str(get_value(row, col_map, "impacto")),
                "Estado": str(get_value(row, col_map, "estado")),
                "Tipo falla ServiceNow": str(get_value(row, col_map, "tipo de falla")),
                "Creado": "" if pd.isna(creado) else creado.strftime("%Y-%m-%d %H:%M:%S"),
                "Cerrado": "" if pd.isna(cerrado) else cerrado.strftime("%Y-%m-%d %H:%M:%S"),
                "Duracion oficial seg": dur_secs,
                "Duracion oficial": seconds_to_human(dur_secs),
                "Tiempo calendario": seconds_to_human(elapsed_secs),
                "Diferencia duracion": discrepancy,
                "Caso proveedor": ", ".join(provider_cases) if provider_cases else "No registrado",
                "Tipo evento": event_type,
                "Grupo evento": event_group(event_type),
                "Causa raiz / causa probable": root_cause,
                "Evidencia relevante": EVIDENCE.get(numero, clean_snippet(full_text)),
                "Ventana mantenimiento": maintenance,
                "Detalle ventana mantenimiento": maintenance_detail,
                "Observacion nosotros": OUR_OBS_BY_CASE.get(numero, ""),
                "Observacion proveedor": PROVIDER_OBS_BY_CASE.get(numero, ""),
                "Breve descripcion": str(get_value(row, col_map, "breve descripcion")),
                "Actualizaciones": str(get_value(row, col_map, "actualizaciones")),
            }
        )

    total_duration = sum(row["Duracion oficial seg"] or 0 for row in detail_rows)
    valid_duration = [row["Duracion oficial seg"] for row in detail_rows if row["Duracion oficial seg"] is not None]
    provider_ticket_count = sum(1 for row in detail_rows if row["Caso proveedor"] != "No registrado")
    maintenance_count = sum(1 for row in detail_rows if row["Ventana mantenimiento"] == "Mencionada")
    discrepancy_count = sum(1 for row in detail_rows if row["Diferencia duracion"])

    summary = {
        "source": str(SOURCE),
        "sheet": "Page 1",
        "generated": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"),
        "total_source_rows": int(len(df)),
        "total_rpost": len(detail_rows),
        "provider_ticket_count": provider_ticket_count,
        "maintenance_count": maintenance_count,
        "duration_total": seconds_to_human(total_duration),
        "duration_avg": seconds_to_human(sum(valid_duration) / len(valid_duration)) if valid_duration else "",
        "duration_median": seconds_to_human(statistics.median(valid_duration)) if valid_duration else "",
        "discrepancy_count": discrepancy_count,
    }

    def counter_rows(key: str) -> list[dict[str, object]]:
        counts = Counter(str(row[key]) for row in detail_rows)
        return [{"Categoria": name, "Cantidad": count} for name, count in counts.most_common()]

    month_counts = Counter(row["Creado"][:7] for row in detail_rows if row["Creado"])
    by_month = [{"Mes": month, "Cantidad": count} for month, count in sorted(month_counts.items())]

    observations_us = [
        {
            "Tema": "Cierre y RCA",
            "Observacion": "Evitar cierres genericos como Cerrado/resuelto por solicitante cuando hubo alarma Down/unknown/time out.",
            "Accion sugerida": "Agregar campo obligatorio de causa raiz tecnica y evidencia de recuperacion.",
        },
        {
            "Tema": "Clasificacion operativa",
            "Observacion": "Separar alertas sin indisponibilidad, intermitencias, indisponibilidad real y casos funcionales de correo.",
            "Accion sugerida": "Usar una taxonomia unica para reportes mensuales y postmortem.",
        },
        {
            "Tema": "Duracion",
            "Observacion": f"Hay {discrepancy_count} casos con diferencia relevante entre Duracion oficial y Creado/Cerrado.",
            "Accion sugerida": "Confirmar si Duracion representa tiempo SLA/operativo o tiempo calendario.",
        },
        {
            "Tema": "Servicio de negocio",
            "Observacion": "RPOST aparece distribuido en Certificacion Digital, Infraestructura y CiberSeguridad.",
            "Accion sugerida": "Crear etiqueta unica RPOST Portal/WBLM para no fragmentar metricas.",
        },
        {
            "Tema": "Ventanas de mantenimiento",
            "Observacion": "No se encontro evidencia documental de ventanas de mantenimiento en los tickets RPOST.",
            "Accion sugerida": "Solicitar al proveedor confirmacion explicita de mantenimiento/no mantenimiento por cada evento escalado.",
        },
    ]

    observations_provider = [
        {
            "Solicitud": "RCA por ticket proveedor",
            "Detalle": "Entregar causa raiz, linea de tiempo, impacto y accion preventiva para tickets 59483, 59743, 59820, 60054, 60183, 60305, 60306 y 60836.",
            "Prioridad": "Alta",
        },
        {
            "Solicitud": "Confirmar ventanas de mantenimiento",
            "Detalle": "Indicar para cada evento si existio ventana programada, cambio, despliegue o mantenimiento no informado.",
            "Prioridad": "Alta",
        },
        {
            "Solicitud": "Eventos WBLM/Agrio",
            "Detalle": "Explicar recurrencia de estados Down/unknown/time out sobre PORTAL.RPOST.COM - WBLM en nodo Agrio.",
            "Prioridad": "Alta",
        },
        {
            "Solicitud": "Monitoreo",
            "Detalle": "Validar endpoints, umbrales y criterios de salud para reducir falsos positivos e intermitencias sin impacto.",
            "Prioridad": "Media",
        },
        {
            "Solicitud": "Caso DKIM",
            "Detalle": "Confirmar llave DKIM correcta, valor esperado, responsable de implementacion y evidencia de validacion.",
            "Prioridad": "Media",
        },
    ]

    maintenance_rows = [
        {
            "Numero": row["Numero"],
            "Creado": row["Creado"],
            "Caso proveedor": row["Caso proveedor"],
            "Ventana mantenimiento": row["Ventana mantenimiento"],
            "Conclusion": row["Detalle ventana mantenimiento"],
            "Requiere confirmacion proveedor": "Si" if row["Caso proveedor"] != "No registrado" else "Recomendado",
        }
        for row in detail_rows
    ]

    output = {
        "summary": summary,
        "detail": detail_rows,
        "by_group": counter_rows("Grupo evento"),
        "by_type": counter_rows("Tipo evento"),
        "by_service": counter_rows("Servicio"),
        "by_failure": counter_rows("Tipo falla ServiceNow"),
        "by_month": by_month,
        "observations_us": observations_us,
        "observations_provider": observations_provider,
        "maintenance": maintenance_rows,
    }

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    OUT_JSON.write_text(json.dumps(output, ensure_ascii=False, indent=2), encoding="utf-8")
    print(OUT_JSON)


if __name__ == "__main__":
    main()
