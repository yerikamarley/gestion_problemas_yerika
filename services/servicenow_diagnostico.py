"""Errores accionables sin imprimir credenciales, cuerpos de mensajes ni valores SQL."""

import logging
from pathlib import Path
import re
import traceback
from uuid import uuid4

from services.servicenow import ServiceNowError

LOGGER = logging.getLogger(__name__)
MENSAJES_DB = {
    "26000": "La conexión a la base perdió una consulta preparada.",
    "42P05": "La conexión a la base encontró una consulta preparada duplicada.",
    "55P03": "La base de datos está ocupada guardando otros cambios.",
    "57014": "La base de datos canceló una operación que tardó demasiado.",
    "25P03": "Se perdió el turno de sincronización por inactividad de la conexión.",
    "23505": "La base de datos rechazó un identificador duplicado.",
    "42P01": "Falta una tabla requerida en la base de datos.",
    "42703": "Falta una columna requerida en la base de datos.",
    "22P02": "La base de datos rechazó el formato de un campo.",
    "22P05": "La base de datos rechazó un carácter no admitido.",
    "22021": "La base de datos rechazó la codificación de un campo.",
}


def diagnosticar_fallo(error, *, etapa, guardados=0, total=None):
    referencia = uuid4().hex[:10]
    estado = getattr(error, "sqlstate", None) or getattr(error, "pgcode", None)
    estado = estado if isinstance(estado, str) and re.fullmatch(r"[0-9A-Z]{5}", estado) else "sin_codigo"
    frames = " > ".join(f"{Path(f.filename).name}:{f.lineno} ({f.name})"
                        for f in traceback.extract_tb(error.__traceback__))
    LOGGER.error("SN_SYNC ref=%s etapa=%s guardados=%d total=%s excepcion=%s sqlstate=%s stack=%s",
                 referencia, etapa, guardados, total, type(error).__name__, estado, frames)
    razon = MENSAJES_DB.get(estado, "El servidor no pudo completar una operación de sincronización.")
    cantidad = f"{guardados} de {total}" if isinstance(total, int) else str(guardados)
    codigo = f" Código de base: {estado}." if estado != "sin_codigo" else ""
    return ServiceNowError(f"{razon} Etapa: {etapa}. Casos guardados: {cantidad}.{codigo} Referencia: {referencia}.")


def cerrar_control(control):
    try:
        control.rollback()
    except Exception as error:
        # No ocultar el error original si el servidor ya cerró esta conexión.
        diagnosticar_fallo(error, etapa="liberar el turno")
    finally:
        try:
            control.close()
        except Exception as error:
            diagnosticar_fallo(error, etapa="cerrar la conexión")
