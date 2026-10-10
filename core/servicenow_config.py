"""Persistencia privada de la conexión, independiente de las sesiones de usuarios."""

import os
import shutil
import tempfile
from threading import RLock

from dotenv import set_key

from core.settings import PROJECT_ROOT, config_value


_config_lock = RLock()
_CLAVES = {"SERVICENOW_USER", "SERVICENOW_PASSWORD", "SERVICENOW_SYNC_ACTOR", "SERVICENOW_SYNC_ENABLED"}
_OPCIONALES = {"SERVICENOW_INSTANCE"}


def guardar_conexion_servidor(valores):
    """Escribe .env de forma atómica; nunca retorna ni registra los secretos."""
    if not _CLAVES.issubset(valores) or set(valores) - (_CLAVES | _OPCIONALES):
        raise ValueError("Configuración de conexión inválida.")
    path = PROJECT_ROOT / ".env"
    with _config_lock:
        fd, tmp = tempfile.mkstemp(prefix=".servicenow-", suffix=".env", dir=PROJECT_ROOT)
        os.close(fd)
        try:
            if path.exists():
                shutil.copyfile(path, tmp)
            for name, value in valores.items():
                set_key(tmp, name, str(value), quote_mode="always", encoding="utf-8")
            os.replace(tmp, path)
            os.environ.update({k: str(v) for k, v in valores.items()})
        finally:
            if os.path.exists(tmp):
                os.unlink(tmp)


def conexion_configurada():
    if str(config_value("SERVICENOW_AUTH_TYPE", "basic")).strip().lower() == "client_credentials":
        return all(config_value(k) for k in ("SERVICENOW_INSTANCE", "SERVICENOW_CLIENT_ID", "SERVICENOW_CLIENT_SECRET"))
    return all(config_value(k) for k in ("SERVICENOW_INSTANCE", "SERVICENOW_USER", "SERVICENOW_PASSWORD"))


def sincronizacion_habilitada():
    return str(config_value("SERVICENOW_SYNC_ENABLED", "false")).lower() == "true"
