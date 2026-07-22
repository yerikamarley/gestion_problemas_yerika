"""Lectura centralizada de configuración y secretos del entorno."""

import os
import tomllib
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent.parent


def streamlit_secrets_path():
    """Localiza el archivo de secretos del proyecto o del usuario."""
    candidates = [PROJECT_ROOT / ".streamlit" / "secrets.toml"]
    try:
        candidates.append(Path.home() / ".streamlit" / "secrets.toml")
    except RuntimeError:
        pass
    return next((str(path) for path in candidates if path.exists()), "")


def local_secret_value(name, default=""):
    """Lee un secreto local sin fallar cuando el archivo no está disponible."""
    secrets_path = streamlit_secrets_path()
    if not secrets_path:
        return default
    try:
        with open(secrets_path, "rb") as file:
            return tomllib.load(file).get(name, default)
    except (OSError, tomllib.TOMLDecodeError):
        return default


def config_value(name, default=""):
    """Resuelve configuración: entorno, TOML local y secretos de Streamlit."""
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
    except (ImportError, KeyError, RuntimeError):
        return default


SUPABASE_URL = config_value("SUPABASE_URL")
SUPABASE_PUBLISHABLE_KEY = config_value("SUPABASE_PUBLISHABLE_KEY")
SUPABASE_DATABASE_URL = config_value("SUPABASE_DATABASE_URL")
ADMIN_EMAIL = config_value("APP_ADMIN_EMAIL")
INITIAL_ADMIN_PASSWORD = config_value("APP_ADMIN_PASSWORD")
