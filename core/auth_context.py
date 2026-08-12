"""Contexto servidor de la vista autenticada en ejecución."""

from contextvars import ContextVar


_actor_email = ContextVar("actor_email", default="")
_view_id = ContextVar("view_id", default="")


def establecer_contexto_autorizacion(actor_email, view_id):
    return _actor_email.set(str(actor_email or "")), _view_id.set(str(view_id or ""))


def obtener_contexto_autorizacion():
    return _actor_email.get(), _view_id.get()


def restaurar_contexto_autorizacion(tokens):
    actor_token, view_token = tokens
    _actor_email.reset(actor_token)
    _view_id.reset(view_token)
