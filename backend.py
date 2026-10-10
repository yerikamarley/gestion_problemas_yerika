"""Proxy HTTP opcional. Autentica usuarios locales; ServiceNow solo se consume en Python."""

import base64
import binascii

from starlette.applications import Starlette
from starlette.concurrency import run_in_threadpool
from starlette.requests import Request
from starlette.responses import JSONResponse
from starlette.routing import Route

from app_logic import (
    AutorizacionError, autenticar_usuario, consultar_traza_servicenow, sincronizar_servicenow,
)
from services.servicenow import ServiceNowError


def _respuesta(data, status=200):
    return JSONResponse(data, status_code=status, headers={"Cache-Control": "no-store"})


def _actor(request):
    """HTTP Basic de la aplicación; nunca utiliza el usuario API de ServiceNow."""
    header = request.headers.get("authorization", "")
    try:
        scheme, encoded = header.split(" ", 1)
        if scheme.lower() != "basic":
            return None
        email, password = base64.b64decode(encoded, validate=True).decode("utf-8").split(":", 1)
    except (ValueError, UnicodeDecodeError, binascii.Error):
        return None
    usuario = autenticar_usuario(email, password)
    return usuario["email"] if usuario else None


def _operar(request, body=None):
    try:
        actor = _actor(request)
        if not actor:
            response = _respuesta({"error": "Autenticación requerida."}, 401)
            response.headers["WWW-Authenticate"] = 'Basic realm="gestion", charset="UTF-8"'
            return response
        tipo = request.path_params.get("tipo", "casos")
        numero = request.path_params.get("numero")
        if request.method == "GET":
            documento = consultar_traza_servicenow(numero, actor_email=actor, tipo=tipo)
            return _respuesta(documento) if documento else _respuesta({"error": "Registro no sincronizado."}, 404)
        if numero:
            resultado = sincronizar_servicenow(actor_email=actor, tipo=tipo, numero=numero)
            if not resultado["cargados"] and not resultado.get("ocupado"):
                return _respuesta({"error": "Registro no encontrado."}, 404)
        else:
            if not isinstance(body, dict) or set(body) not in ({"mes"}, {"fecha_inicio", "fecha_fin"}):
                raise ValueError("Envía mes (AAAA-MM) o fecha_inicio y fecha_fin (AAAA-MM-DD).")
            resultado = sincronizar_servicenow(actor_email=actor, tipo=tipo, **body)
        if resultado.get("ocupado"):
            return _respuesta({"error": "Ya hay una sincronización de casos en curso."}, 409)
        return _respuesta(resultado)
    except AutorizacionError:
        return _respuesta({"error": "No tienes permisos para realizar esta operación."}, 403)
    except ValueError as error:
        return _respuesta({"error": str(error)}, 400)
    except ServiceNowError as error:
        return _respuesta({"error": str(error)}, 502)
    except Exception:
        # No devolver excepciones de requests/psycopg: pueden incluir secretos.
        return _respuesta({"error": "No se pudo completar la operación."}, 503)


async def endpoint(request: Request):
    body = None
    if request.method == "POST" and "numero" not in request.path_params:
        try:
            body = await request.json()
        except ValueError:
            return _respuesta({"error": "JSON inválido."}, 400)
    return await run_in_threadpool(_operar, request, body)


app = Starlette(debug=False, routes=[
    Route("/api/servicenow/{tipo}/sync", endpoint, methods=["POST"]),
    Route("/api/servicenow/{tipo}/{numero}/sync", endpoint, methods=["POST"]),
    Route("/api/servicenow/{tipo}/{numero}", endpoint, methods=["GET"]),
])
