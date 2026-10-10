"""Prueba directa del token OAuth; no lee la BD ni guarda secretos o tokens."""

import argparse
import getpass
import re
from datetime import datetime, timezone

import requests


ERROR_CODES = {
    "invalid_client", "unauthorized_client", "unsupported_grant_type",
    "invalid_scope", "invalid_grant", "invalid_request", "access_denied",
    "server_error", "temporarily_unavailable",
}


def diagnose(instance, client_id, secret, output=print):
    """Devuelve 0 con token emitido, 1 con rechazo o 2 con entrada/red inválida."""
    if not re.fullmatch(r"[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.service-now\.com", instance):
        output("Instancia inválida: usa solamente el dominio terminado en .service-now.com.")
        return 2
    if not client_id or not secret:
        output("Falta el ID o el secreto de cliente.")
        return 2
    if any("\ufdd0" <= char <= "\ufdef" for char in secret):
        output("El secreto contiene marcadores de cifrado: necesitas el valor original, no el almacenado cifrado.")
        return 2
    if secret != secret.strip() or "\n" in secret or "\r" in secret:
        output("El secreto contiene espacios en los extremos o saltos de línea. Revisa la copia.")
        return 2
    output("Fecha UTC: " + datetime.now(timezone.utc).isoformat(timespec="seconds"))
    try:
        with requests.Session() as session:
            # Evita que .netrc añada autenticación Basic a esta prueba OAuth.
            session.trust_env = False
            response = session.post(
                f"https://{instance}/oauth_token.do",
                data={"grant_type": "client_credentials", "client_id": client_id,
                      "client_secret": secret},
                headers={"Accept": "application/json"},
                timeout=(10, 30), allow_redirects=False,
            )
    except requests.RequestException:
        output("No se pudo completar la conexión HTTPS. No se muestran detalles privados de la petición.")
        return 2
    output(f"HTTP: {response.status_code}")
    transaction = response.headers.get("X-Transaction-ID", "")
    if re.fullmatch(r"[a-fA-F0-9-]{16,64}", transaction):
        output("ID de transacción: " + transaction)
    try:
        payload = response.json()
    except ValueError:
        payload = None
    if not isinstance(payload, dict):
        output("Respuesta sin objeto JSON OAuth. El contenido se omite.")
        return 1
    if response.status_code == 200 and isinstance(payload.get("access_token"), str) and payload["access_token"]:
        output("TOKEN EMITIDO. El token se omite y no se guarda.")
        output("Estas credenciales funcionan desde este equipo. Compara los secretos efectivos y reinicia Streamlit.")
        return 0
    code = payload.get("error")
    output("OAuth: " + (code if isinstance(code, str) and code in ERROR_CODES else "código no reconocido (omitido)"))
    if payload.get("error_description") == "access_denied":
        output("Descripción OAuth: access_denied")
    output("No se emitió un token. Esta prueba no usa Streamlit ni Supabase.")
    return 1


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--instance", default="certicamara.service-now.com")
    parser.add_argument("--client-id", default="")
    args = parser.parse_args()
    try:
        client_id = args.client_id or input("ID de cliente de portalgestion: ").strip()
        secret = getpass.getpass("Secreto de cliente (oculto, no se guarda): ")
    except (EOFError, KeyboardInterrupt):
        print("Prueba cancelada.")
        return 2
    return diagnose(args.instance, client_id, secret)


if __name__ == "__main__":
    raise SystemExit(main())
