"""Funciones puras de normalización de identidad y contraseñas."""

import hashlib
import hmac
import os


PASSWORD_ALGORITHM = "pbkdf2_sha256"
PASSWORD_ITERATIONS = 260_000


def normalizar_email(email):
    return str(email or "").strip().lower()


def hash_password(password):
    salt = os.urandom(16)
    password_hash = hashlib.pbkdf2_hmac(
        "sha256",
        str(password).encode("utf-8"),
        salt,
        PASSWORD_ITERATIONS,
    )
    return f"{PASSWORD_ALGORITHM}${PASSWORD_ITERATIONS}${salt.hex()}${password_hash.hex()}"


def verificar_password(password, password_hash):
    try:
        algoritmo, iteraciones, salt_hex, hash_hex = str(password_hash).split("$")
        if algoritmo != PASSWORD_ALGORITHM:
            return False
        calculado = hashlib.pbkdf2_hmac(
            "sha256",
            str(password).encode("utf-8"),
            bytes.fromhex(salt_hex),
            int(iteraciones),
        ).hex()
        return hmac.compare_digest(calculado, hash_hex)
    except (TypeError, ValueError):
        return False

