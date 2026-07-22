"""Detección de clientes clave a partir de nombres y alias conocidos."""

import re

import pandas as pd

from app_logic import normalizar_texto
from config.clientes_clave import CLIENTES_CLAVE_ALIASES


def aliases_clientes_ordenados():
    """Devuelve los alias normalizados, priorizando los textos más específicos."""
    aliases = []
    for cliente, opciones in CLIENTES_CLAVE_ALIASES.items():
        for alias in opciones:
            alias_normalizado = normalizar_texto(alias)
            if alias_normalizado:
                aliases.append((cliente, alias_normalizado))
    return sorted(aliases, key=lambda item: len(item[1]), reverse=True)


CLIENTES_CLAVE_ALIAS_ORDENADOS = aliases_clientes_ordenados()


def texto_contiene_alias(texto_normalizado, alias_normalizado):
    """Comprueba un alias completo para evitar coincidencias dentro de palabras."""
    patron = rf"(?<!\w){re.escape(alias_normalizado)}(?!\w)"
    return re.search(patron, texto_normalizado) is not None


def detectar_cliente_clave(texto):
    """Retorna el nombre oficial del cliente detectado o una cadena vacía."""
    texto_normalizado = normalizar_texto(texto)
    if not texto_normalizado:
        return ""
    for cliente, alias in CLIENTES_CLAVE_ALIAS_ORDENADOS:
        if texto_contiene_alias(texto_normalizado, alias):
            return cliente
    return ""


def _valor_limpio(valor):
    if valor is None or pd.isna(valor):
        return ""
    return str(valor).strip()


def detectar_cliente_en_fila(row, campos):
    """Busca un cliente en una fila y devuelve también el campo de origen."""
    for campo in campos:
        cliente = detectar_cliente_clave(_valor_limpio(row.get(campo)))
        if cliente:
            return cliente, campo
    return "", ""

