"""Detección de clientes clave a partir de nombres y alias conocidos."""

import re

import pandas as pd

from app_logic import normalizar_texto
from config.clientes_clave import CLIENTES_CLAVE_ALIASES, GRUPOS_CLIENTES_CLAVE


GRUPO_POR_CLIENTE = {
    cliente: grupo
    for grupo, clientes in GRUPOS_CLIENTES_CLAVE.items()
    for cliente in clientes
}


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


def detectar_grupo_cliente_clave(texto):
    """Retorna el grupo del cliente encontrado en un texto o una cadena vacía."""
    cliente = detectar_cliente_clave(texto)
    return GRUPO_POR_CLIENTE.get(cliente, "")


def filtrar_por_grupo_cliente_clave(df, columna, grupo):
    """Filtra un DataFrame por grupo sin modificar los registros originales."""
    if df.empty or not grupo:
        return df.copy()
    if columna not in df.columns:
        return df.iloc[0:0].copy()
    grupos_detectados = df[columna].apply(detectar_grupo_cliente_clave)
    return df[grupos_detectados == grupo].copy()


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
