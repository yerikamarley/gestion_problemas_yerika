"""Cruce temporal entre casos y la lista de clientes de la migración PKI."""

import io
import re
import unicodedata

import pandas as pd


GRUPO_PKI_CON_COMPONENTES = "Migración PKI – Con componentes"
GRUPO_PKI_SIN_COMPONENTES = "Migración PKI – Sin componentes"
GRUPO_PKI = "Migración PKI"
GRUPO_PKI_TODOS = GRUPO_PKI
GRUPOS_MIGRACION_PKI = (GRUPO_PKI_CON_COMPONENTES, GRUPO_PKI_SIN_COMPONENTES)


def normalizar(valor):
    if valor is None or pd.isna(valor):
        return ""
    texto = unicodedata.normalize("NFKD", str(valor)).encode("ascii", "ignore").decode().casefold()
    return re.sub(r"\s+", " ", texto).strip()


def _columna(df, opciones):
    columnas = {normalizar(col): col for col in df.columns}
    return next((columnas[normalizar(opcion)] for opcion in opciones if normalizar(opcion) in columnas), None)


def _valores(df, opciones):
    columna = _columna(df, opciones)
    if columna is None:
        return set()
    return {normalizar(valor) for valor in df[columna] if normalizar(valor)}


def leer_lista_migracion(archivo):
    """Lee las dos hojas esperadas sin depender del orden de sus columnas."""
    contenido = archivo.getvalue() if hasattr(archivo, "getvalue") else archivo
    libro = pd.ExcelFile(io.BytesIO(contenido))
    hojas = {normalizar(nombre): nombre for nombre in libro.sheet_names}
    base_nombre = hojas.get(normalizar("mesa de ayuda bdf"))
    componentes_nombre = hojas.get(normalizar("ClientesconComponentes"))
    if base_nombre is None or componentes_nombre is None:
        raise ValueError("El Excel debe contener las hojas 'mesa de ayuda bdf' y 'ClientesconComponentes'.")
    base = pd.read_excel(libro, sheet_name=base_nombre)
    componentes = pd.read_excel(libro, sheet_name=componentes_nombre)
    lista = {
        "correos": _valores(base, ["CORREO", "correo", "email"]),
        "empresas": _valores(base, ["razon social", "empresa", "cuenta"]),
        "nombres": set(),
        "componentes_correos": _valores(componentes, ["correo", "CORREO", "email"]),
        "componentes_empresas": _valores(componentes, ["empresa", "razon social", "cuenta"]),
        "componentes_nombres": _valores(componentes, ["cliente", "nombre"]),
        "filas_base": len(base),
        "filas_componentes": len(componentes),
    }
    lista["correos"] |= lista["componentes_correos"]
    lista["empresas"] |= lista["componentes_empresas"]
    return lista


def _tokens(valor):
    return {token for token in normalizar(valor).split() if len(token) >= 4}


def _coincide(valor, exactos):
    texto = normalizar(valor)
    if not texto:
        return False
    if texto in exactos:
        return True
    tokens = _tokens(texto)
    return any(len(tokens & _tokens(candidato)) >= 2 for candidato in exactos if "@" not in candidato)


def clasificar_fila_migracion(row, lista, campos=("cuenta", "contacto", "creado_por", "empresa")):
    valores = [row.get(campo, "") for campo in campos]
    componente = any(_coincide(valor, lista["componentes_correos"]) or
                     _coincide(valor, lista["componentes_empresas"]) or
                     _coincide(valor, lista["componentes_nombres"]) for valor in valores)
    migracion = componente or any(_coincide(valor, lista["correos"]) or _coincide(valor, lista["empresas"]) for valor in valores)
    if componente:
        return GRUPO_PKI_CON_COMPONENTES
    if migracion:
        return GRUPO_PKI_SIN_COMPONENTES
    return ""


def agregar_grupo_migracion(df, lista):
    trabajo = df.copy()
    trabajo["grupo_migracion_pki"] = trabajo.apply(lambda row: clasificar_fila_migracion(row, lista), axis=1)
    return trabajo


def filtrar_grupo_migracion(df, lista, grupo):
    if not lista or not grupo:
        return df.copy()
    clasificados = agregar_grupo_migracion(df, lista)
    if grupo == GRUPO_PKI_TODOS:
        clasificados = clasificados[clasificados["grupo_migracion_pki"].isin(GRUPOS_MIGRACION_PKI)]
    else:
        clasificados = clasificados.query("grupo_migracion_pki == @grupo")
    return clasificados.drop(columns=["grupo_migracion_pki"])
