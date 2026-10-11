"""Seguimiento local de proveedores y ANS con evidencia y reloj contractual.

Las fechas escritas en las notas se interpretan en Colombia. La fecha del
encabezado de un correo no sustituye una fecha de escalamiento o de respuesta.
"""

import json
import re
from datetime import date, datetime, time, timedelta
from functools import lru_cache
from pathlib import Path

import pandas as pd

from services.casos_historial import eventos_historial, limpiar_mensaje, normalizar, texto
from services.casos_sla import fecha_plataforma


CONFIRMADO = "Escalamiento confirmado"
MENCION = "Mención del proveedor"
SIN_PROVEEDOR = "Proveedor sin identificar"
SIN_CLIENTE = "Cliente sin identificar"
ETAPAS = {"atencion": "Atención", "diagnostico": "Diagnóstico",
          "mitigacion": "Mitigación", "solucion": "Solución"}
CAMPOS_TEXTO = (
    "descripcion", "breve_descripcion", "descripcion_detallada", "sn__description",
    "producto", "servicio_negocio", "tipo_falla", "causa", "causa_raiz_original",
    "notas_trabajo", "lista_notas_trabajo", "observaciones_trabajo",
    "observaciones_adicionales", "notas_resolucion", "nombre_proveedor",
    "codigo_resolucion", "producto_adicional", "caso_referencia", "actualizaciones", "impacto",
)
CAMPOS_INTERNOS = {"notas_trabajo", "lista_notas_trabajo", "observaciones_trabajo",
                   "notas_resolucion", "causa", "causa_raiz_original", "nombre_proveedor"}
FECHA = r"(?:\d{4}-\d{2}-\d{2}|\d{1,2}/\d{1,2}/\d{4})[ T]+\d{1,2}:\d{2}(?::\d{2})?(?:\s*[AP]M|Z|[+-]\d{2}:\d{2})?"
HITOS = (
    ("escalamiento", r"(?:fecha(?:/hora)?\s+(?:de\s+)?)?(?:escalamiento(?:\s+al?\s+proveedor)?|escalad[oa](?:\s+al?\s+proveedor)?)"),
    ("atencion", r"(?:fecha(?:/hora)?\s+(?:de\s+)?)?(?:primera\s+respuesta(?:\s+del?\s+proveedor)?|respuesta\s+(?:inicial|del?\s+proveedor)|atencion(?:\s+del?\s+proveedor)?)"),
    ("diagnostico", r"(?:fecha(?:/hora)?\s+(?:de\s+)?)?diagnostico(?:\s+del?\s+proveedor)?"),
    ("mitigacion", r"(?:fecha(?:/hora)?\s+(?:de\s+)?)?mitigacion(?:\s+del?\s+proveedor)?"),
    ("solucion", r"(?:fecha(?:/hora)?\s+(?:de\s+)?)?(?:solucion(?:\s+del?\s+proveedor)?|resuelto\s+por\s+proveedor)"),
)
ESCALADO = re.compile(r"\b(?:se escalo|escalamos|caso escalado|incidente escalado|"
                      r"escalad[oa] (?:al|a|con)|escalamiento (?:al|a|con))\b")
NEGACION = re.compile(r"\b(?:no|sin|pendiente(?: de)?|por|debe|deberia|solicita|solicitamos|favor)\s+"
                      r"(?:\w+\s+){0,2}(?:escal\w*)\b|\bno se escalo\b")
COLUMNAS = ["Tipo", "Número", "Proveedor", "Cliente", "Relación", "Estado", "Prioridad",
            "Descripción", "Creado", "Cerrado", "Escalamiento", "Primera respuesta",
            "Diagnóstico", "Mitigación", "Solución", "Ciclo", "Fuente", "Evidencia",
            "Evidencia de fechas", "Datos de fechas", "Referencia externa"]


@lru_cache(maxsize=1)
def catalogo_ans():
    return json.loads((Path(__file__).resolve().parents[1] / "config/proveedores_ans.json").read_text(encoding="utf-8"))


@lru_cache(maxsize=1)
def patrones_proveedores():
    return {nombre: re.compile(r"(?<!\w)(?:" + "|".join(re.escape(normalizar(a)) for a in datos["aliases"]) + r")(?!\w)")
            for nombre, datos in catalogo_ans()["proveedores"].items()}


def nombres_proveedores(contenido):
    normalizado = normalizar(contenido)
    return [nombre for nombre, patron in patrones_proveedores().items() if patron.search(normalizado)]


def prioridad_ans(valor):
    valor = normalizar(valor)
    if re.match(r"^(?:p|prioridad\s*)?[5-9]\b", valor):
        return None
    numero = re.search(r"(?:^|\b(?:p|prioridad\s*))([1-4])(?:\b|\s*-)", valor)
    if numero:
        return int(numero.group(1))
    for p, patron in ((1, r"\bcritic"), (2, r"\b(?:alto|alta|high)\b"), (3, r"\b(?:moderad|medi)"), (4, r"\b(?:baj|low)")):
        if re.search(patron, valor):
            return p
    return None


def verdadero(valor):
    return normalizar(valor) in {"1", "true", "verdadero", "si", "yes"}


def cliente_registro(row, tipo):
    valor = texto(row.get("cuenta" if tipo == "Caso" else "empresa"))
    if normalizar(valor) in {"", "(vacio)", "(empty)", "sin dato", "sin informacion", "null", "none", "nan", "n/a"}:
        return SIN_CLIENTE
    return valor


def proveedores_escalados(linea):
    """El destinatario del escalamiento no es todo proveedor citado en la frase."""
    normalizado = normalizar(linea)
    nombres = sorted((m.start(), m.end(), nombre) for nombre, patron in patrones_proveedores().items()
                     for m in patron.finditer(normalizado))
    confirmados = set()
    for accion in ESCALADO.finditer(normalizado):
        comienzo = max(normalizado.rfind(",", 0, accion.start()), normalizado.rfind(";", 0, accion.start())) + 1
        fin = min([p for p in (normalizado.find(",", accion.end()), normalizado.find(";", accion.end())) if p >= 0] or [len(normalizado)])
        if NEGACION.search(normalizado[comienzo:fin]):
            continue
        siguientes = [n for n in nombres if accion.end() <= n[0] < fin and n[0] - accion.end() <= 100]
        if siguientes:
            confirmado = siguientes[0]
            confirmados.add(confirmado[2])
            for siguiente in siguientes[1:]:
                if re.fullmatch(r"\s*(?:y|e|/|,)\s*", normalizado[confirmado[1]:siguiente[0]]):
                    confirmados.add(siguiente[2])
                    confirmado = siguiente
                else:
                    break
        else:
            anteriores = [n for n in nombres if comienzo <= n[1] <= accion.start() and accion.start() - n[1] <= 80]
            if anteriores:
                confirmados.add(anteriores[-1][2])
    return confirmados


def fragmentos_proveedor(row, tipo):
    fragmentos = []
    for campo in CAMPOS_TEXTO:
        valor = row.get(campo)
        if campo == "observaciones_trabajo" and isinstance(valor, str) and " | " in valor:
            # El mismo formato se usa para el historial de incidentes.
            eventos = eventos_historial(valor)
            if any(e.get("tipo") != "Texto sin autor" for e in eventos):
                for evento in eventos:
                    fragmentos.append((campo, limpiar_mensaje(evento.get("contenido")),
                                       normalizar(evento.get("tipo")) in {"nota de trabajo", "correo saliente"}))
                continue
        contenido = limpiar_mensaje(valor)
        if contenido:
            fragmentos.append((campo, contenido, campo in CAMPOS_INTERNOS or tipo == "Incidente"))
    for evento in eventos_historial(row.get("historial_conversacion", [])):
        contenido = limpiar_mensaje(evento.get("contenido"))
        if contenido:
            fragmentos.append(("Historial · " + texto(evento.get("tipo")), contenido,
                               normalizar(evento.get("tipo")) in {"nota de trabajo", "correo saliente"}))
    return fragmentos


def hitos_notas(fragmentos, proveedores, proveedor_campo):
    """Asocia fechas explícitas por bloque y proveedor; no mezcla dos proveedores."""
    hitos = {p: [] for p in proveedores}
    invalidos = set()
    evidencias = {p: [] for p in proveedores}
    for fuente, contenido, interno in fragmentos:
        if not interno:
            continue
        # El campo de proveedor o una única mención en el bloque da contexto.
        presentes = nombres_proveedores(contenido)
        contexto = presentes[0] if len(presentes) == 1 else proveedor_campo if not presentes else None
        if contexto is None and len(proveedores) == 1 and not presentes:
            contexto = proveedores[0]
        for linea in re.split(r"\n| --- |;", contenido):
            linea = normalizar(linea)
            locales = nombres_proveedores(linea)
            if locales:
                contexto = locales[0] if len(locales) == 1 else None
            if contexto not in hitos or NEGACION.search(linea):
                continue
            for etapa, patron in HITOS:
                etiqueta = re.search(r"\b" + patron + r"\s*[:=]\s*(.*)", linea)
                if not etiqueta:
                    continue
                encontrado = re.match(FECHA, etiqueta.group(1), re.I)
                fecha = fecha_plataforma(encontrado.group(0)) if encontrado else pd.NaT
                if pd.isna(fecha):
                    if etiqueta.group(1).strip() not in {"", "pendiente", "sin fecha", "n/a"}:
                        invalidos.add(contexto)
                    continue
                hitos[contexto].append((etapa, fecha))
                evidencia = f"{fuente}: {linea.strip()}"
                if evidencia not in evidencias[contexto]:
                    evidencias[contexto].append(evidencia)
    return hitos, invalidos, evidencias


def registros_proveedores(casos, incidentes):
    filas = []
    for tipo, df in (("Caso", casos), ("Incidente", incidentes)):
        for row in df.to_dict("records"):
            numero = texto(row.get("numero"))
            if not numero:
                continue
            fragmentos = fragmentos_proveedor(row, tipo)
            evidencia, confirmados = {}, set()
            campo = texto(row.get("nombre_proveedor"))
            nombres_campo = nombres_proveedores(campo)
            for fuente, contenido, interno in fragmentos:
                for linea in re.split(r"\n| --- |(?<=[.!?])\s+", contenido):
                    destinos_linea = proveedores_escalados(linea) if interno else set()
                    for nombre in nombres_proveedores(linea):
                        evidencia.setdefault(nombre, (fuente, linea.strip()))
                        if nombre in destinos_linea:
                            confirmados.add(nombre)
                            evidencia[nombre] = (fuente, linea.strip())
            # Conserva proveedores adicionales escritos en el campo, aunque no tengan ANS.
            if campo and not nombres_campo and normalizar(campo) not in {"n/a", "no aplica", "ninguno", "sin dato"}:
                evidencia[campo] = ("nombre_proveedor", campo)
                nombres_campo = [campo]
            escalado = verdadero(row.get("escalado_proveedor")) or verdadero(row.get("sn__u_boolean_3"))
            if escalado:
                destinos = nombres_campo or (list(evidencia) if len(evidencia) == 1 else [SIN_PROVEEDOR])
                for nombre in destinos:
                    if nombre not in confirmados:
                        evidencia[nombre] = ("escalado_proveedor / nombre_proveedor", f"Escalado a proveedor: sí · {nombre}")
                    confirmados.add(nombre)
            proveedores = list(evidencia)
            hitos, invalidos, ev_fechas = hitos_notas(fragmentos, proveedores, nombres_campo[0] if len(nombres_campo) == 1 else None)
            for proveedor in proveedores:
                eventos = sorted(set(hitos[proveedor]), key=lambda h: h[1])
                inicios = sorted({f for etapa, f in eventos if etapa == "escalamiento"})
                if inicios:
                    confirmados.add(proveedor)
                    evidencia[proveedor] = ("Notas de trabajo · fechas", next(
                        (v for v in ev_fechas[proveedor] if "escal" in normalizar(v)), evidencia[proveedor][1]))
                for ciclo, inicio in enumerate(inicios or [pd.NaT], 1):
                    siguiente = inicios[ciclo] if ciclo < len(inicios) else None
                    fechas = {etapa: min((f for e, f in eventos if e == etapa and
                                         (pd.isna(inicio) or f >= inicio) and (siguiente is None or f < siguiente)), default=pd.NaT)
                              for etapa in ETAPAS}
                    anterior = any(f < inicio for e, f in eventos if e != "escalamiento") if len(inicios) == 1 else False
                    fuente, detalle = evidencia[proveedor]
                    filas.append({"Tipo": tipo, "Número": numero, "Proveedor": proveedor,
                                  "Cliente": cliente_registro(row, tipo),
                                  "Relación": CONFIRMADO if proveedor in confirmados else MENCION,
                                  "Estado": texto(row.get("estado")), "Prioridad": prioridad_ans(row.get("prioridad")),
                                  "Descripción": texto(row.get("breve_descripcion")) or texto(row.get("descripcion")),
                                  "Creado": fecha_plataforma(row.get("creado")), "Cerrado": fecha_plataforma(row.get("cerrado")),
                                  "Escalamiento": inicio, "Primera respuesta": fechas["atencion"],
                                  "Diagnóstico": fechas["diagnostico"], "Mitigación": fechas["mitigacion"], "Solución": fechas["solucion"],
                                  "Ciclo": ciclo, "Fuente": fuente, "Evidencia": detalle,
                                  "Evidencia de fechas": "\n".join(ev_fechas[proveedor]),
                                  "Datos de fechas": "Fechas inconsistentes" if proveedor in invalidos or anterior else "",
                                  "Referencia externa": texto(row.get("numero_caso_externo"))})
    return pd.DataFrame(filas, columns=COLUMNAS).drop_duplicates(["Tipo", "Número", "Proveedor", "Ciclo"], keep="last")


@lru_cache(maxsize=32)
def festivos_colombia(anio):
    """Ley 51/1983; Pascua gregoriana y traslados al lunes siguiente."""
    a, b, c = anio % 19, anio // 100, anio % 100
    d, e = b // 4, b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i, k = c // 4, c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    pascua = date(anio, (h + l - 7 * m + 114) // 31, (h + l - 7 * m + 114) % 31 + 1)
    fijos = {date(anio, mes, dia) for mes, dia in ((1, 1), (5, 1), (7, 20), (8, 7), (12, 8), (12, 25))}
    trasladados = [date(anio, mes, dia) for mes, dia in ((1, 6), (3, 19), (6, 29), (8, 15), (10, 12), (11, 1), (11, 11))]
    trasladados += [pascua + timedelta(days=dias) for dias in (39, 60, 68)]
    return frozenset(fijos | {pascua - timedelta(days=3), pascua - timedelta(days=2)} |
                     {fecha + timedelta(days=(-fecha.weekday()) % 7) for fecha in trasladados})


def horario_dia(proveedor, prioridad, dia):
    festivo = dia in festivos_colombia(dia.year)
    if proveedor == "Software Colombia" or (proveedor == "Grafólogos Forenses" and prioridad in {1, 2}) or (proveedor == "AutenTIC" and prioridad == 1):
        return 0, 24
    if proveedor == "Grafólogos Forenses":
        return 8, 22
    if proveedor == "AutenTIC":
        return (8, 19) if dia.weekday() == 6 or festivo else (6, 22)
    if proveedor == "RPost / eSoluciona":
        # El Excel no define una excepción por festivos para este proveedor.
        # Se aplica su horario del día de semana, sin inventar una exclusión.
        if dia.weekday() == 6:
            return None
        return (8, 13) if dia.weekday() == 5 else (7, 18)
    return None


@lru_cache(maxsize=8192)
def minutos_servicio(inicio, fin, proveedor, prioridad):
    if fin < inicio:
        raise ValueError("La fecha final antecede al escalamiento")
    if proveedor == "Software Colombia" or (proveedor == "Grafólogos Forenses" and prioridad in {1, 2}) or (proveedor == "AutenTIC" and prioridad == 1):
        return (fin - inicio).total_seconds() / 60
    total, dia = 0.0, inicio.date()
    while dia <= fin.date():
        ventana = horario_dia(proveedor, prioridad, dia)
        if ventana:
            medianoche = pd.Timestamp(datetime.combine(dia, time()), tz="America/Bogota")
            a, b = medianoche + pd.Timedelta(hours=ventana[0]), medianoche + pd.Timedelta(hours=ventana[1])
            total += max(0, (min(fin, b) - max(inicio, a)).total_seconds() / 60)
        dia += timedelta(days=1)
    return total


def plazo_dias_habiles(inicio, dias):
    # Un día hábil es lunes–viernes, sin festivos colombianos; conserva la hora.
    fin = inicio
    while dias:
        fin += pd.Timedelta(days=1)
        if fin.weekday() < 5 and fin.date() not in festivos_colombia(fin.year):
            dias -= 1
    return fin


def filtrar_periodo_proveedores(df, inicio, fin):
    """Mes de escalamiento y pendientes anteriores; las menciones usan creación.

    Un escalado sin fecha no se excluye como si no existiera: se presenta como
    pendiente de fecha en cada corte posterior a su creación hasta su cierre.
    """
    if df.empty:
        return df.copy()
    conservar, etiquetas = [], []
    for row in df.to_dict("records"):
        escalamiento, solucion, creado, cerrado = (row[c] for c in ("Escalamiento", "Solución", "Creado", "Cerrado"))
        referencia = escalamiento if pd.notna(escalamiento) else creado
        cierre = solucion if pd.notna(solucion) else cerrado
        en_mes = pd.notna(referencia) and inicio <= referencia < fin
        previo = row["Relación"] == CONFIRMADO and (pd.isna(referencia) or referencia < inicio) and (pd.isna(cierre) or cierre >= inicio)
        conservar.append(en_mes or previo)
        etiquetas.append("Escalado en el período" if en_mes and pd.notna(escalamiento) else
                         "Pendiente de fecha de escalamiento" if row["Relación"] == CONFIRMADO and pd.isna(escalamiento) else
                         "Pendiente de período anterior" if previo else "Mención en caso del período")
    resultado = df.loc[conservar].copy()
    resultado["Período del registro"] = [v for v, conservar_fila in zip(etiquetas, conservar) if conservar_fila]
    return resultado


def evaluar_ans_proveedores(df, corte):
    resultado = df.copy()
    catalogo = catalogo_ans()["proveedores"]
    for etapa, etiqueta in ETAPAS.items():
        estados, minutos, objetivos = [], [], []
        for row in df.to_dict("records"):
            datos = catalogo.get(row["Proveedor"], {}).get("prioridades", {}).get(str(int(row["Prioridad"])) if pd.notna(row["Prioridad"]) else "", {})
            objetivo = datos.get("objetivos", {}).get(etapa)
            objetivos.append(datos.get("texto", {}).get(etapa, "Sin ANS definido"))
            inicio = row["Escalamiento"]
            fecha = row["Primera respuesta" if etapa == "atencion" else etiqueta]
            estado, transcurrido = "Sin datos", None
            if row["Relación"] != CONFIRMADO:
                estado = "Mención; no evaluable"
            elif row["Proveedor"] not in catalogo:
                estado = "Sin ANS definido"
            elif pd.isna(row["Prioridad"]):
                estado = "Sin prioridad ANS"
            elif objetivo is None:
                estado = "No aplica"
            elif row["Datos de fechas"] or (pd.notna(fecha) and pd.notna(inicio) and fecha < inicio):
                estado = "Fechas inconsistentes"
            elif pd.notna(inicio) and inicio <= corte:
                terminada = pd.notna(fecha) and fecha <= corte
                finalizado = (pd.notna(row["Solución"]) and row["Solución"] <= corte
                              or pd.notna(row["Cerrado"]) and row["Cerrado"] <= corte
                              or re.search(r"\b(?:cerrado|closed|resuelto|resolved|cancelado)\b", normalizar(row["Estado"])) and pd.isna(row["Cerrado"]))
                if not terminada and finalizado:
                    estado = "Sin datos"
                else:
                    fin = fecha if terminada else corte
                    tiempo_real = minutos_servicio(inicio, fin, row["Proveedor"], int(row["Prioridad"]))
                    transcurrido = round(tiempo_real, 2)
                    dentro = fin <= plazo_dias_habiles(inicio, objetivo["dias_habiles"]) if isinstance(objetivo, dict) else tiempo_real <= objetivo
                    estado = ("Cumple" if dentro else "No cumple") if terminada else ("En plazo" if dentro else "Vencido")
            estados.append(estado)
            minutos.append(transcurrido)
        resultado[f"ANS {etiqueta}"] = estados
        resultado[f"Minutos {etiqueta}"] = minutos
        resultado[f"Objetivo {etiqueta}"] = objetivos
    return resultado


def resumen_ans_proveedores(df):
    filas = []
    for proveedor, grupo in df.groupby("Proveedor", sort=True):
        unicos = grupo.drop_duplicates(["Tipo", "Número"])
        fila = {"Proveedor": proveedor, "Casos": int((unicos["Tipo"] == "Caso").sum()),
                "Incidentes": int((unicos["Tipo"] == "Incidente").sum()),
                "Clientes": unicos.loc[unicos["Cliente"] != SIN_CLIENTE, "Cliente"].nunique(),
                "Escalados": len(grupo[grupo["Relación"] == CONFIRMADO].drop_duplicates(["Tipo", "Número"])),
                "Menciones": len(grupo[grupo["Relación"] == MENCION].drop_duplicates(["Tipo", "Número"])),
                "Ciclos": int((grupo["Relación"] == CONFIRMADO).sum())}
        for etiqueta in ETAPAS.values():
            estados = grupo[f"ANS {etiqueta}"]
            cumple, incumple = int((estados == "Cumple").sum()), int((estados == "No cumple").sum())
            total = cumple + incumple
            fila.update({f"{etiqueta} evaluados": total, f"{etiqueta} cumplen": cumple,
                         f"{etiqueta} no cumplen": incumple,
                         f"{etiqueta} % cumplimiento": round(100 * cumple / total, 2) if total else None,
                         f"{etiqueta} vencidos": int((estados == "Vencido").sum()),
                         f"{etiqueta} en plazo": int((estados == "En plazo").sum()),
                         f"{etiqueta} sin datos": int(estados.isin(["Sin datos", "Fechas inconsistentes", "Sin prioridad ANS", "Sin ANS definido"]).sum())})
        filas.append(fila)
    return pd.DataFrame(filas)
