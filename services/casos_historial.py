"""Análisis local de motivos, causas y atención con evidencia de la conversación."""

import re
import unicodedata

import pandas as pd

from config.equipo_soporte import EQUIPO_SOPORTE_CASOS


COL_TOKEN = "Tipo de token detectado"
COL_MOTIVO = "Motivo observado"
COL_CAUSA = "Causa identificada"
COL_ESTADO_CAUSA = "Estado de la causa"
COL_ATENCION = "Atención registrada"
COL_DESTINO = "Destino de redirección"
COL_CLASIFICACION = "Clasificación principal"
COL_TIPOLOGIA = "Tipologia caso"
COL_EVIDENCIA_TOKEN = "Evidencia del token"
COL_EVIDENCIA_MOTIVO = "Evidencia del motivo"
COL_EVIDENCIA_CAUSA = "Evidencia de la causa"
COL_EVIDENCIA_ATENCION = "Evidencia de atención o redirección"
COL_FISICO = "_token_fisico_historial"
COL_VIRTUAL = "_token_virtual_historial"
REDIRIGIDO_SIN_ATENCION = "Redirigido sin atención técnica registrada"
ENVIADO_OTROS_CANALES = "Enviado a otros canales"
SIN_ATENCION = "Sin evidencia suficiente de atención"
SIN_CAUSA = "Sin causa documentada"
VERSION_ANALISIS = "historial-limpieza-v4"

HEADER_EVENTO = re.compile(
    r"(?:^| --- )(?P<fecha>\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}:\d{2})\s*\|\s*"
    r"(?P<tipo>Nota de Trabajo|Comentario|Correo[^|]*|Nota[^|]*)\s*\|\s*"
    r"(?P<usuario>[^|]*)\|\s*(?P<asunto>[^|]*)\|\s*", re.I)
PIE = re.compile(r"conoce nuestros canales oficiales|aviso de confidencialidad|"
                 r"darse de baja|preferencias de notificaci[oó]n|"
                 r"(?:^| --- |\n)(?:de:|from:|el .{0,100}escribi[oó]:|on .{0,100}wrote:)", re.I)
AUTOMATICO = re.compile(r"hemos recibido tu solicitud|tu n[uú]mero de caso es|"
                        r"ya estamos trabajando en ella|este es un mensaje autom[aá]tico|"
                        r"no responda a este correo|notificaci[oó]n autom[aá]tica", re.I)
CATALOGO_DERIVACION = re.compile(
    r"\bpara\s+(?:los\s+)?procesos\s+de\b.{0,900}?"
    r"(?=\bpor\s+favor\b|\b(?:separe|agende|reserve)\s+(?:su|tu|una)\s+cita\b)", re.I | re.S)
POLITICA_CANAL = re.compile(
    r"\beste\s+canal\s+es\s+exclusivo\s+para\b[^.\n]{0,500}(?:\.|$)", re.I)
FISICO = re.compile(r"\b(?:token|toquen|tocken)\s+fisico\b")
VIRTUAL = re.compile(r"\b(?:certi\s?token|(?:token|toquen|tocken)\s+(?:virtual|movil|digital)|app\s+token)\b")
REDIRECCION = re.compile(r"\b(?:se (?:redirig\w*|redireccion\w*|remit\w*|deriv\w*|traslad\w*|escal\w*|envi\w*)|"
                         r"caso (?:redirig\w*|redireccion\w*|remit\w*|deriv\w*|traslad\w*|escal\w*)|"
                         r"(?:redirig\w*|redireccion\w*|remitim\w*|derivam\w*|trasladam\w*|escalam\w*|enviam\w*)|"
                         r"(?:debe|debes|favor|por favor)\s+(?:contactar|comunicarse|comunicarte|agendar|radicar)|"
                         r"(?:contacte|contacta|comuniquese|comunicate|agende|agenda tu cita|radique)|"
                         r"(?:separe|reserve|programe)\s+(?:su|tu|una)\s+cita|"
                         r"envio (?:del? |un )?(?:enlace|link)|"
                         r"(?:corresponde|competencia)\s+(?:a|del|de la)|canal correcto|fuera del alcance)\b")
DESTINOS = (("PQRS", re.compile(r"\b(?:pqrs|pqrf|pqr|quejas y reclamos)\b")),
            ("Canal de ventas", re.compile(r"\b(?:canal de ventas|ventas)\b")),
            ("Agenda", re.compile(r"\b(?:agenda|agendar|agendamiento|cita|zcal|agendate)\b")),
            ("Comercial", re.compile(r"\b(?:comercial|ejecutivo de cuenta|asesor comercial)\b")))
ACCION_TECNICA = re.compile(
    r"\b(?:se (?:instalo|reinstalo|configuro|actualizo|desbloqueo|restablecio|verifico|valido|reviso|"
    r"realizo|corrigio|soluciono|probo|oriento|explico|brindo|identifico|determino)|"
    r"(?:instalamos|reinstalamos|configuramos|actualizamos|desbloqueamos|restablecimos|verificamos|"
    r"validamos|revisamos|realizamos|corregimos|solucionamos|probamos|orientamos|explicamos)|"
    r"sesion remota|conexion remota|pruebas realizadas|se comparte (?:el |la |un |una )?(?:manual|guia|instructivo)|"
    r"(?:adjuntamos|adjunto|compartimos|enviamos) (?:el |la |un |una )?(?:manual|guia|instructivo)|"
    r"(?:sigue|realice|realiza) (?:estos|los siguientes) pasos)\b")
CONTEXTO_TECNICO = re.compile(r"\b(?:token|toquen|tocken|certificado|driver|controlador|instala\w*|configura\w*|"
                             r"prueba\w*|firma\w*|equipo|error|falla\w*|conexion remota|sesion remota|java|autofirma|"
                             r"adobe|descarg\w*|activa\w*|software|pin|acceso|contrasena|manual|guia|instructivo|pasos|"
                             r"sistema|servicio|plataforma|portal|consumos?)\b")
GESTION_ADMINISTRATIVA = re.compile(r"^se (?:reviso|valido|verifico|realizo)\s+(?:el |la |un |una )?"
                                  r"(?:solicitud|caso|ticket|compra|cotizacion|redireccion|remision|envio|agenda)\b")
CAUSAL = re.compile(r"\b(?:causa raiz(?: confirmada)?|causa identificada|se identifico que|"
                    r"se determino que|se debe a|debido a|causad[oa] por)\s*[:=-]?\s*(?P<detalle>[^.;\n]{5,})")
GENERICA = re.compile(r"^(?:token(?: fisico| virtual)?|e[ -]?pass|firma(?: digital)?|certificado|"
                     r"soporte(?: uso| falla)?|solicitud(?: operativa)?|en analisis|pendiente(?: de revision)?|"
                     r"sin (?:dato|informacion|causa).*|no aplica|n/?a|ninguna)$")
TEMA_TOKEN = re.compile(r"^(?:(?:token fisico|token virtual|e[ -]?pass|certitoken|safenet)\s*(?:[/+-]\s*)?)+$")
MOTIVOS = (
    ("Caída o indisponibilidad", re.compile(
        r"\b(?:sistema|servicio|plataforma|portal|aplicacion)\b.{0,60}"
        r"(?:caid[oa]|se cayo|se callo|no (?:esta|estan) disponible|indisponib\w*|"
        r"fuera de servicio|no funciona|no responde|sin servicio)|"
        r"(?:se cayo|se callo|caid[oa]|indisponibilidad|fuera de servicio).{0,60}"
        r"\b(?:sistema|servicio|plataforma|portal|aplicacion)\b")),
    ("Consulta o visualización de consumos", re.compile(r"\bconsumos?\b")),
    ("Falla del servicio", re.compile(
        r"\b(?:error\w*|falla\w*|lentitud|lento|interrupcion)\b.{0,60}\b(?:sistema|servicio|plataforma|portal|aplicacion)\b|"
        r"\b(?:sistema|servicio|plataforma|portal|aplicacion)\b.{0,60}\b(?:error\w*|falla\w*|lentitud|lento|interrupcion)\b")),
    ("Seguridad o suplantación", re.compile(r"\b(?:phishing|suplantacion|fraude)\b")),
    ("Bloqueo o PIN", re.compile(r"\b(?:pin|puk|bloquead\w*|bloqueo|contrasena incorrecta)\b")),
    ("Token no reconocido", re.compile(r"(?:no (?:reconoce|detecta|lee)|no reconocido|no detectado|no aparece).{0,50}(?:token|dispositivo|usb)|(?:token|dispositivo).{0,50}(?:no (?:reconoce|detecta|lee)|no reconocido|no detectado)")),
    ("Falla de descarga o activación", re.compile(r"(?:error|falla\w*|problema\w*|dificult\w*|no (?:pued\w*|permite|logr\w*)|imposible).{0,65}(?:descarg\w*|activ\w*)|(?:descarg\w*|activ\w*).{0,65}(?:error|falla\w*|no funciona)")),
    ("Falla al firmar", re.compile(r"(?:error|falla\w*|problema\w*|dificult\w*|no (?:pued\w*|permite)).{0,55}(?:firmar|firma)|(?:firma|firmar).{0,55}(?:error|falla|no funciona)")),
    ("Instalación o configuración", re.compile(r"\b(?:instala\w*|reinstala\w*|configura\w*|driver\w*|controlador\w*)\b")),
    ("Descarga o activación", re.compile(r"\b(?:descarg\w*|activa\w*)\b")),
    ("Uso del token", re.compile(r"(?:uso|usar|utilizar|manejo|ayuda|asesoria).{0,45}(?:token|toquen|tocken)|(?:token|toquen|tocken).{0,45}(?:usar|uso|ayuda|dificultad)")),
)
CLASIFICACIONES_LEGADAS = (
    ("Duplicado", re.compile(r"\b(?:duplicad\w*|caso repetido|ticket repetido)\b")),
    ("Captores", re.compile(r"\bcaptores?\b")),
    ("Firma digital", re.compile(r"\b(?:firma|firmar|certifirma)\b")),
    ("Solicitud operativa", re.compile(r"\b(?:solicitud|pago|pagar|biometria|renovacion|cotizacion|factura|compra)\b")),
)
ETIQUETAS_GESTION = {"agenda", "envio agenda / manual de uso", "enviado a otros canales",
                    "9 - redireccionamiento agenda", "soporte uso", "soporte falla",
                    "instalacion o configuracion", "instalacion", "descarga o activacion",
                    "caida o indisponibilidad", "consulta o visualizacion de consumos", "duplicado", "caso duplicado", "captores"}
CLASIFICACION_POR_MOTIVO = {
    "Caída o indisponibilidad": "Caída o indisponibilidad",
    "Consulta o visualización de consumos": "Consulta o visualización de consumos",
    "Falla del servicio": "Falla del servicio",
    "Seguridad o suplantación": "Seguridad o suplantación",
    "Bloqueo o PIN": "Uso o funcionamiento del token",
    "Token no reconocido": "Uso o funcionamiento del token",
    "Uso del token": "Uso o funcionamiento del token",
    "Falla de descarga o activación": "Descarga o activación",
    "Descarga o activación": "Descarga o activación",
    "Falla al firmar": "Firma digital",
    "Instalación o configuración": "Instalación o configuración",
}
CLASIFICACIONES_PRINCIPALES = (ENVIADO_OTROS_CANALES, *dict.fromkeys(CLASIFICACION_POR_MOTIVO.values()),
                              "Duplicado", "Captores", "Solicitud operativa", "Por revisar")
CAUSAS = (
    ("Controlador ausente, incompatible o corrupto", re.compile(r"(?:driver|controlador).{0,65}(?:ausente|incompatible|corrupt\w*|no instalad\w*|desactualizad\w*)|(?:falta|ausencia|incompatibilidad).{0,45}(?:driver|controlador)")),
    ("Certificado vencido o revocado", re.compile(r"certificado.{0,45}(?:vencid\w*|expirad\w*|revocad\w*)")),
    ("PIN bloqueado por intentos fallidos", re.compile(r"(?:pin|token).{0,65}bloquead\w*.{0,65}intent\w*|intent\w*.{0,65}(?:bloqueo|bloquead\w*)")),
    ("Incompatibilidad de software", re.compile(r"(?:java|autofirma|adobe|navegador|version).{0,65}(?:incompatible|desactualizad\w*)|incompatibilidad.{0,45}(?:java|autofirma|adobe|navegador|version)")),
    ("Permisos o restricciones del equipo", re.compile(r"(?:falta|ausencia|restriccion|bloqueo).{0,50}(?:permiso|administrador|antivirus|firewall)|(?:antivirus|firewall).{0,45}bloque")),
)


def texto(valor):
    if valor is None:
        return ""
    if isinstance(valor, (list, dict)):
        return ""
    return "" if pd.isna(valor) else str(valor).strip()


def normalizar(valor):
    valor = unicodedata.normalize("NFKD", texto(valor)).casefold()
    return " ".join("".join(c for c in valor if not unicodedata.combining(c)).split())


def limpiar_mensaje(valor):
    valor = texto(valor)
    corte = PIE.search(valor)
    return valor[:corte.start()].strip(" -\n") if corte else valor


def eventos_historial(historial):
    if isinstance(historial, list):
        return [dict(evento) for evento in historial if isinstance(evento, dict)]
    historial = texto(historial)
    encabezados = list(HEADER_EVENTO.finditer(historial))
    if not encabezados:
        return [{"tipo": "Texto sin autor", "contenido": historial}] if historial else []
    eventos = []
    for indice, encabezado in enumerate(encabezados):
        fin = encabezados[indice + 1].start() if indice + 1 < len(encabezados) else len(historial)
        eventos.append({**{key: value.strip() for key, value in encabezado.groupdict().items()},
                        "contenido": historial[encabezado.end():fin].strip(" -")})
    return eventos


def fragmentos_caso(row):
    partes = []
    for campo, interno in (("descripcion", False), ("descripcion_detallada", False), ("sn__description", False),
                           ("causa", True), ("producto", False), ("notas_resolucion", True),
                           ("observaciones_adicionales", False), ("observaciones_trabajo", False),
                           ("notas_trabajo", True), ("codigo_resolucion", True)):
        contenido = limpiar_mensaje(row.get(campo))
        if contenido:
            partes.append({"contenido": contenido, "fuente": campo, "tipo": "Campo del caso", "interno": interno})
    historial = row.get("_historial_eventos", row.get("historial_conversacion", []))
    for evento in eventos_historial(historial):
        # El acuse genérico no prueba atención ni un producto afectado.
        contenido = limpiar_mensaje(evento.get("contenido"))
        if AUTOMATICO.search(contenido[:600]):
            continue
        asunto = limpiar_mensaje(evento.get("asunto"))
        if contenido or asunto:
            partes.append({**evento, "contenido": (asunto + ". " + contenido).strip(". "),
                           "fuente": "Historial de la conversación", "interno": False})
    return partes


def texto_util_caso(row):
    return " ".join(normalizar(_texto_temas(p)) for p in fragmentos_caso(row))


def _texto_temas(parte):
    """Un catálogo de servicios al redirigir no identifica el producto del cliente."""
    contenido = parte["contenido"]
    if not _es_gestion(parte):
        return contenido

    def quitar_catalogo(match):
        catalogo = normalizar(match.group())
        if re.search(r"\btoken\s+(?:fisico|usb)\b", catalogo) and VIRTUAL.search(catalogo):
            return ""
        return match.group()

    return POLITICA_CANAL.sub("", CATALOGO_DERIVACION.sub(quitar_catalogo, contenido))


def _coincidencia(patron, valor):
    for match in patron.finditer(valor):
        previo = valor[max(0, match.start() - 65):match.start()]
        if not re.search(r"\b(?:no (?:usa|utiliza|tiene|es|se trata de)|sin)\s+(?:un |el )?$", previo):
            return match
    return None


def _evidencia(parte, match=None):
    contenido = parte["contenido"]
    # Las reglas operan sobre texto normalizado: la evidencia conserva el original.
    posicion = 0
    if match:
        contador, espacio, encontrado = 0, True, False
        for indice, caracter in enumerate(contenido):
            for c in unicodedata.normalize("NFKD", caracter).casefold():
                if unicodedata.combining(c):
                    continue
                if c.isspace():
                    if espacio:
                        continue
                espacio = c.isspace()
                if contador == match.start():
                    posicion, encontrado = indice, True
                    break
                contador += 1
            if encontrado:
                break
    inicio = max(0, posicion - 90)
    fragmento = contenido[inicio:inicio + 420]
    if inicio:
        fragmento = "…" + fragmento
    if inicio + 420 < len(contenido):
        fragmento += "…"
    fecha = texto(parte.get("fecha"))
    actor = texto(parte.get("usuario"))
    origen = " · ".join(v for v in (parte["fuente"], fecha, actor) if v)
    return f"{origen}: {fragmento}"


AUTORES_SOPORTE = {normalizar(nombre) for nombre in EQUIPO_SOPORTE_CASOS}


def _es_gestion(parte):
    if parte.get("interno"):
        return True
    tipo = normalizar(parte.get("tipo"))
    if "entrante" in tipo or "pendiente" in tipo or "fallido" in tipo or "omitido" in tipo:
        return False
    actor = normalizar(parte.get("usuario")).split("@")[0]
    actor = re.sub(r"[._-]", " ", actor)
    if tipo == "nota de trabajo":
        return actor in AUTORES_SOPORTE or not actor  # Nota interna sin autor identificable.
    return tipo == "correo saliente" or actor in AUTORES_SOPORTE or actor == "mesadeayuda"


def analizar_caso(row):
    partes = fragmentos_caso(row)
    normalizadas = [(parte, normalizar(parte["contenido"])) for parte in partes]
    temas = []
    for parte in partes:
        contenido = _texto_temas(parte)
        temas.append(({**parte, "contenido": contenido}, normalizar(contenido)))
    fisico = virtual = False
    pruebas_token = []
    for parte, contenido in temas:
        mf, mv = _coincidencia(FISICO, contenido), _coincidencia(VIRTUAL, contenido)
        if mf or mv:
            fisico |= bool(mf)
            virtual |= bool(mv)
            if len(pruebas_token) < 2:
                pruebas_token.append(_evidencia(parte, mf or mv))
    tipo = "Físico y virtual" if fisico and virtual else "Token físico" if fisico else "Token virtual" if virtual else "No identificado"
    motivo = "Sin motivo específico"
    prueba_motivo = ""
    # El problema reportado pesa antes que los temas de una guía de instalación.
    primarias = [(p, t) for p, t in temas if p["fuente"] in
                 {"descripcion", "descripcion_detallada", "sn__description"} or "entrante" in normalizar(p.get("tipo"))]
    for candidatos in (primarias, temas):
        for etiqueta, patron in MOTIVOS:
            hallazgo = next(((p, m) for p, t in candidatos if (m := _coincidencia(patron, t))), None)
            if hallazgo:
                motivo, prueba_motivo = etiqueta, _evidencia(*hallazgo)
                break
        if prueba_motivo:
            break
    atenciones = []
    destinos = []
    redirecciones = []
    causa, estado_causa, prueba_causa = SIN_CAUSA, "Pendiente de validar", ""
    causa_registrada = limpiar_mensaje(row.get("causa"))
    if causa_registrada and not (GENERICA.fullmatch(normalizar(causa_registrada)) or
            TEMA_TOKEN.fullmatch(normalizar(causa_registrada)) or normalizar(causa_registrada) in ETIQUETAS_GESTION):
        causa, estado_causa = causa_registrada, "Registrada en el caso"
        prueba_causa = f"Causa Raiz*: {causa_registrada}"
    for parte, contenido in normalizadas:
        if not _es_gestion(parte):
            continue
        redireccion = REDIRECCION.search(contenido)
        if redireccion and not re.search(r"\b(?:no se|no fue|sin)\s+(?:redirig|redireccion|remit|deriv|traslad|escal|envi)", contenido):
            # Asociar destino con la frase de redirección, no con cualquier mención del correo.
            contexto = contenido[max(0, redireccion.start() - 70):redireccion.end() + 180]
            encontrados = [nombre for nombre, patron in DESTINOS if patron.search(contexto)]
            if encontrados:
                destinos.extend(encontrados)
                redirecciones.append(_evidencia(parte, redireccion))
        for accion in ACCION_TECNICA.finditer(contenido):
            contexto = contenido[max(0, accion.start() - 70):accion.end() + 120]
            prefijo = contenido[max(0, accion.start() - 25):accion.start()]
            administrativo = GESTION_ADMINISTRATIVA.search(contenido[accion.start():accion.end() + 70])
            if CONTEXTO_TECNICO.search(contexto) and not administrativo and not re.search(r"\b(?:no|sin)\s*$", prefijo):
                atenciones.append(_evidencia(parte, accion))
                break
        if estado_causa == "Pendiente de validar" and (hallazgo := CAUSAL.search(contenido)):
            prefijo = contenido[max(0, hallazgo.start() - 30):hallazgo.start()]
            if re.search(r"\b(?:no|sin|desconocemos la)\s*$", prefijo):
                continue
            detalle = hallazgo.group("detalle").split(" --- ")[0].strip()
            if detalle and not GENERICA.fullmatch(detalle):
                # Conservar la explicación del agente; una acción exitosa sola no confirma la causa.
                causa, estado_causa = detalle, "Explicada en la atención"
                prueba_causa = _evidencia(parte, hallazgo)
    grupo_causa = causa
    if estado_causa != "Pendiente de validar":
        grupo_causa = next((nombre for nombre, patron in CAUSAS if patron.search(normalizar(causa))), causa)
    if destinos:
        atencion = "Redirigido con atención técnica registrada" if atenciones else REDIRIGIDO_SIN_ATENCION
    else:
        atencion = "Atención técnica registrada" if atenciones else SIN_ATENCION
    # Una clasificación principal por caso; producto y destino son atributos.
    clasificacion = CLASIFICACION_POR_MOTIVO.get(motivo, "Por revisar")
    if motivo == "Sin motivo específico":
        texto_principal = " ".join(t for _, t in primarias) or " ".join(t for _, t in temas)
        clasificacion = next((etiqueta for etiqueta, patron in CLASIFICACIONES_LEGADAS
                              if patron.search(texto_principal)), "Por revisar")
    anotacion = normalizar(causa_registrada)
    if anotacion in {"duplicado", "caso duplicado", "captores"}:
        clasificacion = "Captores" if anotacion == "captores" else "Duplicado"
    if atencion == REDIRIGIDO_SIN_ATENCION or normalizar(causa_registrada) in {
            "enviado a otros canales", "envio agenda / manual de uso", "9 - redireccionamiento agenda"}:
        if not atenciones:
            clasificacion = ENVIADO_OTROS_CANALES
    if clasificacion == ENVIADO_OTROS_CANALES:
        tipologia = ENVIADO_OTROS_CANALES
    elif clasificacion in {"Caída o indisponibilidad", "Falla del servicio", "Seguridad o suplantación"} or motivo in {
            "Falla de descarga o activación", "Falla al firmar", "Token no reconocido", "Bloqueo o PIN"}:
        tipologia = "Incidentes"
    elif clasificacion in {"Solicitud operativa", "Consulta o visualización de consumos", "Duplicado", "Captores"}:
        tipologia = "Solicitudes"
    elif clasificacion == "Por revisar":
        tipologia = "Por revisar"
    else:
        tipologia = "Soporte Uso"
    return {COL_TOKEN: tipo, COL_FISICO: fisico, COL_VIRTUAL: virtual,
            COL_CLASIFICACION: clasificacion,
            COL_TIPOLOGIA: tipologia,
            COL_MOTIVO: motivo, COL_CAUSA: causa, COL_ESTADO_CAUSA: estado_causa,
            "Grupo de causa": grupo_causa, COL_ATENCION: atencion,
            COL_DESTINO: ", ".join(dict.fromkeys(destinos)), COL_EVIDENCIA_TOKEN: " | ".join(pruebas_token),
            COL_EVIDENCIA_MOTIVO: prueba_motivo, COL_EVIDENCIA_CAUSA: prueba_causa,
            COL_EVIDENCIA_ATENCION: " | ".join((redirecciones[:1] + atenciones[:1])),
            "_version_analisis_historial": VERSION_ANALISIS}


def enviado_a_otros_canales(row, analisis=None):
    """Identifica la categoría de gestión conservando la causa técnica original."""
    if analisis is None:
        analisis = row if row.get("_version_analisis_historial") == VERSION_ANALISIS else analizar_caso(row)
    return analisis.get(COL_CLASIFICACION) == ENVIADO_OTROS_CANALES


def agregar_analisis_historial(df):
    if "_version_analisis_historial" in df and df["_version_analisis_historial"].eq(VERSION_ANALISIS).all():
        return df.copy()
    resultados = [analizar_caso(row) for _, row in df.iterrows()]
    columnas = list(analizar_caso({}))
    extras = pd.DataFrame(resultados, index=df.index, columns=columnas)
    return pd.concat([df.drop(columns=[col for col in columnas if col in df], errors="ignore"), extras], axis=1)


def analizar_evidencias_completas(df, registros):
    """Añade descripción y eventos locales con autor sin cambiar las filas originales."""
    trabajo = df.drop(columns=["_version_analisis_historial"], errors="ignore").copy()
    por_numero = {numero: (descripcion, eventos) for numero, descripcion, eventos in registros}
    descripciones, eventos = [], []
    for _, row in trabajo.iterrows():
        descripcion, historial = por_numero.get(row.get("numero"), (None, None))
        if isinstance(descripcion, dict):
            descripcion = descripcion.get("display_value") or descripcion.get("value")
        descripciones.append(texto(descripcion))
        eventos.append(historial if historial else row.get("historial_conversacion", []))
    trabajo["descripcion_detallada"] = descripciones
    trabajo["_historial_eventos"] = eventos
    return agregar_analisis_historial(trabajo)


def resumen_causas_historial(df):
    columnas = ["Grupo de causa", COL_ESTADO_CAUSA, "Casos", "% casos", "Ejemplos"]
    if df.empty:
        return pd.DataFrame(columns=columnas)
    base = df.drop_duplicates("numero") if "numero" in df else df
    grupos = base.groupby(["Grupo de causa", COL_ESTADO_CAUSA], dropna=False)
    filas = []
    for (causa, estado), grupo in grupos:
        ejemplos = ", ".join(grupo["numero"].astype(str).head(3)) if "numero" in grupo else ""
        filas.append({"Grupo de causa": causa, COL_ESTADO_CAUSA: estado, "Casos": len(grupo),
                      "% casos": round(len(grupo) * 100 / len(base), 2), "Ejemplos": ejemplos})
    return pd.DataFrame(filas, columns=columnas).sort_values(["Casos", "Grupo de causa"], ascending=[False, True])


def resumen_clasificacion_casos(df):
    """Distribución excluyente: el total de categorías coincide con casos únicos."""
    columnas = [COL_CLASIFICACION, "Casos", "% casos"]
    if df.empty:
        return pd.DataFrame(columns=columnas)
    base = df.drop_duplicates("numero", keep="last") if "numero" in df else df
    base = agregar_analisis_historial(base)
    resumen = base.groupby(COL_CLASIFICACION).size().rename("Casos").reset_index()
    resumen["% casos"] = (resumen["Casos"] * 100 / len(base)).round(2)
    return resumen.sort_values(["Casos", COL_CLASIFICACION], ascending=[False, True])[columnas]
