"""Cliente servidor de Table API, sin dependencias de interfaz o base de datos."""

import json
import hashlib
from copy import copy
import re
import math
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, wait, FIRST_COMPLETED
from queue import SimpleQueue, Empty
from html.parser import HTMLParser
from time import monotonic
from threading import RLock, Event
from urllib.parse import urlsplit
from datetime import date, datetime, timedelta, timezone
from zoneinfo import ZoneInfo

import requests
from requests.adapters import HTTPAdapter
from requests.auth import HTTPBasicAuth
from urllib3.util.retry import Retry

from core.settings import config_value


TABLES = {"casos": "sn_customerservice_case", "incidentes": "incident"}
PAGE_SIZE = 100
SOURCE_FIELDS = {
    "task_sla": "sys_id,task,sla,stage,start_time,planned_end_time,has_breached",
    "sys_journal_field": "sys_id,element_id,sys_created_on,sys_created_by,element,value",
    "sys_email": "sys_id,instance,sys_created_on,sys_created_by,type,origemail,sender,subject,recipients,body_text",
}
SOURCE_LABELS = {"sn_customerservice_case": "casos", "incident": "incidentes",
                 "task_sla": "SLAs", "sys_journal_field": "notas y comentarios",
                 "sys_email": "correos del historial"}
RELATED_SOURCES = (("task_sla", "task"), ("sys_journal_field", "element_id"), ("sys_email", "instance"))
_sesiones_api = {}
_sesiones_lock = RLock()
_tokens_oauth = {}
_tokens_lock = RLock()


class ServiceNowError(RuntimeError):
    """Error público sin URLs, credenciales ni cuerpo de respuesta remoto."""


class ServiceNowAuthenticationError(ServiceNowError):
    """El administrador debe volver a autenticar la conexión API."""


class ServiceNowQueryError(ServiceNowError):
    """Rechazo de Table API con estado HTTP, sin guardar el cuerpo remoto."""

    def __init__(self, mensaje, http_status):
        super().__init__(mensaje)
        self.http_status = http_status


class ServiceNowMFARequiredError(ServiceNowAuthenticationError):
    """La API pide al administrador renovar su verificación multifactor."""


class _ConsultaCancelada(Exception):
    """Detiene las siguientes páginas de una fuente que ya no se necesita."""


def _adaptador_http():
    return HTTPAdapter(max_retries=Retry(
        total=1, connect=1, read=0, status=1, backoff_factor=0.5, status_forcelist=[429, 502, 503, 504],
        allowed_methods=["GET"], respect_retry_after_header=False,
    ))


def validar_tipo(tipo):
    if tipo not in TABLES:
        raise ValueError("Tipo inválido: usa casos o incidentes.")
    return TABLES[tipo]


def validar_numero(numero, tipo="casos"):
    validar_tipo(tipo)
    patron = r"CS\d{5,}" if tipo == "casos" else r"INC\d{5,}"
    if not isinstance(numero, str) or not re.fullmatch(patron, numero):
        raise ValueError("Número de registro inválido.")
    return numero


def limites_mes_utc(mes):
    """El mes de negocio es colombiano; Table API filtra fechas brutas en UTC."""
    if not isinstance(mes, str) or not re.fullmatch(r"\d{4}-\d{2}", mes):
        raise ValueError("El mes debe tener formato AAAA-MM.")
    try:
        inicio = date.fromisoformat(mes + "-01")
        fin = date(inicio.year + (inicio.month == 12), inicio.month % 12 + 1, 1)
    except ValueError:
        raise ValueError("Mes inválido.") from None
    zona = ZoneInfo("America/Bogota")
    return tuple(datetime(d.year, d.month, 1, tzinfo=zona).astimezone(timezone.utc)
                 .strftime("%Y-%m-%d %H:%M:%S") for d in (inicio, fin))


def limites_rango_utc(inicio, fin):
    """Incluye ambos días de Colombia y usa un límite superior exclusivo en UTC."""
    try:
        inicio, fin = (date.fromisoformat(d) if isinstance(d, str) else d for d in (inicio, fin))
        if any(type(d) is not date for d in (inicio, fin)) or fin < inicio:
            raise ValueError
        fin = fin + timedelta(days=1)
        zona = ZoneInfo("America/Bogota")
        return tuple(datetime(d.year, d.month, d.day, tzinfo=zona).astimezone(timezone.utc)
                     .strftime("%Y-%m-%d %H:%M:%S") for d in (inicio, fin))
    except (ValueError, TypeError, OverflowError):
        raise ValueError("Selecciona una fecha inicial y final válidas, en ese orden.") from None


def valor(campo, legible=False):
    if isinstance(campo, dict):
        return (campo.get("display_value") or campo.get("value")) if legible else campo.get("value")
    return campo


def limpiar_texto(texto):
    return re.sub(r"\r\n|[\r\n\t]", " --- ", str(texto or ""))


class _TextoCorreoHTML(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.textos = []
        self.oculto = 0

    def handle_starttag(self, tag, attrs):
        if tag in {"script", "style"}:
            self.oculto += 1
        elif not self.oculto and tag in {"br", "p", "div", "li", "tr"}:
            self.textos.append("\n")

    def handle_endtag(self, tag):
        if tag in {"script", "style"}:
            self.oculto = max(0, self.oculto - 1)
        elif not self.oculto and tag in {"p", "div", "li", "tr"}:
            self.textos.append("\n")

    def handle_data(self, data):
        if not self.oculto:
            self.textos.append(data)


def texto_correo(row):
    texto = valor(row.get("body_text"))
    if texto and str(texto).strip():
        return limpiar_texto(texto)
    parser = _TextoCorreoHTML()
    parser.feed(str(valor(row.get("body")) or ""))
    return limpiar_texto("".join(parser.textos).strip())


def texto_historial(timeline):
    """Texto analítico calculado al sincronizar, sin HTML ni IDs técnicos."""
    return " --- ".join(" | ".join(str(r.get(k) or "") for k in
                                  ("fecha", "tipo", "usuario", "asunto", "contenido"))
                        for r in timeline if r.get("contenido") or r.get("asunto"))


def mensaje_rechazo_oauth(response):
    """Traduce solo códigos conocidos; nunca muestra el texto remoto del error."""
    try:
        payload = response.json()
        codigo = payload.get("error") if isinstance(payload, dict) else None
        denegado = isinstance(payload, dict) and payload.get("error_description") == "access_denied"
    except (ValueError, TypeError):
        codigo, denegado = None, False
    mensajes = {
        "invalid_client": "ServiceNow no reconoce el ID o el secreto de cliente. "
                          "Comprueba que ambos pertenecen a portalgestion y vuelve a copiarlos en los secretos del servidor.",
        "unauthorized_client": "ServiceNow no autoriza esta integración para usar credenciales de cliente. "
                               "Revisa que portalgestion esté activa y use la concesión Credenciales del cliente.",
        "unsupported_grant_type": "ServiceNow no admite la concesión client_credentials en esta instancia. "
                                  "Revisa que glide.oauth.inbound.client.credential.grant_type.enabled tenga el valor true.",
        "invalid_scope": "ServiceNow rechazó el alcance OAuth solicitado. "
                         "Revisa SERVICENOW_OAUTH_SCOPE y los alcances autorizados de portalgestion.",
        "invalid_grant": "ServiceNow rechazó la concesión OAuth. "
                         "Revisa el usuario de aplicación OAuth y la configuración de credenciales de cliente.",
        "invalid_request": "ServiceNow considera inválida la solicitud del token OAuth. "
                           "Consulta el registro OAuth de ServiceNow para identificar el parámetro rechazado.",
        "access_denied": "ServiceNow denegó la autorización OAuth. Revisa el usuario de aplicación OAuth de portalgestion, "
                         "que esté activo y que glide.oauth.inbound.client.credential.grant_type.enabled sea true. "
                         "El registro OAuth de ServiceNow permite identificar el motivo concreto.",
    }
    if codigo == "server_error" and denegado:
        codigo = "access_denied"
    if isinstance(codigo, str) and codigo in mensajes:
        return f"{mensajes[codigo]} (HTTP {response.status_code}; OAuth: {codigo})."
    return (f"ServiceNow rechazó la solicitud del token OAuth (HTTP {response.status_code}). "
            "Revisa el ID, el secreto, el usuario de aplicación OAuth y que las credenciales de cliente estén habilitadas. "
            "Consulta el registro OAuth de ServiceNow para identificar el motivo concreto.")


def mensaje_rechazo_consulta(response, tabla):
    """Reconoce motivos de Table API sin publicar correos ni texto remoto libre."""
    base = f"ServiceNow rechazó la consulta de {tabla} (HTTP {response.status_code})."
    try:
        payload = response.json()
        error = payload.get("error") if isinstance(payload, dict) else None
        message = error.get("message") if isinstance(error, dict) else None
    except (ValueError, TypeError):
        message = None
    motivos = {
        "Pagination not supported": " La instancia rechazó los enlaces de paginación de una consulta larga.",
        "Invalid table": " La tabla no está disponible mediante Table API; revisa su acceso por servicios web.",
        "Invalid query": " La instancia rechazó el filtro de consulta; revisa los nombres de sus campos.",
    }
    return base + (motivos.get(message, "") if isinstance(message, str) else "")


def normalizar_instancia(instance):
    """Admite el enlace del portal; conserva solo el dominio permitido para la API."""
    instance = str(instance or "").strip()
    if not instance:
        raise ServiceNowError("Falta la dirección de ServiceNow. Complétala en la pantalla de conexión del administrador.")
    if "://" in instance:
        try:
            parsed = urlsplit(instance)
            if parsed.scheme.lower() != "https" or parsed.username or parsed.password or parsed.port not in (None, 443):
                raise ValueError
            instance = parsed.hostname or ""
        except ValueError:
            raise ServiceNowError("Introduce un enlace HTTPS de tu instancia de ServiceNow, sin credenciales ni puertos adicionales.") from None
    instance = instance.lower()
    if not re.fullmatch(r"[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.service-now\.com", instance):
        raise ServiceNowError("Introduce el dominio de tu instancia terminado en .service-now.com o su enlace HTTPS.")
    return instance


class ServiceNowClient:
    def __init__(self, *, user=None, password=None, mfa_code=None, instance=None, auth_type=None, progress_callback=None):
        self.progress_callback = progress_callback
        self.total_registros = None
        self._lectores = {}
        self._executor = None
        instance = normalizar_instancia(config_value("SERVICENOW_INSTANCE") if instance is None else instance)
        mode = auth_type or ("basic" if user is not None or password is not None
                             else config_value("SERVICENOW_AUTH_TYPE", "basic"))
        self._oauth = str(mode).strip().lower() == "client_credentials"
        if str(mode).strip().lower() not in {"basic", "client_credentials"}:
            raise ServiceNowError("El método de conexión de ServiceNow no está admitido.")
        user = config_value("SERVICENOW_USER") if user is None else user
        password = config_value("SERVICENOW_PASSWORD") if password is None else password
        self._client_id = config_value("SERVICENOW_CLIENT_ID") if self._oauth else ""
        self._client_secret = config_value("SERVICENOW_CLIENT_SECRET") if self._oauth else ""
        if self._oauth and (not self._client_id or not self._client_secret):
            raise ServiceNowError("Faltan las credenciales de portalgestion en la configuración privada del servidor.")
        if not self._oauth and (not user or not password):
            raise ServiceNowError("No hay una cuenta de ServiceNow conectada. Introduce tu usuario y contraseña en Sincronizar Casos.")
        self._base_url = f"https://{instance}/api/now/table"
        self._token_url = f"https://{instance}/oauth_token.do"
        self._session = requests.Session()
        if not self._oauth:
            self._session.auth = HTTPBasicAuth(user, password)
        self._session.headers.update({"Accept": "application/json"})
        credenciales = (self._client_id, self._client_secret) if self._oauth else (user, password)
        self._scope = str(config_value("SERVICENOW_OAUTH_SCOPE", "")) if self._oauth else ""
        self._session_key = hashlib.sha256(f"{instance}\0{mode}\0{credenciales[0]}\0{credenciales[1]}\0{self._scope}".encode()).hexdigest()
        with _sesiones_lock:
            cookies = _sesiones_api.get(self._session_key)
            if cookies is not None:
                self._session.cookies.update(copy(cookies))
        if mfa_code and not self._oauth:
            if not re.fullmatch(r"\d{6,10}", mfa_code):
                self._session.close()
                raise ValueError("El código de verificación debe tener entre 6 y 10 dígitos.")
            self._session.headers["X-MFA_TOKEN"] = mfa_code
        self._session.mount("https://", _adaptador_http())

    def __enter__(self):
        return self

    def __exit__(self, *args):
        try:
            if self._executor is not None:
                self._executor.shutdown(wait=True, cancel_futures=True)
        finally:
            for lector in self._lectores.values():
                lector._session.close()
            self._session.close()

    def validar_acceso(self):
        """Comprueba autenticación y acceso HTTP a las cuatro fuentes de casos."""
        for tabla in (TABLES["casos"], "task_sla", "sys_journal_field", "sys_email"):
            # Consultar un ID vacío evita descargar datos durante el inicio de sesión.
            list(self._paginas(tabla, "sys_id=" + "0" * 32))

    def recordar_sesion(self):
        """Solo las cookies de sesión API viven en RAM; el código MFA no se almacena."""
        if self._oauth:
            return
        with _sesiones_lock:
            _sesiones_api.clear()
            _sesiones_api[self._session_key] = copy(self._session.cookies)
        self._session.headers.pop("X-MFA_TOKEN", None)

    def _obtener_token(self, rechazado=None):
        # Un solo POST por renovación, compartido por los clientes del proceso.
        with _tokens_lock:
            actual = _tokens_oauth.get(self._session_key)
            if actual and actual[1] > monotonic() and actual[0] != rechazado:
                return actual[0]
            _tokens_oauth.pop(self._session_key, None)
            datos = {"grant_type": "client_credentials", "client_id": self._client_id,
                     "client_secret": self._client_secret}
            if self._scope:
                datos["scope"] = self._scope
            inicio = monotonic()
            try:
                response = self._session.post(self._token_url, data=datos, timeout=(10, 30), allow_redirects=False)
                if response.status_code in {400, 401, 403}:
                    raise ServiceNowAuthenticationError(mensaje_rechazo_oauth(response))
                if response.status_code != 200:
                    raise ServiceNowError(f"No se pudo obtener autorización de ServiceNow (HTTP {response.status_code}).")
                payload = response.json()
                token = payload.get("access_token")
                duracion = float(payload.get("expires_in", 0))
                if (not isinstance(token, str) or not re.fullmatch(r"[A-Za-z0-9._~+/-]+=*", token)
                        or not math.isfinite(duracion) or duracion <= 0
                        or str(payload.get("token_type", "Bearer")).lower() != "bearer"):
                    raise ValueError
            except (requests.RequestException, ValueError, TypeError, AttributeError):
                raise ServiceNowError("No se pudo obtener una autorización válida de ServiceNow.") from None
            if len(_tokens_oauth) >= 16:
                _tokens_oauth.clear()
            _tokens_oauth[self._session_key] = (token, inicio + duracion - min(60, duracion * 0.1))
            return token

    def _consultar_api(self, url, **kwargs):
        if not self._oauth:
            return self._session.get(url, **kwargs)
        token = self._obtener_token()
        response = self._session.get(url, headers={"Authorization": f"Bearer {token}"}, **kwargs)
        if response.status_code == 401:
            token = self._obtener_token(rechazado=token)
            response = self._session.get(url, headers={"Authorization": f"Bearer {token}"}, **kwargs)
        return response

    def _informar_consulta(self, mensaje):
        if self.progress_callback:
            self.progress_callback(0, mensaje)

    def _paginas(self, tabla, query, *, fields=None):
        offset = 0
        cursor = None
        pagina = 0
        # No se interpolan tablas ni filtros procedentes del cliente HTTP.
        if tabla not in {*TABLES.values(), "task_sla", "sys_journal_field", "sys_email"}:
            raise ValueError("Tabla no permitida.")
        seen = set()
        while True:
            pagina += 1
            etiqueta = SOURCE_LABELS[tabla]
            self._informar_consulta(f"Consultando {etiqueta} · página {pagina}...")
            try:
                filtro = query + (f"^sys_id>{cursor}" if cursor else "")
                params = {"sysparm_query": filtro + "^ORDERBYsys_id",
                          "sysparm_display_value": "all" if tabla in TABLES.values() or (tabla == "task_sla" and fields is None) else "false",
                          "sysparm_exclude_reference_link": "true",
                          "sysparm_suppress_pagination_header": "true",
                          "sysparm_limit": PAGE_SIZE, "sysparm_offset": offset}
                proyeccion = fields or SOURCE_FIELDS.get(tabla)
                if proyeccion:
                    params["sysparm_fields"] = proyeccion
                response = self._consultar_api(
                    f"{self._base_url}/{tabla}",
                    params=params,
                    timeout=(10, 30), allow_redirects=False,
                )
                if response.status_code != 200:
                    if response.status_code == 401 and response.headers.get("X-MFA_TOKEN") == "invalid":
                        raise ServiceNowMFARequiredError(
                            "ServiceNow requiere un código de verificación vigente. Un administrador debe reconectar."
                        )
                    if response.status_code == 401:
                        if self._oauth:
                            raise ServiceNowAuthenticationError(
                                "ServiceNow no aceptó la autorización de la integración tras renovarla. Revisa su configuración de acceso API."
                            )
                        raise ServiceNowAuthenticationError(
                            "ServiceNow rechazó el inicio de sesión por API (HTTP 401). Comprueba tu usuario y contraseña de ServiceNow. "
                            "Si accedes con inicio de sesión corporativo o biometría, puede requerirse otra autorización para la API."
                        )
                    if response.status_code == 403:
                        raise ServiceNowError(
                            f"ServiceNow no permite consultar {tabla} por API (HTTP 403). Revisa los permisos de lectura y la política de acceso API."
                        )
                    raise ServiceNowQueryError(mensaje_rechazo_consulta(response, tabla), response.status_code)
                result = response.json().get("result")
                if not isinstance(result, list) or any(not isinstance(row, dict) for row in result):
                    raise ServiceNowError(f"ServiceNow devolvió datos inválidos para {tabla}.")
                total = response.headers.get("X-Total-Count")
                total = int(total) if total is not None else None
                if total is not None and total < 0:
                    raise ValueError
                if tabla in TABLES.values() and pagina == 1:
                    self.total_registros = total
                    if total is not None:
                        self._informar_consulta(f"La API reporta {total} {etiqueta} para esta consulta.")
                with _sesiones_lock:
                    if self._session_key in _sesiones_api:
                        _sesiones_api[self._session_key] = copy(self._session.cookies)
            except requests.Timeout:
                raise ServiceNowError(
                    f"ServiceNow tardó demasiado al consultar {etiqueta}. "
                    "La consulta se interrumpió; puedes volver a sincronizar."
                ) from None
            except (requests.RequestException, ValueError, AttributeError):
                raise ServiceNowError(f"No se pudo consultar {tabla} en ServiceNow.") from None
            for row in result:
                identity = valor(row.get("sys_id"))
                if (not isinstance(identity, str) or not re.fullmatch(r"[0-9a-fA-F]{32}", identity)
                        or identity in seen or (cursor and identity <= cursor)):
                    raise ServiceNowError(f"Paginación inconsistente en {tabla}; vuelve a sincronizar.")
                seen.add(identity)
            if result:
                if tabla == "sys_email" and fields is None:
                    self._completar_correos_html(result)
                self._informar_consulta(f"Recibidos {len(result)} registros de {etiqueta} · página {pagina}.")
                yield result
            if total is not None and offset + len(result) >= total:
                break
            if result and (cursor or len(result) < PAGE_SIZE):
                # Una página corta puede deberse a ACLs o a un límite del servidor.
                # Continuar tras el último ID recibido evita saltar casos visibles.
                cursor = max(valor(row["sys_id"]) for row in result)
                offset = 0
            else:
                # Si una ventana entera queda oculta por ACLs, avanzar dentro del
                # filtro actual hasta encontrar IDs visibles o agotar su total.
                offset += PAGE_SIZE
                if (total is not None and offset >= total) or (total is None and not result):
                    break

    def _completar_correos_html(self, rows):
        """El HTML solo se descarga cuando el correo no contiene una versión de texto."""
        pendientes = {valor(row.get("sys_id")): row for row in rows
                      if not str(valor(row.get("body_text")) or "").strip() and not valor(row.get("body"))}
        if not pendientes:
            return
        self._informar_consulta("Recuperando el contenido de correos sin versión de texto...")
        for page in self._paginas("sys_email", "sys_idIN" + ",".join(pendientes), fields="sys_id,body"):
            for row in page:
                original = pendientes.get(valor(row.get("sys_id")))
                if original is not None:
                    original["body"] = row.get("body", "")

    def consolidar(self, raw, tipo="casos", relacionados=None):
        tabla = validar_tipo(tipo)
        sys_id = valor(raw.get("sys_id"))
        if not isinstance(sys_id, str) or not re.fullmatch(r"[0-9a-fA-F]{32}", sys_id):
            raise ServiceNowError("ServiceNow devolvió un sys_id inválido.")
        validar_numero(valor(raw.get("number")), tipo)
        def paginas(tabla_fuente, query):
            if relacionados is not None:
                return [relacionados.get(tabla_fuente, [])]
            return self._paginas(tabla_fuente, query)
        slas = []
        for page in paginas("task_sla", f"task={sys_id}"):
            for row in page:
                sla = {key: valor(row.get(key), legible=key in {"sla", "stage"})
                       for key in ("sla", "stage", "start_time", "planned_end_time", "has_breached")}
                sla["has_breached"] = str(sla["has_breached"]).lower() == "true"
                sla["sys_id"] = valor(row.get("sys_id"))
                slas.append(sla)
        timeline = []
        for page in paginas("sys_journal_field", f"element_id={sys_id}"):
            for row in page:
                timeline.append({"id": valor(row.get("sys_id")), "fecha": valor(row.get("sys_created_on")), "origen": "Journal",
                                 "tipo": "Nota de Trabajo" if valor(row.get("element")) == "work_notes" else "Comentario",
                                 "usuario": valor(row.get("sys_created_by")),
                                 "contenido": limpiar_texto(valor(row.get("value")))})
        for page in paginas("sys_email", f"target_table={tabla}^instance={sys_id}"):
            for row in page:
                timeline.append({"id": valor(row.get("sys_id")), "fecha": valor(row.get("sys_created_on")), "origen": "Email",
                                 "tipo": {"received": "Correo Entrante", "received-ignored": "Correo Entrante (omitido)",
                                          "sent": "Correo Saliente", "send-ready": "Correo Pendiente", "send-failed": "Correo Fallido"}
                                          .get(valor(row.get("type")), "Correo"),
                                 "usuario": valor(row.get("origemail")) or valor(row.get("sender")) or valor(row.get("sys_created_by")),
                                 "asunto": limpiar_texto(valor(row.get("subject"))),
                                 "destinatarios": valor(row.get("recipients")),
                                 "contenido": texto_correo(row)})
        timeline.sort(key=lambda row: (str(row["fecha"] or ""), str(row["id"] or "")))
        return {"case_id": sys_id, "tabla": tabla,
                "campos_caso": {key: valor(field, legible=True) for key, field in raw.items()},
                "campos_raw": raw, "slas": slas, "traza_unificada": timeline,
                "historial_conversacion": timeline}

    def _paginas_referencias(self, fuente, referencia, ids, tabla):
        if not ids:
            return
        prefijo = f"target_table={tabla}^" if fuente == "sys_email" else ""
        query = f"{prefijo}{referencia}IN{','.join(ids)}"
        try:
            # No entregar resultados parciales si hay que reintentar el lote.
            paginas = list(self._paginas(fuente, query))
        except ServiceNowQueryError as error:
            if error.http_status != 400 or len(ids) <= 1:
                raise
            mitad = len(ids) // 2
            for parte in (ids[:mitad], ids[mitad:]):
                yield from self._paginas_referencias(fuente, referencia, parte, tabla)
            return
        yield from paginas

    def _consolidar_pagina(self, rows, tipo="casos"):
        tabla = validar_tipo(tipo)
        ids = [valor(row.get("sys_id")) for row in rows]
        if any(not isinstance(s, str) or not re.fullmatch(r"[0-9a-fA-F]{32}", s) for s in ids):
            raise ServiceNowError("ServiceNow devolvió un sys_id inválido.")
        grupos = {s: defaultdict(list) for s in ids}
        if self._oauth:
            fuentes = self._leer_fuentes_paralelas(ids, tabla)
        else:
            # Las sesiones antiguas con cookies/MFA conservan su secuencia.
            fuentes = {fuente: self._paginas_referencias(fuente, referencia, ids, tabla)
                       for fuente, referencia in RELATED_SOURCES}
        for fuente, referencia in RELATED_SOURCES:
            for page in fuentes[fuente]:
                for row in page:
                    referencia_id = valor(row.get(referencia))
                    if referencia_id in grupos:
                        grupos[referencia_id][fuente].append(row)
        return [self.consolidar(row, tipo, relacionados=grupos[valor(row["sys_id"])]) for row in rows]

    def _lector_fuente(self, fuente):
        if fuente not in self._lectores:
            # Crear el lector con la configuración ya validada en el hilo principal:
            # ningún trabajador accede a st.secrets ni comparte una sesión mutable.
            lector = object.__new__(ServiceNowClient)
            for nombre in ("_oauth", "_base_url", "_token_url", "_client_id", "_client_secret", "_scope", "_session_key"):
                setattr(lector, nombre, getattr(self, nombre))
            lector.total_registros = None
            lector._session = requests.Session()
            lector._session.headers.update(self._session.headers)
            lector._session.mount("https://", _adaptador_http())
            self._lectores[fuente] = lector
        return self._lectores[fuente]

    def _leer_fuentes_paralelas(self, ids, tabla):
        if self._executor is None:
            self._executor = ThreadPoolExecutor(max_workers=3, thread_name_prefix="servicenow-historial")
        mensajes = SimpleQueue()
        cancelar = Event()

        def progreso(_, mensaje):
            if cancelar.is_set():
                raise _ConsultaCancelada
            mensajes.put(mensaje)

        futuros = {}
        resultados = {}
        try:
            for fuente, referencia in RELATED_SOURCES:
                lector = self._lector_fuente(fuente)
                lector.progress_callback = progreso
                futuros[self._executor.submit(lambda l=lector, f=fuente, r=referencia:
                    list(l._paginas_referencias(f, r, ids, tabla)))] = fuente
            ultima = monotonic()
            while futuros:
                terminados, _ = wait(futuros, timeout=0.2, return_when=FIRST_COMPLETED)
                for futuro in terminados:
                    resultados[futuros.pop(futuro)] = futuro.result()
                # Solo el hilo principal actualiza Streamlit o la conexión que
                # coordina el turno; los trabajadores entregan mensajes en cola.
                while True:
                    try:
                        mensaje = mensajes.get_nowait()
                    except Empty:
                        break
                    self._informar_consulta(mensaje)
                    ultima = monotonic()
                if futuros and monotonic() - ultima >= 1:
                    fuentes = ", ".join(SOURCE_LABELS[f] for f in futuros.values())
                    self._informar_consulta(f"Esperando {fuentes} del lote actual...")
                    ultima = monotonic()
            return resultados
        except BaseException:
            # No dejar consultas huérfanas al cancelar desde Streamlit o ceder el
            # turno. Las peticiones ya enviadas finalizan con su timeout normal.
            cancelar.set()
            for futuro in futuros:
                futuro.cancel()
            wait(futuros)
            raise

    def obtener(self, numero, tipo="casos"):
        tabla = validar_tipo(tipo)
        validar_numero(numero, tipo)
        for page in self._paginas(tabla, f"number={numero}"):
            return self.consolidar(page[0], tipo)
        return None

    def listar_mes(self, mes, tipo="casos"):
        tabla = validar_tipo(tipo)
        inicio, fin = limites_mes_utc(mes)
        for page in self._paginas(tabla, f"sys_created_on>={inicio}^sys_created_on<{fin}"):
            yield self._consolidar_pagina(page, tipo)

    def listar_rango(self, inicio, fin, tipo="casos"):
        tabla = validar_tipo(tipo)
        desde, hasta = limites_rango_utc(inicio, fin)
        for page in self._paginas(tabla, f"sys_created_on>={desde}^sys_created_on<{hasta}"):
            yield self._consolidar_pagina(page, tipo)

    def listar_casos_actualizados(self, desde, hasta):
        """Incluye actividad nueva de casos antiguos aunque el caso no cambie."""
        if desde.tzinfo is None or hasta.tzinfo is None or desde > hasta:
            raise ValueError("Intervalo de sincronización inválido.")
        inicio, fin = (d.astimezone(timezone.utc).strftime("%Y-%m-%d %H:%M:%S") for d in (desde, hasta))
        rango = f"sys_updated_on>={inicio}^sys_updated_on<{fin}"
        vistos = set()
        for page in self._paginas(TABLES["casos"], rango):
            documentos = self._consolidar_pagina(page)
            vistos.update(doc["case_id"] for doc in documentos)
            yield documentos
        relacionados = set()
        fuentes = (("task_sla", "task", rango),
                   ("sys_journal_field", "element_id", f"name={TABLES['casos']}^{rango}"),
                   ("sys_email", "instance", f"target_table={TABLES['casos']}^{rango}"))
        for tabla, referencia, query in fuentes:
            for page in self._paginas(tabla, query, fields=f"sys_id,{referencia}"):
                for row in page:
                    sys_id = valor(row.get(referencia))
                    if isinstance(sys_id, str) and re.fullmatch(r"[0-9a-fA-F]{32}", sys_id):
                        relacionados.add(sys_id)
        pendientes = sorted(relacionados - vistos)
        for offset in range(0, len(pendientes), PAGE_SIZE):
            ids = ",".join(pendientes[offset:offset + PAGE_SIZE])
            for page in self._paginas(TABLES["casos"], f"sys_idIN{ids}"):
                yield self._consolidar_pagina(page)


def obtener_traza_caso_servicenow(numero_caso):
    """Consulta interna del servidor; los consumidores públicos usan el proxy autorizado."""
    with ServiceNowClient() as client:
        return client.obtener(numero_caso)


CASE_FIELDS = {
    "numero": ("number",), "descripcion": ("short_description",), "estado": ("state",),
    "creado_por": ("sys_created_by", "opened_by"), "contacto": ("contact",), "cuenta": ("account",),
    "asignado": ("assigned_to",), "prioridad": ("priority",), "actualizado": ("sys_updated_on",),
    "creado": ("sys_created_on",), "cerrado": ("closed_at",), "canal": ("contact_type",),
    "escalado_proveedor": ("u_escalado_a_proveedor",), "fecha_vencimiento_sla": ("sla_due",),
    "producto": ("u_producto", "product"), "codigo_resolucion": ("close_code", "resolution_code"),
    "notas_resolucion": ("close_notes", "resolution_notes"), "causa": ("u_causa_raiz",),
    "nombre_proveedor": ("u_nombre_del_proveedor",), "numero_caso_externo": ("u_numero_de_caso_externo",),
    "observaciones_adicionales": ("u_observaciones_adicionales",),
    "sla_creado": ("u_sla_creado", "made_sla"),
    "observaciones_trabajo": ("comments_and_work_notes",),
    "notas_trabajo": ("work_notes",),
    "caso_referencia": ("case",),
}
INCIDENT_FIELDS = {
    "numero": ("number",), "solicitante": ("caller_id",), "breve_descripcion": ("short_description",),
    "descripcion": ("description",), "categoria": ("category",), "estado": ("state",),
    "prioridad": ("priority",), "grupo_asignacion": ("assignment_group",), "asignado_a": ("assigned_to",),
    "empresa": ("company",), "creado_por": ("sys_created_by", "opened_by"),
    "creado": ("sys_created_on",), "cerrado": ("closed_at",), "fecha_vencimiento_sla": ("sla_due",),
    "servicio_negocio": ("business_service",), "impacto": ("impact",),
    "escalado_proveedor": ("u_escalado_a_proveedor",), "nombre_proveedor": ("u_nombre_del_proveedor",),
    "tipo_falla": ("u_tipo_de_falla",),
}
DATE_COLUMNS = {"creado", "cerrado", "actualizado", "fecha_vencimiento_sla"}


def mapear_documento(documento, tipo="casos"):
    """Conserva campos desconocidos en JSON y admite nombres personalizados de la instancia."""
    validar_tipo(tipo)
    fields = dict(CASE_FIELDS if tipo == "casos" else INCIDENT_FIELDS)
    env_name = "SERVICENOW_CASE_FIELD_MAP" if tipo == "casos" else "SERVICENOW_INCIDENT_FIELD_MAP"
    try:
        overrides = json.loads(config_value(env_name, "{}"))
        if not isinstance(overrides, dict) or any(not isinstance(k, str) or not isinstance(v, str)
                                                for k, v in overrides.items()):
            raise ValueError
    except (ValueError, TypeError):
        raise ServiceNowError(f"{env_name} debe contener un objeto JSON de nombres de campos.") from None
    fields.update({key: (field,) for key, field in overrides.items()})
    raw = documento["campos_raw"]
    mapped = {}
    for column, candidates in fields.items():
        presentes = [field for field in candidates if field in raw]
        # Un campo personalizado vacío no debe ocultar su alternativa estándar
        # cuando ambas claves vienen en la respuesta completa de la API.
        field = next((field for field in presentes if valor(raw[field]) not in (None, "")),
                     presentes[0] if presentes else None)
        if field is None:
            continue
        content = valor(raw[field], legible=column not in DATE_COLUMNS and column != "numero")
        if column in DATE_COLUMNS and content:
            try:
                content = datetime.fromisoformat(str(content).replace("Z", "+00:00"))
                if content.tzinfo is None:
                    content = content.replace(tzinfo=timezone.utc)
                content = content.astimezone(ZoneInfo("America/Bogota")).strftime("%Y-%m-%d %H:%M:%S")
            except ValueError:
                raise ServiceNowError(f"Fecha inválida en el campo {field}.") from None
        # Mantener False/0: son valores reales de origen, no campos vacíos.
        mapped[column] = "" if content is None else content
    timeline = documento.get("historial_conversacion", documento.get("traza_unificada", []))
    traza = texto_historial(timeline)
    notas = " --- ".join(r["contenido"] for r in timeline if r["tipo"] == "Nota de Trabajo")
    # No vaciar las notas históricas de Excel cuando la API no entrega actividad.
    if traza:
        mapped["historial_conversacion" if tipo == "casos" else "observaciones_trabajo"] = traza
    if notas:
        mapped["notas_trabajo" if tipo == "casos" else "lista_notas_trabajo"] = notas
    if tipo == "incidentes" and traza:
        mapped["actualizaciones"] = traza
    if "numero" not in mapped:
        raise ServiceNowError("Falta el número del registro para guardar los datos.")
    validar_numero(mapped["numero"], tipo)
    return mapped
