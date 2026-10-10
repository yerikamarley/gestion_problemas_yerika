"""Cliente servidor de Table API, sin dependencias de interfaz o base de datos."""

import json
import hashlib
from copy import copy
import re
from threading import RLock
from datetime import date, datetime, timezone
from zoneinfo import ZoneInfo

import requests
from requests.adapters import HTTPAdapter
from requests.auth import HTTPBasicAuth
from urllib3.util.retry import Retry

from core.settings import config_value


TABLES = {"casos": "sn_customerservice_case", "incidentes": "incident"}
PAGE_SIZE = 100
_sesiones_api = {}
_sesiones_lock = RLock()


class ServiceNowError(RuntimeError):
    """Error público sin URLs, credenciales ni cuerpo de respuesta remoto."""


class ServiceNowMFARequiredError(ServiceNowError):
    """La API pide al administrador renovar su verificación multifactor."""


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


def valor(campo, legible=False):
    if isinstance(campo, dict):
        return (campo.get("display_value") or campo.get("value")) if legible else campo.get("value")
    return campo


def limpiar_texto(texto):
    return re.sub(r"\r\n|[\r\n\t]", " --- ", str(texto or ""))


class ServiceNowClient:
    def __init__(self, *, user=None, password=None, mfa_code=None):
        instance = str(config_value("SERVICENOW_INSTANCE")).strip()
        user = config_value("SERVICENOW_USER") if user is None else user
        password = config_value("SERVICENOW_PASSWORD") if password is None else password
        if not instance or not user or not password:
            raise ServiceNowError("Configura SERVICENOW_INSTANCE, SERVICENOW_USER y SERVICENOW_PASSWORD en el servidor.")
        if not re.fullmatch(r"[A-Za-z0-9](?:[A-Za-z0-9-]*[A-Za-z0-9])?\.service-now\.com", instance):
            raise ServiceNowError("SERVICENOW_INSTANCE debe ser un dominio de instancia de ServiceNow, sin protocolo ni ruta.")
        self._base_url = f"https://{instance}/api/now/table"
        self._session = requests.Session()
        self._session.auth = HTTPBasicAuth(user, password)
        self._session.headers.update({"Accept": "application/json"})
        self._session_key = hashlib.sha256(f"{instance}\0{user}\0{password}".encode()).hexdigest()
        with _sesiones_lock:
            cookies = _sesiones_api.get(self._session_key)
            if cookies is not None:
                self._session.cookies.update(copy(cookies))
        if mfa_code:
            if not re.fullmatch(r"\d{6,10}", mfa_code):
                self._session.close()
                raise ValueError("El código de verificación debe tener entre 6 y 10 dígitos.")
            self._session.headers["X-MFA_TOKEN"] = mfa_code
        self._session.mount("https://", HTTPAdapter(max_retries=Retry(
            total=3, backoff_factor=0.5, status_forcelist=[429, 502, 503, 504],
            allowed_methods=["GET"], respect_retry_after_header=False,
        )))

    def __enter__(self):
        return self

    def __exit__(self, *args):
        self._session.close()

    def validar_acceso(self):
        """Comprueba autenticación y acceso HTTP a las cuatro fuentes de casos."""
        for tabla in (TABLES["casos"], "task_sla", "sys_journal_field", "sys_email"):
            # Consultar un ID vacío evita descargar datos durante el inicio de sesión.
            list(self._paginas(tabla, "sys_id=" + "0" * 32))

    def recordar_sesion(self):
        """Solo las cookies de sesión API viven en RAM; el código MFA no se almacena."""
        with _sesiones_lock:
            _sesiones_api.clear()
            _sesiones_api[self._session_key] = copy(self._session.cookies)
        self._session.headers.pop("X-MFA_TOKEN", None)

    def _paginas(self, tabla, query):
        offset = 0
        # No se interpolan tablas ni filtros procedentes del cliente HTTP.
        if tabla not in {*TABLES.values(), "task_sla", "sys_journal_field", "sys_email"}:
            raise ValueError("Tabla no permitida.")
        seen = set()
        while True:
            try:
                response = self._session.get(
                    f"{self._base_url}/{tabla}",
                    params={"sysparm_query": query + "^ORDERBYsys_id",
                            "sysparm_display_value": "all", "sysparm_exclude_reference_link": "true",
                            "sysparm_limit": PAGE_SIZE, "sysparm_offset": offset},
                    timeout=(10, 60), allow_redirects=False,
                )
                if response.status_code != 200:
                    if response.status_code == 401 and response.headers.get("X-MFA_TOKEN") == "invalid":
                        raise ServiceNowMFARequiredError(
                            "ServiceNow requiere un código de verificación vigente. Un administrador debe reconectar."
                        )
                    raise ServiceNowError(f"ServiceNow rechazó la consulta de {tabla} (HTTP {response.status_code}).")
                result = response.json().get("result")
                if not isinstance(result, list) or any(not isinstance(row, dict) for row in result):
                    raise ServiceNowError(f"ServiceNow devolvió datos inválidos para {tabla}.")
                total = response.headers.get("X-Total-Count")
                total = int(total) if total is not None else None
                if total is not None and total < 0:
                    raise ValueError
                with _sesiones_lock:
                    if self._session_key in _sesiones_api:
                        _sesiones_api[self._session_key] = copy(self._session.cookies)
            except (requests.RequestException, ValueError, AttributeError):
                raise ServiceNowError(f"No se pudo consultar {tabla} en ServiceNow.") from None
            for row in result:
                identity = valor(row.get("sys_id"))
                if not identity or identity in seen:
                    raise ServiceNowError(f"Paginación inconsistente en {tabla}; vuelve a sincronizar.")
                seen.add(identity)
            if result:
                yield result
            offset += PAGE_SIZE
            # ACLs pueden devolver páginas cortas/vacías antes del último offset.
            if (total is not None and offset >= total) or (total is None and not result):
                break

    def consolidar(self, raw, tipo="casos"):
        tabla = validar_tipo(tipo)
        sys_id = valor(raw.get("sys_id"))
        if not isinstance(sys_id, str) or not re.fullmatch(r"[0-9a-fA-F]{32}", sys_id):
            raise ServiceNowError("ServiceNow devolvió un sys_id inválido.")
        validar_numero(valor(raw.get("number")), tipo)
        slas = []
        for page in self._paginas("task_sla", f"task={sys_id}"):
            for row in page:
                sla = {key: valor(row.get(key), legible=key in {"sla", "stage"})
                       for key in ("sla", "stage", "start_time", "planned_end_time", "has_breached")}
                sla["has_breached"] = str(sla["has_breached"]).lower() == "true"
                sla["sys_id"] = valor(row.get("sys_id"))
                slas.append(sla)
        timeline = []
        for page in self._paginas("sys_journal_field", f"element_id={sys_id}"):
            for row in page:
                timeline.append({"id": valor(row.get("sys_id")), "fecha": valor(row.get("sys_created_on")), "origen": "Journal",
                                 "tipo": "Nota de Trabajo" if valor(row.get("element")) == "work_notes" else "Comentario",
                                 "usuario": valor(row.get("sys_created_by")),
                                 "contenido": limpiar_texto(valor(row.get("value")))})
        for page in self._paginas("sys_email", f"target_table={tabla}^instance={sys_id}"):
            for row in page:
                timeline.append({"id": valor(row.get("sys_id")), "fecha": valor(row.get("sys_created_on")), "origen": "Email",
                                 "tipo": "Correo Entrante" if valor(row.get("type")) == "received" else "Correo Saliente",
                                 "usuario": valor(row.get("origemail")) or valor(row.get("sys_created_by")),
                                 "contenido": limpiar_texto(valor(row.get("body_text")) or "Sin texto")})
        timeline.sort(key=lambda row: str(row["fecha"] or ""))
        return {"case_id": sys_id, "tabla": tabla,
                "campos_caso": {key: valor(field, legible=True) for key, field in raw.items()},
                "campos_raw": raw, "slas": slas, "traza_unificada": timeline}

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
            yield [self.consolidar(row, tipo) for row in page]

    def listar_casos_actualizados(self, desde, hasta):
        """Incluye actividad nueva de casos antiguos aunque el caso no cambie."""
        if desde.tzinfo is None or hasta.tzinfo is None or desde > hasta:
            raise ValueError("Intervalo de sincronización inválido.")
        inicio, fin = (d.astimezone(timezone.utc).strftime("%Y-%m-%d %H:%M:%S") for d in (desde, hasta))
        rango = f"sys_updated_on>={inicio}^sys_updated_on<{fin}"
        vistos = set()
        for page in self._paginas(TABLES["casos"], rango):
            documentos = [self.consolidar(row) for row in page]
            vistos.update(doc["case_id"] for doc in documentos)
            yield documentos
        relacionados = set()
        fuentes = (("task_sla", "task", rango),
                   ("sys_journal_field", "element_id", f"name={TABLES['casos']}^{rango}"),
                   ("sys_email", "instance", f"target_table={TABLES['casos']}^{rango}"))
        for tabla, referencia, query in fuentes:
            for page in self._paginas(tabla, query):
                for row in page:
                    sys_id = valor(row.get(referencia))
                    if isinstance(sys_id, str) and re.fullmatch(r"[0-9a-fA-F]{32}", sys_id):
                        relacionados.add(sys_id)
        pendientes = sorted(relacionados - vistos)
        for offset in range(0, len(pendientes), PAGE_SIZE):
            ids = ",".join(pendientes[offset:offset + PAGE_SIZE])
            for page in self._paginas(TABLES["casos"], f"sys_idIN{ids}"):
                yield [self.consolidar(row) for row in page]


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
DATE_COLUMNS = {"creado", "cerrado", "actualizado", "fecha_vencimiento_sla", "sla_creado"}


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
        field = next((field for field in candidates if field in raw), None)
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
        mapped[column] = content or ""
    timeline = documento["traza_unificada"]
    traza = " --- ".join(f"{r['fecha']} | {r['tipo']} | {r['usuario']} | {r['contenido']}" for r in timeline)
    notas = " --- ".join(r["contenido"] for r in timeline if r["tipo"] == "Nota de Trabajo")
    # No vaciar las notas históricas de Excel cuando la API no entrega actividad.
    if traza:
        mapped["observaciones_trabajo"] = traza
    if notas:
        mapped["notas_trabajo" if tipo == "casos" else "lista_notas_trabajo"] = notas
    if tipo == "incidentes" and traza:
        mapped["actualizaciones"] = traza
    if "numero" not in mapped:
        raise ServiceNowError("Falta el número del registro para guardar los datos.")
    validar_numero(mapped["numero"], tipo)
    return mapped
