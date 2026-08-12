"""Matriz central y funciones puras de autorización por rol.

Este módulo no consulta sesiones ni base de datos. El rol ``viewer`` se conserva
temporalmente como rol heredado para facilitar la migración de usuarios.
"""

from types import MappingProxyType


ROLE_ADMIN = "admin"
ROLE_SOPORTE = "soporte"
ROLE_EXPERIENCIA = "experiencia"
ROLE_GERENCIAS = "gerencias"
ROLE_AREA_IA = "area_ia"
ROLE_VIEWER_LEGACY = "viewer"

ACTION_MANAGE_USERS = "manage_users"
ACTION_WRITE_CASES = "write_cases"
ACTION_WRITE_INCIDENTS = "write_incidents"
ACTION_PURGE_INCIDENTS = "purge_incidents"

ROLES_PERMITIDOS = (
    ROLE_ADMIN,
    ROLE_SOPORTE,
    ROLE_EXPERIENCIA,
    ROLE_GERENCIAS,
    ROLE_AREA_IA,
)
ROLES_HEREDADOS = (ROLE_VIEWER_LEGACY,)

NOMBRES_ROLES = MappingProxyType(
    {
        ROLE_ADMIN: "Admin",
        ROLE_SOPORTE: "Soporte",
        ROLE_EXPERIENCIA: "Experiencia",
        ROLE_GERENCIAS: "Gerencias",
        ROLE_AREA_IA: "Área de IA",
        ROLE_VIEWER_LEGACY: "Viewer (heredado)",
    }
)

# Identificadores estables de las 17 vistas. No dependen del texto del menú.
VIEW_CARGAR_CASOS = "cargar_casos"
VIEW_CASOS = "casos"
VIEW_DASHBOARD_CASOS_SOPORTE = "dashboard_casos_soporte"
VIEW_CONTROL_DIARIO_SOPORTE = "control_diario_soporte"
VIEW_KPI_CASOS_CLIENTE_EXTERNO = "kpi_casos_cliente_externo"
VIEW_CARGAR_INCIDENTES = "cargar_incidentes"
VIEW_INCIDENTES = "incidentes"
VIEW_DASHBOARD_INCIDENTES = "dashboard_incidentes"
VIEW_KPI_INCIDENTES = "kpi_incidentes"
VIEW_KPI_2025_2026 = "kpi_2025_2026"
VIEW_REINCIDENCIAS_PROBLEMAS = "reincidencias_problemas_sugeridos"
VIEW_SEGUIMIENTO_RPOST = "seguimiento_rpost"
VIEW_SEGUIMIENTO_AUTENTIC = "seguimiento_autentic"
VIEW_SEGUIMIENTO_INCIDENTES = "seguimiento_incidentes"
VIEW_KPI_CLIENTES_CLAVE = "kpi_clientes_clave"
VIEW_CLIENTES_CLAVE = "clientes_clave"
VIEW_ADMINISTRAR_USUARIOS = "administrar_usuarios"

VISTAS = (
    VIEW_CARGAR_CASOS,
    VIEW_CASOS,
    VIEW_DASHBOARD_CASOS_SOPORTE,
    VIEW_CONTROL_DIARIO_SOPORTE,
    VIEW_KPI_CASOS_CLIENTE_EXTERNO,
    VIEW_CARGAR_INCIDENTES,
    VIEW_INCIDENTES,
    VIEW_DASHBOARD_INCIDENTES,
    VIEW_KPI_INCIDENTES,
    VIEW_KPI_2025_2026,
    VIEW_REINCIDENCIAS_PROBLEMAS,
    VIEW_SEGUIMIENTO_RPOST,
    VIEW_SEGUIMIENTO_AUTENTIC,
    VIEW_SEGUIMIENTO_INCIDENTES,
    VIEW_KPI_CLIENTES_CLAVE,
    VIEW_CLIENTES_CLAVE,
    VIEW_ADMINISTRAR_USUARIOS,
)

_PERMISOS_SOPORTE = frozenset(
    {
        VIEW_CASOS,
        VIEW_DASHBOARD_CASOS_SOPORTE,
        VIEW_CONTROL_DIARIO_SOPORTE,
        VIEW_INCIDENTES,
        VIEW_DASHBOARD_INCIDENTES,
        VIEW_SEGUIMIENTO_INCIDENTES,
    }
)
_PERMISOS_EXPERIENCIA = frozenset(
    {
        VIEW_CASOS,
        VIEW_KPI_CASOS_CLIENTE_EXTERNO,
        VIEW_INCIDENTES,
        VIEW_KPI_INCIDENTES,
        VIEW_KPI_2025_2026,
        VIEW_KPI_CLIENTES_CLAVE,
        VIEW_CLIENTES_CLAVE,
    }
)
_PERMISOS_GERENCIAS = frozenset(
    {
        VIEW_KPI_CASOS_CLIENTE_EXTERNO,
        VIEW_KPI_INCIDENTES,
        VIEW_KPI_2025_2026,
        VIEW_KPI_CLIENTES_CLAVE,
        VIEW_CLIENTES_CLAVE,
    }
)
_PERMISOS_AREA_IA = _PERMISOS_EXPERIENCIA

# Conserva las 13 opciones no administrativas que hoy recibe ``viewer``.
# Se eliminará después de migrar todos los usuarios a uno de los cinco roles.
_PERMISOS_VIEWER_LEGACY = frozenset(
    {
        VIEW_CASOS,
        VIEW_DASHBOARD_CASOS_SOPORTE,
        VIEW_CONTROL_DIARIO_SOPORTE,
        VIEW_KPI_CASOS_CLIENTE_EXTERNO,
        VIEW_INCIDENTES,
        VIEW_KPI_INCIDENTES,
        VIEW_KPI_2025_2026,
        VIEW_REINCIDENCIAS_PROBLEMAS,
        VIEW_SEGUIMIENTO_RPOST,
        VIEW_SEGUIMIENTO_AUTENTIC,
        VIEW_SEGUIMIENTO_INCIDENTES,
        VIEW_KPI_CLIENTES_CLAVE,
        VIEW_CLIENTES_CLAVE,
    }
)

PERMISOS_POR_ROL = MappingProxyType(
    {
        ROLE_ADMIN: frozenset(VISTAS),
        ROLE_SOPORTE: _PERMISOS_SOPORTE,
        ROLE_EXPERIENCIA: _PERMISOS_EXPERIENCIA,
        ROLE_GERENCIAS: _PERMISOS_GERENCIAS,
        ROLE_AREA_IA: _PERMISOS_AREA_IA,
        ROLE_VIEWER_LEGACY: _PERMISOS_VIEWER_LEGACY,
    }
)

CAPACIDADES = (
    ACTION_MANAGE_USERS,
    ACTION_WRITE_CASES,
    ACTION_WRITE_INCIDENTS,
    ACTION_PURGE_INCIDENTS,
)

# Las mutaciones disponibles actualmente son administrativas. Se mantiene
# separado de las vistas para no conceder pantallas adicionales por compartir
# una operación interna.
CAPACIDADES_POR_ROL = MappingProxyType(
    {
        rol: (frozenset(CAPACIDADES) if rol == ROLE_ADMIN else frozenset())
        for rol in (*ROLES_PERMITIDOS, *ROLES_HEREDADOS)
    }
)


def normalizar_rol(rol):
    """Normaliza un identificador de rol sin asignar valores por defecto."""
    return str(rol or "").strip().casefold().replace(" ", "_")


def rol_valido(rol, incluir_heredados=True):
    """Indica si el rol es definitivo o, temporalmente, heredado."""
    rol = normalizar_rol(rol)
    roles = ROLES_PERMITIDOS + (ROLES_HEREDADOS if incluir_heredados else ())
    return rol in roles


def obtener_permisos_rol(rol):
    """Devuelve un conjunto inmutable; un rol desconocido no recibe acceso."""
    return PERMISOS_POR_ROL.get(normalizar_rol(rol), frozenset())


def puede_acceder(rol, vista):
    """Comprueba el acceso de un rol a un identificador de vista conocido."""
    return vista in VISTAS and vista in obtener_permisos_rol(rol)


def obtener_vistas_permitidas(rol):
    """Devuelve las vistas autorizadas conservando el orden canónico."""
    permisos = obtener_permisos_rol(rol)
    return tuple(vista for vista in VISTAS if vista in permisos)


def obtener_capacidades_rol(rol):
    """Devuelve las capacidades de acción del rol normalizado."""
    return CAPACIDADES_POR_ROL.get(normalizar_rol(rol), frozenset())


def puede_ejecutar(rol, capacidad):
    """Comprueba una acción sin mezclarla con permisos de navegación."""
    return capacidad in CAPACIDADES and capacidad in obtener_capacidades_rol(rol)
