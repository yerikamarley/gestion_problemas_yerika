"""Catálogo corporativo y prioridades para riesgos materializados."""

EXCLUSION_CATEGORIES = {
    "Falsas Alarmas (Monitoreo NOC/SOC)": {
        "description": "Alertas normalizadas o sin indisponibilidad/afectación material comprobada.",
        "priority": 10,
    },
    "Requerimientos y Soporte Rutinario (BAU)": {
        "description": "Consultas, acompañamientos y actividades ordinarias sin materialización de riesgo.",
        "priority": 20,
    },
    "Pruebas y Simulacros": {
        "description": "Pruebas controladas o tickets generados intencionalmente para validar servicios.",
        "priority": 5,
    },
    "Eventos de Seguridad Contenidos": {
        "description": "Eventos detectados y contenidos sin pérdida de información ni indisponibilidad material.",
        "priority": 15,
    },
}


def _risk(risk_id, name, failure, owner, impact, impact_pct, raci, rules, exclusions=(), priority=100):
    return {
        "risk_id": risk_id,
        "name": name,
        "description": name,
        "associated_failure": failure,
        "owner": owner,
        "impact": impact,
        "impact_percentage": impact_pct,
        "responsible_r": raci[0],
        "accountable_a": raci[1],
        "consulted_c": raci[2],
        "informed_i": raci[3],
        "classification_rules": rules,
        "exclusions": tuple(exclusions),
        "active": True,
        "priority": priority,
    }


# Una regla exige evidencia de al menos dos grupos, salvo que ``strong`` coincida.
RISK_CATALOG = (
    _risk("R-MON", "Posibilidad de fallas en la detección, generación o atención oportuna de alertas tecnológicas.",
          "Fallas del monitoreo para detectar, generar o escalar alertas ante afectaciones reales.",
          "DIR. DE INFRAESTRUCTURA DE TI", "MODERADO (60%)", 0.60,
          ("NOC / Monitoreo", "Dir. Infra. TI", "Ciberseguridad", "Soporte TI"),
          {"product": ("monitoreo", "noc", "zabbix", "grafana", "prometheus", "solarwinds"),
           "failure": ("no genero alerta", "no generó alerta", "sin alerta", "alerta tardia", "alerta tardía", "no detecto", "no detectó"),
           "strong": ("monitoreo no genero alerta", "noc no detecto", "alerta no generada")}, priority=5),
    _risk("R164", "Posibilidad de fallas en el despliegue, la configuración o la integración en entornos de nube, on-premise o híbridos.",
          "Caída de máquinas virtuales/EC2, despliegues, configuración, APIs e integraciones técnicas.",
          "DIR. DE INFRAESTRUCTURA DE TI", "MENOR (40%)", 0.40,
          ("Analista Ciberseg.", "Dir. Infra. TI", "Gestor Capacidad", "Dir. Registro y Uso"),
          {"technical": ("ec2", "maquina virtual", "maquinas virtuales", "despliegue", "configuracion", "api", "integracion", "consulta de listas"),
           "failure": ("error 500", "error 503", "error 504", "caida", "indisponibilidad", "perdida de integracion", "no responde"),
           "strong": ("ec2 error 504", "error 504 integracion", "falla de despliegue", "caida ec2")}, priority=10),
    _risk("R76", "Posibilidad de inadecuación o debilidades en los procesos de validación de identidad.",
          "Problemas de biometría, VDI, correos de validación, OTP y SMS.", "DIR. DE REGISTRO Y USO",
          "MODERADO (60%)", 0.60, ("Soporte Nivel 2 / Oficial Revisión", "Dir. Registro y Uso", "Dir. Registro y Uso", "Cliente"),
          {"product": ("biometria", "biometrico", "vdi", "validacion de identidad", "autentic", "authentic"),
           "failure": ("otp", "sms", "correo de validacion", "codigo de validacion", "no recibe codigo", "no llega codigo", "falla de validacion"),
           "strong": ("validacion biometrica", "validacion de identidad")}, priority=20),
    _risk("R75", "Posibilidad de fallas o debilidades en los procesos de revisión, aprobación y emisión de certificados digitales.",
          "Errores de emisión en portal y fallas en descarga de Tokens.", "DIR. DE REGISTRO Y USO",
          "MODERADO (60%)", 0.60, ("Soporte Nivel 2", "Dir. Registro y Uso", "Desarrollo", "Negocios"),
          {"product": ("certificado digital", "certitoken", "token virtual", "token fisico", "crl", "pki"),
           "failure": ("emision", "emitir", "descarga", "no descarga", "error al firmar", "no permite firmar", "validar firma"),
           "strong": ("error de emision", "falla descarga token", "certitoken error al firmar")}, priority=30),
    _risk("R113", "Posibilidad de errores en información, omisión de productos o servicios y/o errores de cálculo.",
          "Errores Pegaso/Dynamics, facturación, código de barras e inconsistencias de datos.", "DIR. DE FACTURACIÓN",
          "MAYOR (80%)", 0.80, ("Desarrollo / Soporte N2", "Dir. de Facturación", "Desarrollo", "Negocios"),
          {"product": ("pegaso", "dynamics", "facturacion", "codigo de barras"),
           "failure": ("inconsistencia", "error", "omision", "calculo", "integracion"),
           "strong": ("problema de facturacion", "inconsistencia de facturacion")}, priority=40),
    _risk("R132", "Posibilidad de que no se gestionen o atiendan con soluciones de fondo los casos de soporte radicados.",
          "Casos bloqueados, errores ERP, ventas masivas y fallas operativas de soporte.", "DIR. DE IMPLEMENTACIÓN Y SERVICIO",
          "MODERADO (60%)", 0.60, ("Soporte Nivel 2", "Dir. Impl y Serv", "Infraestructura", "Negocios"),
          {"product": ("erp", "portal de ventas masivas", "soporte"),
           "failure": ("bloqueado", "sin solucion", "sin solución", "falla operativa", "no gestionado", "reincidente"),
           "strong": ("caso de soporte bloqueado", "error erp")}, priority=50),
    _risk("R140", "Posibilidad de interrupción del funcionamiento de la infraestructura tecnológica.",
          "Caídas de infraestructura core, portales, RPOST, OCSP y MongoDB.", "DIR. DE INFRAESTRUCTURA DE TI",
          "MAYOR (80%)", 0.80, ("NOC / Soporte TI", "Dir. Infra. TI", "Ciberseguridad", "Dir. Registro y Uso"),
          {"product": ("infraestructura core", "ocsp", "osps", "rpost", "autentic", "ssps", "ssp", "epss", "certitoken", "token virtual", "token fisico", "mongodb", "portal"),
           "failure": ("caida", "caido", "indisponibilidad", "no responde", "alarma real", "interrupcion"),
           "strong": ("caida ocsp", "ocsp no responde", "indisponibilidad ocsp")},
          exclusions=("ec2", "maquina virtual", "despliegue", "integracion", "api", "error 500", "error 503", "error 504"), priority=60),
    _risk("R47", "Posibilidad de pérdida parcial o total de información.",
          "Falla en almacenamiento de documentos en bucket AWS.", "DIR. DE INFRAESTRUCTURA DE TI",
          "MAYOR (80%)", 0.80, ("Soporte Nivel 2", "Dir. Infra. TI", "Ciberseguridad", "Dirección de Registro y Uso"),
          {"product": ("bucket", "s3", "aws", "almacenamiento"),
           "failure": ("perdida de informacion", "documento no almacenado", "no guarda documento", "falla"),
           "strong": ("perdida de informacion", "falla almacenamiento bucket")}, priority=70),
)

RISK_BY_ID = {risk["risk_id"]: risk for risk in RISK_CATALOG}
REINCIDENCE_THRESHOLD = 2


def recurrence_status(count):
    return "🔥 REINCIDENTE" if int(count) >= REINCIDENCE_THRESHOLD else "AISLADO"
