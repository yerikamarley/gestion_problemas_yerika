from pathlib import Path

import pandas as pd

from app_logic import construir_matriz_riesgo_incidentes
from services.incident_risk_service import build_analysis, classify_incidents
from utils.risk_excel_export import build_risk_workbook


incidents = pd.DataFrame([
    {"numero": "INC-DEMO-01", "creado": "2026-01-10", "servicio_negocio": "EC2", "descripcion": "Error 504 integración consulta de listas"},
    {"numero": "INC-DEMO-02", "creado": "2026-02-11", "servicio_negocio": "Validación de identidad", "descripcion": "Falla OTP validación de identidad"},
    {"numero": "INC-DEMO-03", "creado": "2026-03-12", "servicio_negocio": "Certificado digital", "descripcion": "Error de emisión de certificado digital"},
    {"numero": "INC-DEMO-04", "creado": "2026-04-13", "servicio_negocio": "Facturación", "descripcion": "Inconsistencia de facturación"},
    {"numero": "INC-DEMO-05", "creado": "2026-05-14", "servicio_negocio": "ERP", "descripcion": "Caso de soporte bloqueado"},
    {"numero": "INC-DEMO-06", "creado": "2026-06-15", "servicio_negocio": "OCSP", "descripcion": "Caída OCSP no responde", "tipificacion_auto": "Indisponibilidad", "causa_raiz_auto": "Agotamiento de conexiones"},
    {"numero": "INC-DEMO-10", "creado": "2026-06-18", "servicio_negocio": "OCSP", "descripcion": "Nueva caída OCSP no responde", "tipificacion_auto": "Indisponibilidad", "causa_raiz_auto": "Agotamiento de conexiones"},
    {"numero": "INC-DEMO-11", "creado": "2026-06-18", "servicio_negocio": "NOC", "descripcion": "El monitoreo no generó alerta durante la caída", "tipificacion_auto": "Falla de alertamiento", "causa_raiz_auto": "Regla de monitoreo incompleta"},
    {"numero": "INC-DEMO-12", "creado": "2026-06-20", "servicio_negocio": "Certitoken", "descripcion": "Solicitud de instalación y activación en nuevo PC"},
    {"numero": "INC0017514", "creado": "2026-08-11", "servicio_negocio": "Certificación Digital", "descripcion": "La URL responde HTTP 200, pero el certificado SSL se encuentra vencido y requiere renovación.", "causa_raiz_auto": "Certificado SSL vencido o renovación no oportuna"},
    {"numero": "INC0017551", "creado": "2026-08-19", "servicio_negocio": "Servicios de Infraestructura", "breve_descripcion": "Alarmas sobre infraestructura de PKI y monitoreo OCSP - TSA", "descripcion": "Down repositorio LDAP PKI. Down OCSP. Down SSPS. Down token virtual. Unreachable HSM.", "tipificacion_auto": "Indisponibilidad múltiple de infraestructura PKI"},
    {"numero": "INC-DEMO-07", "creado": "2026-07-16", "servicio_negocio": "AWS S3", "descripcion": "Pérdida de información por falla almacenamiento bucket"},
    {"numero": "INC-DEMO-08", "creado": "2026-08-17", "descripcion": "Novedad pendiente de análisis"},
    {"numero": "INC-DEMO-09", "creado": "2026-08-18", "descripcion": "Olvido de contraseña"},
])

analysis = build_analysis(classify_incidents(incidents), range(1, 13))
analysis["patterns"] = construir_matriz_riesgo_incidentes(incidents)
output = Path("outputs/matriz_riesgos_ajustada/Matriz_Riesgos_Materializados_2026.xlsx")
output.parent.mkdir(parents=True, exist_ok=True)
output.write_bytes(build_risk_workbook(analysis, 2026, 1, 12))
print(output.resolve())
