from pathlib import Path

import pandas as pd

from services.incident_risk_service import build_analysis, classify_incidents
from utils.risk_excel_export import build_risk_workbook


incidents = pd.DataFrame([
    {"numero": "INC-DEMO-01", "creado": "2026-01-10", "servicio_negocio": "EC2", "descripcion": "Error 504 integración consulta de listas"},
    {"numero": "INC-DEMO-02", "creado": "2026-02-11", "servicio_negocio": "Validación de identidad", "descripcion": "Falla OTP validación de identidad"},
    {"numero": "INC-DEMO-03", "creado": "2026-03-12", "servicio_negocio": "Certificado digital", "descripcion": "Error de emisión de certificado digital"},
    {"numero": "INC-DEMO-04", "creado": "2026-04-13", "servicio_negocio": "Facturación", "descripcion": "Inconsistencia de facturación"},
    {"numero": "INC-DEMO-05", "creado": "2026-05-14", "servicio_negocio": "ERP", "descripcion": "Caso de soporte bloqueado"},
    {"numero": "INC-DEMO-06", "creado": "2026-06-15", "servicio_negocio": "OCSP", "descripcion": "Caída OCSP no responde"},
    {"numero": "INC-DEMO-07", "creado": "2026-07-16", "servicio_negocio": "AWS S3", "descripcion": "Pérdida de información por falla almacenamiento bucket"},
    {"numero": "INC-DEMO-08", "creado": "2026-08-17", "descripcion": "Novedad pendiente de análisis"},
    {"numero": "INC-DEMO-09", "creado": "2026-08-18", "descripcion": "Olvido de contraseña"},
])

analysis = build_analysis(classify_incidents(incidents), range(1, 13))
output = Path("outputs/matriz_riesgos_ajustada/Matriz_Riesgos_Materializados_2026.xlsx")
output.parent.mkdir(parents=True, exist_ok=True)
output.write_bytes(build_risk_workbook(analysis, 2026, 1, 12))
print(output.resolve())
