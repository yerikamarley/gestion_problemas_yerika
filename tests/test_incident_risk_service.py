import unittest

import pandas as pd
from openpyxl import load_workbook

from services.incident_risk_service import build_analysis, classify_incident, classify_incidents
from utils.risk_excel_export import build_risk_workbook


class IncidentRiskServiceTest(unittest.TestCase):
    def classify(self, **values):
        values.setdefault("numero", "INC0001")
        return classify_incident(values)

    def test_r76(self):
        result = self.classify(servicio_negocio="Biometría", descripcion="Falla de validación de identidad con código OTP")
        self.assertEqual(("RISK", "R76"), (result["classification_type"], result["risk_id"]))

    def test_false_alarm(self):
        result = self.classify(descripcion="Alerta tipificada como falsa alarma sin afectación")
        self.assertEqual(("EXCLUSION", "Falsas Alarmas (Monitoreo NOC/SOC)"), (result["classification_type"], result["exclusion_category"]))

    def test_ec2_504_integration_is_r164_not_r140(self):
        result = self.classify(servicio_negocio="EC2", descripcion="Error 504 integración consulta de listas")
        self.assertEqual("R164", result["risk_id"]); self.assertNotIn("R140", result["risk_id"])

    def test_real_ocsp_is_r140(self):
        result = self.classify(servicio_negocio="OCSP", descripcion="Caída real, OCSP no responde")
        self.assertEqual("R140", result["risk_id"])

    def test_password_is_bau(self):
        result = self.classify(descripcion="Olvido de contraseña")
        self.assertEqual(("EXCLUSION", "Requerimientos y Soporte Rutinario (BAU)"), (result["classification_type"], result["exclusion_category"]))

    def test_ambiguous_is_pending(self):
        self.assertEqual("PENDING", self.classify(descripcion="Se reporta una novedad")["classification_type"])

    def test_overlap_has_single_priority_result(self):
        result = self.classify(servicio_negocio="OCSP EC2", descripcion="Error 504 integración, caída OCSP")
        self.assertEqual("R164", result["risk_id"])

    def test_reconciliation_counts_unique_incidents(self):
        df = pd.DataFrame([
            {"numero":"INC1", "creado":"2026-01-01", "servicio_negocio":"Biometría", "descripcion":"Falla OTP validación de identidad"},
            {"numero":"INC2", "creado":"2026-01-02", "descripcion":"Olvido de contraseña"},
            {"numero":"INC3", "creado":"2026-01-03", "descripcion":"Novedad no concluyente"},
        ])
        analysis = build_analysis(classify_incidents(df), [1])
        v = analysis["validation"]
        self.assertTrue(v["reconciled"]); self.assertEqual(v["total"], v["risk"] + v["exclusion"] + v["pending"])
        self.assertTrue(all(analysis["risks"]["Cantidad tickets asociados"] == analysis["risks"]["Tickets Asociados"].str.split(", ").str.len()))

    def test_excel_has_required_sheets(self):
        df = pd.DataFrame([{"numero":"INC1", "creado":"2026-01-01", "descripcion":"Novedad"}])
        analysis = build_analysis(classify_incidents(df), [1])
        from io import BytesIO
        wb = load_workbook(BytesIO(build_risk_workbook(analysis, 2026, 1, 1)))
        self.assertEqual({"Resumen","Riesgos","Conciliación","Exclusiones","Detalle","Pendientes"}, set(wb.sheetnames))


if __name__ == "__main__": unittest.main()
