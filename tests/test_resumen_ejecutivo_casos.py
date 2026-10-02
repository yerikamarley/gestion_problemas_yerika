from datetime import date
import unittest
from unittest.mock import patch

import pandas as pd

from components.resumen_ejecutivo_casos import lamina_resumen_casos
from services.resumen_ejecutivo_casos import construir_resumen_ejecutivo, periodos_corte, causas_para_lamina


def causa(row):
    return row.get("causa", "")


def caso(numero="CS1", **kwargs):
    return {"numero": numero, "asignado": "Yader Neira", "estado": "Nuevo",
            "creado": "2026-09-01 08:00:00", "cerrado": "",
            "fecha_vencimiento_sla": "2026-09-04 08:00:00", "causa": "Firma digital", **kwargs}


class ResumenEjecutivoTest(unittest.TestCase):
    def test_cortes_equivalentes_incluyen_bisiesto_y_cambio_anio(self):
        for corte, anterior in [
            (date(2026, 9, 15), date(2026, 8, 15)),
            (date(2026, 3, 31), date(2026, 2, 28)),
            (date(2024, 3, 30), date(2024, 2, 29)),
            (date(2026, 5, 31), date(2026, 4, 30)),
            (date(2026, 1, 15), date(2025, 12, 15)),
        ]:
            with self.subTest(corte=corte):
                periodos = periodos_corte(corte)
                self.assertEqual((anterior.replace(day=1), anterior), periodos["anterior"])
                self.assertEqual((corte.replace(day=1), corte), periodos["actual"])

    def test_soporte_exclusivo_corte_completo_y_duplicados(self):
        df = pd.DataFrame([
            caso("SI1", creado="2026-09-15 23:59:59", asignado="PAULA PÁEZ"),
            caso("SI2"), caso("SI2"),
            caso("NO1", asignado=""), caso("NO2", asignado="Otro equipo"),
            caso("NO3", creado="2026-09-16 00:00:00"),
            caso("NO4", creado="2026-08-31 23:59:59"),
            caso("NO5", creado="fecha inválida"),
            caso("NO6", actualizado="2026-09-01"),
            caso("NO6", asignado="Otro equipo", actualizado="2026-09-02"),
            caso(""),
        ])
        reporte = construir_resumen_ejecutivo(pd.DataFrame(), df, date(2026, 9, 15), causa)
        self.assertEqual({"SI1", "SI2", "NO1"}, set(reporte["bases"]["actual"]["numero"]))
        self.assertEqual(3, reporte["metricas"]["actual"]["total"])
        self.assertEqual(3, reporte["metricas"]["actual"]["abiertos"])
        self.assertEqual(3, reporte["causas"]["Casos"].sum())
        self.assertEqual(3, reporte["diferencia"])
        self.assertIsNone(reporte["diferencia_sla"])

    def test_sla_denominador_y_diferencias_verificables(self):
        anterior = pd.DataFrame([caso("A", creado="2026-08-01", cerrado="2026-08-02", estado="Cerrado", fecha_vencimiento_sla="2026-08-03")])
        actual = pd.DataFrame([
            caso("B", estado="Cerrado", cerrado="2026-09-04 08:00:00"),
            caso("C", estado="Cerrado", cerrado="2026-09-04 08:00:01"),
            caso("D", estado="Cerrado", cerrado="2026-09-02", fecha_vencimiento_sla=""),
            caso("E"),
        ])
        reporte = construir_resumen_ejecutivo(anterior, actual, date(2026, 9, 15), causa)
        self.assertEqual(100, reporte["metricas"]["anterior"]["sla"])
        self.assertEqual(50, reporte["metricas"]["actual"]["sla"])
        self.assertEqual(2, reporte["metricas"]["actual"]["evaluados"])
        self.assertEqual(1, reporte["metricas"]["actual"]["sin_evaluar"])
        self.assertEqual(-50, reporte["diferencia_sla"])
        self.assertEqual(3, reporte["diferencia"])
        self.assertEqual(300, reporte["variacion"])

    def test_causas_reconcilian_y_no_duplican_casos(self):
        df = pd.DataFrame([caso(str(i), causa=valor) for i, valor in enumerate(["Firma", "Firma", "Token", "Pago", "", "Agenda"])])
        reporte = construir_resumen_ejecutivo(pd.DataFrame(), df, date(2026, 9, 15), causa)
        self.assertEqual(6, reporte["causas"]["Casos"].sum())
        self.assertEqual(6, reporte["causas_lamina"]["Casos"].sum())
        self.assertAlmostEqual(100, reporte["causas_lamina"]["Porcentaje"].sum())
        self.assertEqual(3, len(reporte["causas_lamina"]))
        self.assertIn("Solicitud operativa", reporte["causas"]["Causa"].tolist())
        self.assertTrue(reporte["causas_lamina"]["Porcentaje"].is_monotonic_decreasing)

    def test_causa_raiz_prevalece_sobre_producto_en_descripcion(self):
        from app_ui import inferir_causa_comun_caso
        for raiz, categoria in [
            ("Caso duplicado", "Duplicado"),
            ("Captores", "Captores"),
            ("Token fisico", "Token fisico / ePass"),
            ("Firma digital", "Firma digital"),
        ]:
            with self.subTest(raiz=raiz):
                self.assertEqual(categoria, inferir_causa_comun_caso(pd.Series({
                    "causa": raiz, "descripcion": "Solicitud de firma digital con token",
                })))

    def test_causas_coinciden_con_dashboard_soporte(self):
        from app_ui import preparar_kpi_casos_cliente_externo, inferir_causa_comun_caso, resumen_principales_causas_servicios_casos
        df = pd.DataFrame([
            caso("A", causa="Firma"), caso("B", causa="Token", asignado=""),
            caso("C", causa="", asignado="Yader Neira"),
            caso("D", causa="Pago", asignado="Otro equipo"),
        ])
        base, metricas = preparar_kpi_casos_cliente_externo(df)
        reporte = construir_resumen_ejecutivo(pd.DataFrame(), df, date(2026, 9, 15), inferir_causa_comun_caso)
        self.assertEqual(metricas["total"], reporte["metricas"]["actual"]["total"])
        self.assertEqual(base["causa_comun"].value_counts().to_dict(),
                         reporte["causas"].set_index("Causa")["Casos"].to_dict())
        tabla, _ = resumen_principales_causas_servicios_casos(base)
        pd.testing.assert_frame_equal(reporte["causas_lamina"], tabla.rename(columns={
            "Causa raíz": "Causa", "Cantidad": "Casos", "% casos": "Porcentaje",
        }))

    def test_tabla_kpi_dos_principales_y_acumulado_maximo_cuatro_nombres(self):
        from app_ui import preparar_kpi_casos_cliente_externo, resumen_principales_causas_servicios_casos
        raices = ["duplicado", "captores", "solicitud operativa", "token", "firma", "activacion"]
        base, _ = preparar_kpi_casos_cliente_externo(pd.DataFrame([
            caso(str(i), causa=raiz) for i, raiz in enumerate(raices)
        ]))
        tabla, _ = resumen_principales_causas_servicios_casos(base)
        self.assertEqual(3, len(tabla))
        self.assertEqual(6, tabla["Cantidad"].sum())
        self.assertAlmostEqual(100, tabla["% casos"].sum())
        etiqueta = tabla.iloc[0]["Causa raíz"]
        self.assertEqual(4, etiqueta.count("(1 casos)"))
        self.assertNotIn("+", etiqueta)
        self.assertNotIn("\n", etiqueta)
        self.assertEqual(4, tabla.iloc[0]["Cantidad"])
        self.assertTrue(tabla["% casos"].is_monotonic_decreasing)

    def test_porcentaje_descendente_incluye_acumulado(self):
        for cantidades, esperado in [([40, 30, 20, 10], [40, 30, 30]),
                                     ([30, 25, 24, 21], [45, 30, 25]),
                                     ([20, 50, 30], [50, 30, 20]),
                                     ([25, 75], [75, 25])]:
            causas = pd.DataFrame({"Causa": list("ABCD")[:len(cantidades)],
                                   "Casos": cantidades, "Porcentaje": cantidades})
            self.assertEqual(esperado, causas_para_lamina(causas)["Porcentaje"].tolist())

    def test_no_usa_etiqueta_sin_causa_identificada(self):
        from app_ui import inferir_causa_comun_caso
        self.assertEqual("Solicitud operativa", inferir_causa_comun_caso(pd.Series(dtype=object)))

    def test_catalogo_limitado_a_captura_mas_duplicado_y_captores(self):
        from app_ui import CASE_COMMON_CAUSE_RULES, inferir_causa_comun_caso
        permitidas = {"Firma digital", "Token fisico / ePass",
                      "Activacion o descarga de certificado", "Solicitud operativa",
                      "Duplicado", "Captores"}
        self.assertEqual(permitidas, {nombre for nombre, _ in CASE_COMMON_CAUSE_RULES})
        for anterior in ["Sin causa comun", "Sin causa identificada", "Pendiente de revisión",
                         "Falla técnica", "Instalacion o configuracion", "Plataforma externa",
                         "Enviado a otros canales", "Solicitud básica de soporte"]:
            self.assertIn(inferir_causa_comun_caso(pd.Series({"causa": anterior})), permitidas)

    def test_sin_datos_y_sin_soporte_no_inventan_sla(self):
        for df in [pd.DataFrame(), pd.DataFrame([caso(asignado="Otro")])]:
            reporte = construir_resumen_ejecutivo(df, df, date(2026, 9, 15), causa)
            self.assertEqual(0, reporte["metricas"]["actual"]["total"])
            self.assertIsNone(reporte["metricas"]["actual"]["sla"])
            html = lamina_resumen_casos(reporte)
            self.assertIn("Sin dato", html)
            self.assertIn("Sin casos para agrupar", html)

    def test_textos_no_inyectan_html_y_se_mantiene_formato(self):
        df = pd.DataFrame([caso(causa='<script>alert(1)</script>')])
        reporte = construir_resumen_ejecutivo(df, df, date(2026, 9, 15), causa)
        html = lamina_resumen_casos(reporte, foco='<img src=x>', decision='<script>acción</script>')
        self.assertNotIn('<script>', html)
        self.assertIn('&lt;script&gt;', html)
        for texto in ["DECISIÓN REQUERIDA", "BACKLOG DEL PERÍODO", "Casos por mes", "Causas agrupadas", "Impacto al cliente"]:
            self.assertIn(texto, html)

    def test_interfaz_consulta_solo_los_dos_periodos_y_muestra_soporte(self):
        import app_ui
        from streamlit.testing.v1 import AppTest
        app = AppTest.from_string('''
from unittest.mock import patch
import pandas as pd
import streamlit as st
import app_ui as ui
from tests.test_resumen_ejecutivo_casos import caso
def cargar(**kwargs):
    st.session_state.setdefault("consultas", []).append(kwargs)
    return pd.DataFrame([caso(), caso("OTRO", asignado="Otro"), caso("VACIO", asignado="")])
with patch.object(ui, "cargar_casos_soporte_filtrados_cache", side_effect=cargar):
    ui.dashboard_kpi_casos_cliente_externo()
''').run()
        self.assertFalse(app.exception)
        self.assertEqual("Resumen ejecutivo", app.radio(key="kpi_casos_cliente_externo_vista").value)
        self.assertEqual(2, len(app.session_state["consultas"]))
        app.date_input(key="corte_resumen_ejecutivo_casos").set_value(date(2026, 9, 15)).run()
        self.assertFalse(app.exception)
        self.assertEqual([
            {"fecha_inicio": date(2026, 8, 1), "fecha_fin": date(2026, 8, 15)},
            {"fecha_inicio": date(2026, 9, 1), "fecha_fin": date(2026, 9, 15)},
        ], app.session_state["consultas"][-2:])
        self.assertEqual(2, app.dataframe[0].value.iloc[1]["total"])


if __name__ == "__main__":
    unittest.main()
