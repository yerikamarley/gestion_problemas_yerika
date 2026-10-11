from unittest import TestCase

import pandas as pd

from services.casos_historial import (
    analisis_token_fisico, COL_CAUSA, COL_CLASIFICACION, COL_FISICO,
    COL_EVIDENCIA_CAUSA, COL_EVIDENCIA_MOTIVO, COL_EVIDENCIA_TOKEN,
    COL_MOTIVO_INICIAL, COL_DESTINO, COL_TOKEN, ENVIADO_OTROS_CANALES,
)
from tests.test_casos_historial import evento, DERIVACION_AGENDA, ACUSE
from tests.test_canales_documentados import COMERCIAL
from tests.test_resumen_ejecutivo_casos import caso


class AnalisisTokenFisicoTest(TestCase):
    def datos(self):
        return pd.DataFrame([
            caso("INST1", causa="", descripcion="Necesito instalar mi token físico",
                 historial_conversacion=[evento(DERIVACION_AGENDA, "Correo Saliente")]),
            caso("INST2", causa="Driver incompatible", descripcion="Necesito instalar mi token físico"),
            caso("DESC", causa="Enlace de descarga vencido", descripcion="No puedo descargar mi token físico",
                 historial_conversacion=[evento(COMERCIAL, "Correo Saliente")]),
            caso("PIN", causa="", descripcion="Mi token físico tiene el PIN bloqueado",
                 historial_conversacion=[evento("Causa raíz: PIN bloqueado por intentos fallidos.")]),
            caso("AMBOS", causa="", descripcion="Necesito instalar un token físico y un token virtual"),
            caso("VIRTUAL", causa="", descripcion="Necesito instalar un token virtual"),
            caso("USB", causa="", descripcion="Necesito instalar ePass2003 USB"),
            caso("ACUSE", causa="", descripcion="Consulta sobre factura",
                 historial_conversacion=[evento(ACUSE, "Correo Saliente", "system")]),
            caso("CATALOGO", causa="", descripcion="Ayuda con token virtual",
                 historial_conversacion=[evento(DERIVACION_AGENDA, "Correo Saliente")]),
            caso("NEGADO", causa="", descripcion="No utiliza token físico, tiene un token virtual"),
            caso("ACTUALIZADO", causa="", descripcion="Instalar token físico"),
            caso("ACTUALIZADO", causa="", descripcion="Instalar token virtual"),
            caso("INST2", causa="Driver incompatible", descripcion="Necesito instalar mi token físico"),
        ])

    def test_cohorte_explicita_deduplicada_y_mismos_casos_de_grafica(self):
        import app_ui as ui
        datos = self.datos()
        resultado = analisis_token_fisico(datos)
        self.assertEqual({"INST1", "INST2", "DESC", "PIN", "AMBOS"}, set(resultado["casos"]["numero"]))
        self.assertEqual({"total": 5, "con_causa": 3, "sin_causa": 2, "enviados": 2}, resultado["metricas"])
        self.assertEqual(5, resultado["motivos"]["Casos"].sum())
        self.assertEqual(100, resultado["motivos"]["% token físico"].sum())
        datos[ui.TEXT_CREADO_DT_DASHBOARD] = pd.to_datetime(datos["creado"])
        grafica = ui.resumen_entradas_diarias_casos(datos, solo_token=True)
        self.assertEqual(grafica[ui.TEXT_CASOS_2].sum(), resultado["metricas"]["total"])

    def test_motivo_inicial_no_se_reemplaza_por_derivacion(self):
        resultado = analisis_token_fisico(self.datos())
        motivos = resultado["motivos"].set_index("Motivo de contacto")
        self.assertEqual(3, motivos.loc["Instalación o configuración", "Casos"])
        self.assertEqual(1, motivos.loc["Instalación o configuración", "Enviados a otros canales"])
        self.assertEqual(1, motivos.loc["Falla de descarga o activación", "Enviados a otros canales"])
        self.assertNotIn(ENVIADO_OTROS_CANALES, motivos.index)
        base = resultado["casos"].set_index("numero")
        self.assertEqual(ENVIADO_OTROS_CANALES, base.loc["INST1", COL_CLASIFICACION])
        self.assertEqual("Agenda", base.loc["INST1", COL_DESTINO])
        self.assertEqual("Instalación o configuración", base.loc["INST1", COL_MOTIVO_INICIAL])

    def test_causas_reales_evidencia_y_denominador_documentados(self):
        resultado = analisis_token_fisico(self.datos())
        causas = resultado["causas"]
        self.assertEqual(3, causas["Casos"].sum())
        self.assertEqual({"Driver incompatible", "Enlace de descarga vencido", "pin bloqueado por intentos fallidos"},
                         set(causas["Causa documentada"]))
        self.assertTrue(causas["% documentados"].eq(33.33).all())
        pin = resultado["casos"].set_index("numero").loc["PIN"]
        self.assertIn("Causa raíz:", pin[COL_EVIDENCIA_CAUSA])
        self.assertIn("token físico", pin[COL_EVIDENCIA_MOTIVO])
        self.assertIn("token físico", pin[COL_EVIDENCIA_TOKEN])
        ambos = resultado["casos"].set_index("numero").loc["AMBOS"]
        self.assertEqual("Físico y virtual", ambos[COL_TOKEN])
        self.assertEqual("Sin causa documentada", ambos[COL_CAUSA])

    def test_vacios_y_sin_token_fisico(self):
        for datos in [pd.DataFrame(), pd.DataFrame([caso("V", descripcion="Instalar token virtual")])]:
            resultado = analisis_token_fisico(datos)
            self.assertEqual(0, resultado["metricas"]["total"])
            self.assertTrue(resultado["motivos"].empty)
            self.assertTrue(resultado["causas"].empty)
            self.assertTrue(resultado["casos"].empty)

    def test_render_navega_motivos_causas_y_evidencia_sin_consultas(self):
        from streamlit.testing.v1 import AppTest
        app = AppTest.from_string('''
from unittest.mock import patch
import app_ui as ui
from tests.test_analisis_token_fisico import AnalisisTokenFisicoTest
with patch("services.servicenow.ServiceNowClient", side_effect=AssertionError("Sin nueva consulta")), \\
     patch.object(ui, "cargar_evidencia_casos_cache", side_effect=AssertionError("Sin nueva lectura")):
    ui.render_analisis_token_fisico(AnalisisTokenFisicoTest().datos(), "Septiembre 2026")
''').run()
        self.assertFalse(app.exception)
        self.assertEqual(0, len(app.get("plotly_chart")))
        self.assertEqual(5, app.dataframe[0].value["Casos"].sum())
        app.radio(key="vista_analisis_token_fisico").set_value("Causas documentadas").run()
        self.assertFalse(app.exception)
        self.assertEqual(3, app.dataframe[0].value["Casos"].sum())
        app.radio(key="vista_analisis_token_fisico").set_value("Casos y evidencia").run()
        self.assertFalse(app.exception)
        self.assertEqual(5, len(app.dataframe[0].value))
        app.selectbox(key="motivo_token_fisico").set_value("Instalación o configuración").run()
        self.assertFalse(app.exception)
        self.assertEqual({"INST1", "INST2", "AMBOS"}, set(app.dataframe[0].value["Número"]))
        self.assertIn(COL_EVIDENCIA_CAUSA, app.dataframe[0].value.columns)

    def test_sin_diagnostico_muestra_motivos_y_ausencia_de_causas(self):
        from streamlit.testing.v1 import AppTest
        app = AppTest.from_string('''
import pandas as pd
import app_ui as ui
ui.render_analisis_token_fisico(pd.DataFrame([
    {"numero": "A", "descripcion": "Instalar token f\\u00edsico"}]), "Octubre 2026")
''').run()
        self.assertFalse(app.exception)
        app.radio(key="vista_analisis_token_fisico").set_value("Causas documentadas").run()
        self.assertFalse(app.exception)
        self.assertTrue(any("todavía no hay causas técnicas documentadas" in aviso.value for aviso in app.info))
