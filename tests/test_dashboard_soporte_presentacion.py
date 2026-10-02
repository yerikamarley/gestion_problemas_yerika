import unittest

import pandas as pd

import app_ui as ui
from tests.test_resumen_ejecutivo_casos import caso


class DashboardSoporteTest(unittest.TestCase):
    def base(self):
        datos = pd.DataFrame([
            caso("A", causa="Token fisico", creado="2026-09-01 08:00"),
            caso("B", causa="ePass", creado="2026-09-01 09:00"),
            caso("C", causa="Firma digital", creado="2026-09-02 10:00"),
            caso("D", causa="Duplicado", descripcion="token", creado="2026-09-03"),
            caso("E", causa="Token fisico", creado="2026-09-03", asignado="Otro equipo"),
        ])
        datos[ui.TEXT_TIPIFICACION_2] = "4 - Solicitudes"
        base, _ = ui.preparar_kpi_casos_cliente_externo(datos)
        base[ui.TEXT_CREADO_DT_DASHBOARD] = pd.to_datetime(base["creado"], format="mixed")
        return base

    def test_token_exclusivo_y_dias_sin_casos(self):
        base = self.base()
        total = ui.resumen_entradas_diarias_casos(base, inicio="2026-09-01", fin="2026-09-04")
        token = ui.resumen_entradas_diarias_casos(base, True, "2026-09-01", "2026-09-04")
        self.assertEqual([2, 1, 1, 0], total[ui.TEXT_CASOS_2].tolist())
        self.assertEqual([2, 0, 0, 0], token[ui.TEXT_CASOS_2].tolist())
        self.assertEqual(total[ui.TEXT_FECHA].tolist(), token[ui.TEXT_FECHA].tolist())

    def test_periodo_sin_token_mantiene_ceros(self):
        base = self.base().iloc[2:]
        token = ui.resumen_entradas_diarias_casos(base, True, "2026-09-01", "2026-09-04")
        self.assertEqual([0, 0, 0, 0], token[ui.TEXT_CASOS_2].tolist())

    def test_dashboard_renderiza_dos_graficas_compactas(self):
        from streamlit.testing.v1 import AppTest
        app = AppTest.from_string('''
from unittest.mock import patch
import app_ui as ui
from tests.test_dashboard_soporte_presentacion import DashboardSoporteTest
base = DashboardSoporteTest().base()
with patch.object(ui, "selector_periodo_sql", return_value=(2026, 9, "Septiembre 2026")), \\
     patch.object(ui, "cargar_casos_soporte_filtrados_cache", return_value=base), \\
     patch.object(ui, "render_distribucion_productos_soporte"), \\
     patch.object(ui, "render_carga_agentes"), \\
     patch.object(ui, "render_seguimiento_casos"):
    ui.dashboard_casos()
''').run()
        self.assertFalse(app.exception)
        self.assertEqual(2, len(app.get("plotly_chart")))
        self.assertTrue(any("Token físico / ePass" in item.value for item in app.markdown))


if __name__ == "__main__":
    unittest.main()
