from datetime import date, timedelta
import unittest
from unittest.mock import patch

import pandas as pd
import app_logic
import app_ui
from repositories.tables import agregar_condicion_fechas


class PeriodoCasosTest(unittest.TestCase):
    def test_sql_incluye_fin_completo_y_cruce_de_mes(self):
        where, params = [], []
        agregar_condicion_fechas(where, params, date(2026, 8, 31), date(2026, 9, 1))
        self.assertEqual(["(creado >= ? AND creado < ?)"], where)
        self.assertEqual(["2026-08-31", "2026-09-02"], params)

    def test_rechaza_rango_incompleto_o_invertido(self):
        for inicio, fin in [(date(2026, 9, 1), None), (date(2026, 9, 7), date(2026, 9, 1))]:
            with self.assertRaises(ValueError):
                agregar_condicion_fechas([], [], inicio, fin)

    def test_carga_envia_rango_hasta_repositorio(self):
        inicio, fin = date(2026, 9, 1), date(2026, 9, 7)
        app_ui.cargar_casos_soporte_filtrados_cache.clear()
        app_ui.cargar_casos_filtrados_cache.clear()
        with patch.object(app_logic, "exigir_contexto_lectura"), patch.object(app_logic, "read_table_filtered", return_value=pd.DataFrame()) as leer:
            app_ui.cargar_casos_soporte_filtrados_cache(fecha_inicio=inicio, fecha_fin=fin)
        self.assertEqual(inicio, leer.call_args.kwargs["fecha_inicio"])
        self.assertEqual(fin, leer.call_args.kwargs["fecha_fin"])
        self.assertIsNone(leer.call_args.kwargs["anio"])
        self.assertIsNone(leer.call_args.kwargs["mes"])
        app_ui.cargar_casos_soporte_filtrados_cache.clear()
        app_ui.cargar_casos_filtrados_cache.clear()

    def test_vista_hoy_mes_y_rango_sin_ampliar_consulta_vacia(self):
        from streamlit.testing.v1 import AppTest
        app = AppTest.from_string('''
from unittest.mock import patch
import streamlit as st
import pandas as pd
import app_ui as ui
def cargar(**kwargs):
    st.session_state["consulta"] = kwargs
    st.session_state["llamadas"] = st.session_state.get("llamadas", 0) + 1
    return pd.DataFrame()
with patch.object(ui, "cargar_casos_soporte_filtrados_cache", side_effect=cargar), patch.object(ui, "selector_periodo_sql", side_effect=AssertionError("No consultar todos los meses")):
    ui.vista_casos()
''').run(timeout=10)
        hoy = pd.Timestamp.now(tz="America/Bogota").date()
        self.assertFalse(app.exception)
        self.assertEqual("Hoy", app.radio(key="periodo_consulta_casos").value)
        self.assertEqual({"fecha_inicio": hoy, "fecha_fin": hoy}, app.session_state["consulta"])
        self.assertEqual(1, app.session_state["llamadas"])
        app.radio(key="periodo_consulta_casos").set_value("Ayer y hoy").run()
        self.assertEqual(hoy - timedelta(days=1), app.session_state["consulta"]["fecha_inicio"])
        app.radio(key="periodo_consulta_casos").set_value("Mes completo").run()
        self.assertEqual(hoy.replace(day=1), app.session_state["consulta"]["fecha_inicio"])
        self.assertEqual((pd.Timestamp(hoy) + pd.offsets.MonthEnd(0)).date(), app.session_state["consulta"]["fecha_fin"])
        app.number_input(key="anio_consulta_casos").set_value(2024).run()
        app.selectbox(key="mes_consulta_casos").select(2).run()
        self.assertEqual(date(2024, 2, 29), app.session_state["consulta"]["fecha_fin"])
        app.radio(key="periodo_consulta_casos").set_value("Rango de fechas").run()
        app.date_input(key="rango_consulta_casos").set_value((date(2026, 9, 1), date(2026, 9, 7))).run()
        self.assertFalse(app.exception)
        self.assertEqual({"fecha_inicio": date(2026, 9, 1), "fecha_fin": date(2026, 9, 7)}, app.session_state["consulta"])
        llamadas = app.session_state["llamadas"]
        app.date_input(key="rango_consulta_casos").set_value((date(2026, 9, 1),)).run()
        self.assertFalse(app.exception)
        self.assertEqual(llamadas, app.session_state["llamadas"])


if __name__ == "__main__":
    unittest.main()
