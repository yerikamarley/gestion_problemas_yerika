import unittest

import pandas as pd

from services.clientes_experiencia import seguimiento_clientes


def datos():
    casos = pd.DataFrame([
        {"numero": "C1", "cliente_clave": "Banco A", "estado": "En proceso", "creado": "2026-10-01 08:00:00",
         "fecha_vencimiento_sla": "2026-10-09 10:00:00", "prioridad": "3 - Moderada", "asignado": "Ana",
         "descripcion": "Necesito instalar el token físico"},
        {"numero": "C2", "cliente_clave": "Banco B", "estado": "Nuevo", "creado": "2026-10-02 09:00:00",
         "fecha_vencimiento_sla": "2026-10-15 10:00:00", "prioridad": "1 - Crítica", "asignado": "",
         "descripcion": "No funciona la firma digital"},
        {"numero": "C3", "cliente_clave": "Banco B", "estado": "Cerrado", "creado": "2026-10-08 09:00:00",
         "cerrado": "2026-10-09 09:00:00", "fecha_vencimiento_sla": "2026-10-12 09:00:00",
         "descripcion": "Falla de descarga"},
        {"numero": "C4", "cliente_clave": "Banco A", "estado": "Nuevo", "creado": "2026-10-03 09:00:00",
         "prioridad": "4 - Baja", "descripcion": "Consulta de consumos"},
    ])
    incidentes = pd.DataFrame([
        {"numero": "I1", "cliente_clave": "Banco A", "estado": "En proceso", "creado": "2026-10-02 09:00:00",
         "asignado_a": "Luis", "breve_descripcion": "Portal caído", "sla_objetivo_horas": 1},
        {"numero": "I2", "cliente_clave": "Banco C", "estado": "Resuelto", "creado": "2026-10-09 08:00:00",
         "cerrado": "2026-10-09 09:00:00"},
    ])
    return casos, incidentes


class SeguimientoClientesTest(unittest.TestCase):
    def test_conteos_unicos_solo_abiertos_y_motivos(self):
        casos, incidentes = datos()
        duplicados = pd.concat([casos, casos.iloc[[0]]], ignore_index=True)
        analisis = seguimiento_clientes(duplicados, incidentes, ahora="2026-10-10 15:00:00")
        self.assertEqual(4, len(analisis["casos"]))
        self.assertEqual(4, len(analisis["pendientes"]))
        self.assertEqual(3, analisis["motivos"]["Casos abiertos"].sum())
        clientes = analisis["clientes"].set_index("Cliente")
        self.assertEqual(3, clientes.loc["Banco A", "Pendientes"])
        self.assertEqual(1, clientes.loc["Banco A", "Casos vencidos"])
        self.assertEqual(0, clientes.loc["Banco C", "Pendientes"])
        self.assertNotIn("C3", analisis["pendientes"]["Número"].tolist())
        self.assertEqual("Banco A", analisis["clientes"].iloc[0]["Cliente"])

    def test_vencidos_prioridad_antiguedad_y_cerrados_al_final(self):
        casos, incidentes = datos()
        analisis = seguimiento_clientes(casos, incidentes, ahora="2026-10-10 15:00:00")
        self.assertEqual(["C1", "C2", "C4", "I1"], analisis["pendientes"]["Número"].tolist())
        self.assertEqual(["C1", "C2", "C4", "C3"], analisis["casos"]["numero"].tolist())
        self.assertEqual("Vencido", analisis["pendientes"].iloc[0]["Estado SLA"])
        self.assertEqual("Sin asignar", analisis["pendientes"].iloc[1]["Responsable"])

    def test_fecha_de_colombia_y_sla_desconocido_no_se_inventa(self):
        casos, incidentes = datos()
        analisis = seguimiento_clientes(casos, incidentes, ahora="2026-10-10T20:00:00Z")
        pendientes = analisis["pendientes"].set_index("Número")
        self.assertEqual("Sin fecha de vencimiento", pendientes.loc["I1", "Estado SLA"])
        self.assertEqual(9.3, pendientes.loc["C1", "Días abiertos"])
        self.assertEqual("America/Bogota", str(pendientes.loc["C1", "Creado"].tz))
        casos.loc[0, "creado"] = "2026-10-12 09:00:00"
        analisis = seguimiento_clientes(casos, incidentes, ahora="2026-10-10 15:00:00")
        fila = analisis["pendientes"].set_index("Número").loc["C1"]
        self.assertEqual("Fechas inconsistentes", fila["Estado SLA"])
        self.assertTrue(pd.isna(fila["Días abiertos"]))

    def test_incidencia_vencida_por_fecha_de_origen(self):
        casos, incidentes = datos()
        incidentes.loc[0, "fecha_vencimiento_sla"] = "2026-10-03 09:00:00"
        analisis = seguimiento_clientes(casos, incidentes, ahora="2026-10-10 15:00:00")
        self.assertEqual(1, analisis["clientes"].iloc[0]["Incidentes vencidos"])
        self.assertEqual("I1", analisis["pendientes"].iloc[1]["Número"])

    def test_base_vacia_y_solo_cerrados(self):
        vacio = seguimiento_clientes(pd.DataFrame(), pd.DataFrame())
        self.assertTrue(vacio["pendientes"].empty)
        self.assertTrue(vacio["clientes"].empty)
        self.assertTrue(vacio["motivos"].empty)
        casos, incidentes = datos()
        cerrado = seguimiento_clientes(casos.iloc[[2]], incidentes.iloc[[1]])
        self.assertTrue(cerrado["pendientes"].empty)
        self.assertTrue(cerrado["motivos"].empty)


class ClientesExperienciaUiTest(unittest.TestCase):
    def app(self):
        import app_ui  # Inicializar fuera del contexto de AppTest.
        from streamlit.testing.v1 import AppTest
        return AppTest.from_string('''
from unittest.mock import patch
import app_ui as ui
from tests.test_clientes_experiencia import datos
casos, incidentes = datos()
with patch.object(ui, "selector_periodo_multi_sql", return_value=(2026, 10, "Octubre 2026")), \\
     patch.object(ui, "cargar_casos_clientes_clave_filtrados_cache", return_value=casos), \\
     patch.object(ui, "cargar_incidentes_clientes_clave_filtrados_cache", return_value=incidentes):
    ui.dashboard_clientes_clave()
''', default_timeout=30).run()

    def test_pendientes_por_defecto_filtro_y_historico(self):
        app = self.app()
        self.assertEqual(0, len(app.exception))
        self.assertEqual("Pendientes", app.radio(key="clientes_clave_seguimiento").value)
        self.assertEqual(2, len(app.get("plotly_chart")))
        self.assertEqual("Banco A", app.dataframe[0].value.iloc[0]["Cliente"])
        self.assertEqual("C1", app.dataframe[1].value.iloc[0]["Número"])
        self.assertNotIn("C3", app.dataframe[1].value["Número"].tolist())
        app.radio(key="clientes_clave_filtro_pendientes").set_value("Sin asignar").run()
        self.assertEqual(0, len(app.exception))
        self.assertEqual({"Sin asignar"}, set(app.dataframe[1].value["Responsable"]))
        app.radio(key="clientes_clave_seguimiento").set_value("Histórico del período").run()
        self.assertEqual(0, len(app.exception))
        self.assertEqual(0, len(app.get("plotly_chart")))
        self.assertEqual("C3", app.dataframe[1].value.iloc[-1]["Número"])

    def test_grupo_y_cliente_individual_y_seleccion_vacia(self):
        app = self.app()
        app.radio(key="clientes_clave_grupo").set_value("Asobancaria").run()
        self.assertEqual(0, len(app.exception))
        app.radio[1].set_value("Un solo cliente").run()
        self.assertEqual(0, len(app.exception))
        self.assertTrue(any("No hay casos" in mensaje.value for mensaje in app.success))
        app.radio[1].set_value("Selección personalizada").run()
        self.assertEqual(0, len(app.exception))
        self.assertEqual([], app.multiselect[0].value)
        self.assertTrue(any("Selecciona al menos" in mensaje.value for mensaje in app.warning))


if __name__ == "__main__":
    unittest.main()
