import unittest

import pandas as pd

import app_ui as ui
from services.casos_historial import (
    COL_CAUSA, COL_ESTADO_CAUSA, agregar_analisis_historial,
    analizar_caso, detalle_causas_documentadas,
)
from tests.test_resumen_ejecutivo_casos import caso


class CausasDocumentadasTest(unittest.TestCase):
    def test_amplia_grupos_solo_con_explicacion_registrada(self):
        for causa, grupo in [
            ("Caída del portal", "Caída o indisponibilidad del servicio"),
            ("Enlace de descarga vencido", "Enlace o proceso de descarga inválido"),
            ("Configuración incorrecta", "Instalación o configuración incorrecta"),
            ("Token físico defectuoso", "Dispositivo no reconocido o defectuoso"),
            ("Credenciales incorrectas", "Acceso o autenticación rechazados"),
            ("Timeout del servicio", "Lentitud o tiempo de espera del servicio"),
            ("No podemos ver consumos", "Consumos o reportes no disponibles"),
            ("Correo rebotado", "Correo o acuse no entregado"),
            ("Firma inválida", "Error de firma o validación documental"),
            ("Datos inconsistentes", "Datos o documentos inconsistentes"),
            ("Certificado no disponible", "Certificado ausente o no disponible"),
        ]:
            with self.subTest(causa=causa):
                registrado = analizar_caso({"causa": causa})
                self.assertEqual(grupo, registrado["Grupo de causa"])
                self.assertEqual(causa, registrado[COL_CAUSA])
                solicitud = analizar_caso({"descripcion": causa})
                self.assertEqual("Pendiente de validar", solicitud[COL_ESTADO_CAUSA])

    def test_detalle_conserva_desconocidas_evidencia_y_denominador(self):
        base = agregar_analisis_historial(pd.DataFrame([
            {"numero": "A", "causa": "Parámetro X deshabilitado"},
            {"numero": "B", "causa": "parametro  x deshabilitado"},
            {"numero": "C", "causa": "Inconsistencia en licencia contratada"},
            {"numero": "D", "causa": "Driver incompatible"},
            {"numero": "E", "descripcion": "Necesito instalar token físico"},
            {"numero": "D", "causa": "Driver incompatible"},
        ]))
        detalle = detalle_causas_documentadas(base)
        self.assertEqual(4, detalle["Casos"].sum())
        self.assertEqual(3, len(detalle))
        parametro = detalle.iloc[0]
        self.assertEqual("Parámetro X deshabilitado", parametro["Causa documentada"])
        self.assertEqual(2, parametro["Casos"])
        self.assertEqual(50, parametro["% documentados"])
        self.assertEqual("A, B", parametro["Ejemplos"])
        self.assertIn("Parámetro X deshabilitado", parametro["Evidencia"])
        otras = detalle_causas_documentadas(base, solo_sin_grupo=True)
        self.assertEqual(2, len(otras))
        self.assertEqual(75, otras["% documentados"].sum())
        self.assertFalse(otras["Causa documentada"].eq("Otras causas documentadas").any())

    def test_tabla_vacia_y_sin_diagnostico(self):
        for base in [pd.DataFrame(), agregar_analisis_historial(pd.DataFrame([
            {"numero": "A", "descripcion": "Necesito descargar certificado"}]))]:
            for solo_sin_grupo in [True, False]:
                self.assertTrue(detalle_causas_documentadas(base, solo_sin_grupo).empty)


class OperacionSoporteTest(unittest.TestCase):
    def agentes(self):
        return pd.DataFrame([
            caso("A1", asignado="Ana", estado="Cerrado", cerrado="2026-09-03 08:00"),
            caso("B1", asignado="Beto", estado="Cerrado", cerrado="2026-09-03 08:00"),
            caso("B2", asignado="Beto", estado="Cerrado", cerrado="2026-09-05 08:00"),
            caso("B3", asignado="Beto", estado="Cerrado", cerrado="2026-09-05 08:00"),
            caso("B4", asignado="Beto", estado="Cerrado", cerrado="fecha inválida"),
            caso("B5", asignado="Beto"),
            caso("B6", asignado="Beto", fecha_vencimiento_sla="2026-11-01 08:00"),
            caso("C1", asignado=""),
            caso("B1", asignado="Beto", estado="Cerrado", cerrado="2026-09-03 08:00"),
        ])

    def test_score_cierres_evaluables_y_vencidos_abiertos(self):
        resumen = ui.resumen_agentes_soporte(self.agentes(), ahora="2026-10-10 12:00")
        self.assertEqual(8, resumen["Total asignado"].sum())
        self.assertEqual(5, resumen[ui.TEXT_CERRADOS].sum())
        self.assertEqual(3, resumen[ui.TEXT_ABIERTOS].sum())
        self.assertEqual(4, resumen["Cerrados evaluables"].sum())
        self.assertEqual(2, resumen["Vencidos abiertos"].sum())
        agentes = resumen.set_index(ui.TEXT_RESPONSABLE)
        self.assertEqual(100, agentes.loc["Ana", ui.TEXT_SCORE])
        self.assertEqual(33.33, agentes.loc["Beto", ui.TEXT_SCORE])
        self.assertEqual(1, agentes.loc["Beto", "Sin evaluación SLA"])
        self.assertTrue(pd.isna(agentes.loc["Sin agente", ui.TEXT_SCORE]))
        self.assertEqual(agentes.loc["Beto", ui.TEXT_SCORE], agentes.loc["Beto", "Cumplimiento SLA (%)"])

    def test_total_sla_pesa_casos_no_promedio_de_agentes_y_no_grafica(self):
        from streamlit.testing.v1 import AppTest
        app = AppTest.from_string('''
import app_ui as ui
from tests.test_soporte_operacion import OperacionSoporteTest
ui.render_agentes_soporte(OperacionSoporteTest().agentes())
''').run()
        self.assertFalse(app.exception)
        self.assertEqual("50%", app.metric[0].value)
        self.assertEqual(0, len(app.get("plotly_chart")))
        self.assertEqual(1, len(app.dataframe))

    def test_productos_conserva_todos_deduplica_y_separa_motivos(self):
        datos = [caso(f"C{i}", producto=f"Producto {i}", causa="Token físico",
                      descripcion="Necesito instalar token físico") for i in range(8)]
        datos += [caso("A", producto="Producto 0", causa="Driver incompatible", descripcion="No puedo firmar"),
                  caso("B", producto="Producto 0", causa="Driver incompatible", descripcion="No puedo firmar"),
                  caso("A", producto="Producto 0", causa="Driver incompatible", descripcion="No puedo firmar"),
                  caso("Z", producto="", causa="Correo rebotado")]
        resumen = ui.resumen_productos_causas_soporte(pd.DataFrame(datos))
        self.assertEqual(11, resumen["Casos"].sum())
        self.assertEqual(9, len(resumen))
        self.assertEqual(3, resumen["Con causa documentada"].sum())
        self.assertEqual(8, resumen["Sin causa documentada"].sum())
        productos = resumen.set_index("Producto")
        self.assertEqual(3, productos.loc["Producto 0", "Casos"])
        self.assertEqual("Driver incompatible (2)", productos.loc["Producto 0", "Principales causas"])
        self.assertIn("Instalación o configuración (1)", productos.loc["Producto 0", "Principales motivos iniciales"])
        self.assertEqual("Sin causa documentada", productos.loc["Producto 1", "Principales causas"])
        self.assertIn("Sin producto", productos.index)
        self.assertNotIn("Otros productos", productos.index)

    def test_operacion_vacia(self):
        self.assertTrue(ui.resumen_agentes_soporte(pd.DataFrame()).empty)
        self.assertTrue(ui.resumen_productos_causas_soporte(pd.DataFrame()).empty)


if __name__ == "__main__":
    unittest.main()
