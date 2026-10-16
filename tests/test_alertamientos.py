import unittest

import pandas as pd
from streamlit.testing.v1 import AppTest

import app_logic
from services.incidentes_sla import (
    CONSULTA, INCIDENTE, SEGURIDAD, agregar_evaluacion, clasificacion_gestion, evaluar_incidente,
)


def registro(**cambios):
    return {"numero": "INC001", "categoria": INCIDENTE, "prioridad": "Alto",
            "breve_descripcion": "Alarma de CPU en servidor", "creado": "2026-10-01 08:00",
            **cambios}


class AlertamientosTest(unittest.TestCase):
    def test_alarma_sin_afectacion_usa_consulta_y_sus_tiempos_sin_modificar_origen(self):
        row = registro()
        resultado = evaluar_incidente(row, "2026-10-16 10:00")
        self.assertEqual(CONSULTA, resultado["categoria_gestion"])
        self.assertTrue(resultado["categoria_ajustada"])
        self.assertTrue(resultado["revision_alertamiento"])
        self.assertEqual(36, resultado["objetivo_resolucion_horas"])
        self.assertEqual("Moderado", resultado["prioridad_gestion"])
        self.assertEqual(INCIDENTE, row["categoria"])
        self.assertEqual("Alto", row["prioridad"])
        self.assertEqual(app_logic.TIPIFICACION_ATENCION_NOC, app_logic.tipificacion_incidente(row))
        self.assertFalse(app_logic.es_incidente_real(row))
        self.assertEqual("Si", app_logic.es_alerta_incidente(row))
        self.assertEqual(36, app_logic.sla_objetivo_horas_incidente(row))
        self.assertEqual(30, app_logic.obtener_sla_respuesta_minutos(row))

    def test_falsa_alarma_y_negaciones_no_se_toman_como_caidas(self):
        for nota in ("Falsa alarma sin afectación del servicio", "Alarma sin indisponibilidad",
                     "No se evidenció caída", "No hubo degradación", "Se descarta una caída",
                     "Falso positivo; no se detectó falla del servicio"):
            with self.subTest(nota=nota):
                resultado = clasificacion_gestion(registro(observaciones_trabajo=nota))
                self.assertEqual(CONSULTA, resultado["categoria_gestion"])
                self.assertEqual("", resultado["evidencia_afectacion"])
        resultado = clasificacion_gestion(registro(observaciones_trabajo="Falsa alarma sin afectación"))
        self.assertEqual("", resultado["revision_alertamiento"])
        self.assertIn("falsa alarma", resultado["evidencia_descarte"])

    def test_alarma_con_afectacion_conserva_incidente_y_sla(self):
        for campo in ("descripcion", "observaciones_trabajo", "lista_notas_trabajo", "actualizaciones"):
            for nota in ("Caída del servicio confirmada", "Servicio con degradación", "OCSP down",
                         "El portal no responde", "Certificado SSL vencido", "No generó alerta"):
                with self.subTest(campo=campo, nota=nota):
                    row = registro(**{campo: nota})
                    resultado = evaluar_incidente(row, "2026-10-16 10:00")
                    self.assertEqual(INCIDENTE, resultado["categoria_gestion"])
                    self.assertFalse(resultado["categoria_ajustada"])
                    self.assertIn(campo, resultado["evidencia_afectacion"])
                    self.assertEqual(8, resultado["objetivo_resolucion_horas"])
                    self.assertTrue(app_logic.es_incidente_real(row))

    def test_contradiccion_y_consulta_con_caida_requieren_revision(self):
        row = registro(descripcion="Caída del portal", observaciones_trabajo="Falsa alarma sin afectación")
        resultado = clasificacion_gestion(row)
        self.assertEqual(INCIDENTE, resultado["categoria_gestion"])
        self.assertIn("revisar el historial", resultado["revision_alertamiento"])
        resultado = clasificacion_gestion(registro(categoria=CONSULTA, descripcion="Caída confirmada"))
        self.assertEqual(CONSULTA, resultado["categoria_gestion"])
        self.assertIn("reclasificar como incidente", resultado["revision_alertamiento"])

    def test_descartar_una_falla_no_oculta_otra_en_la_misma_nota(self):
        resultado = clasificacion_gestion(registro(descripcion="Sin caída, pero hay lentitud confirmada"))
        self.assertEqual(INCIDENTE, resultado["categoria_gestion"])
        self.assertIn("lentitud", resultado["evidencia_afectacion"])

    def test_categoria_noc_numero_inc_y_palabras_parciales_no_crean_alertas(self):
        for cambios in ({"breve_descripcion": "Revisión del servidor", "creado_por": "noc"},
                        {"breve_descripcion": "Consultar escalarmas y desalertado"},
                        {"breve_descripcion": "Incidente de red", "grupo_asignacion": "Alertas NOC"}):
            resultado = clasificacion_gestion(registro(**cambios))
            self.assertFalse(resultado["es_alertamiento"])
            self.assertEqual(INCIDENTE, resultado["categoria_gestion"])

    def test_alerta_de_seguridad_con_evidencia_se_conserva(self):
        row = registro(categoria=SEGURIDAD, descripcion="Alerta por acceso no autorizado")
        self.assertEqual(SEGURIDAD, clasificacion_gestion(row)["categoria_gestion"])
        self.assertEqual(app_logic.TIPO_INCIDENTE_SEGURIDAD, app_logic.tipo_incidente_auto(row))

    def test_alarma_con_categoria_desconocida_y_caida_no_se_pierde_como_consulta(self):
        row = registro(categoria="Infraestructura", descripcion="Caída confirmada del servicio")
        self.assertTrue(app_logic.es_incidente_real(row))
        self.assertIn(app_logic.tipificacion_incidente(row), app_logic.TIPIFICACIONES_INCIDENTE_REAL)
        self.assertTrue(clasificacion_gestion(row)["revision_alertamiento"])

    def test_reevaluar_historial_y_datos_nulos_no_pierde_originales_ni_usa_inferencias_viejas(self):
        base = pd.DataFrame([registro(descripcion=pd.NA, tipificacion_auto="Incidente Interno"),
                             registro(numero="INC002", descripcion="Caída del servicio")])
        evaluado = app_logic.agregar_campos_sla_incidentes(base)
        repetido = agregar_evaluacion(evaluado)
        self.assertEqual([CONSULTA, INCIDENTE], repetido["categoria_gestion"].tolist())
        self.assertEqual([INCIDENTE, INCIDENTE], repetido["categoria"].tolist())
        self.assertEqual([False, True], evaluado["aplica_sla_incidente"].tolist())
        self.assertEqual([36, 8], evaluado["sla_objetivo_horas"].tolist())
        self.assertTrue(agregar_evaluacion(pd.DataFrame()).empty)

    def test_panel_permite_revisar_y_filtrar_sin_perder_el_incidente_real(self):
        codigo = '''
import pandas as pd
from tests.test_alertamientos import registro
from components.incidentes_ans import render_panel_incidentes_ans
render_panel_incidentes_ans(pd.DataFrame([
 registro(numero="INC_ALERTA"),
 registro(numero="INC_FALSA", descripcion="Falsa alarma sin afectación"),
 registro(numero="INC_CAIDA", descripcion="Caída del servicio confirmada"),
]), "alertas_test", "Incidentes")
'''
        app = AppTest.from_string(codigo.encode("ascii", "backslashreplace").decode("ascii"), default_timeout=20).run()
        self.assertFalse(app.exception)
        self.assertEqual("3", app.metric[0].value)
        revision = next(t.value for t in app.dataframe if "Registro" in t.value)
        self.assertEqual(3, len(revision))
        app.checkbox(key="alertas_test_alertas_pendientes").check().run()
        self.assertFalse(app.exception)
        revision = next(t.value for t in app.dataframe if "Registro" in t.value)
        self.assertEqual(["INC_ALERTA"], revision["Registro"].tolist())
        app.selectbox(key="alertas_test_categoria").select(CONSULTA).run()
        self.assertFalse(app.exception)
        self.assertEqual("2", app.metric[0].value)
        detalle = next(t.value for t in app.dataframe if "Incidente" in t.value)
        self.assertEqual({"INC_ALERTA", "INC_FALSA"}, set(detalle["Incidente"]))
        self.assertEqual({INCIDENTE}, set(detalle["Categoría de origen"]))
        self.assertIn("Motivo de clasificación", detalle)


if __name__ == "__main__":
    unittest.main()
