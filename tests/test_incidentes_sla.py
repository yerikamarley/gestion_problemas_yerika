import unittest
from unittest.mock import MagicMock, patch

import pandas as pd
from streamlit.testing.v1 import AppTest

import app_logic
from components.incidentes_ans import grafica_cumplimiento, tabla_resumen_categorias
from services.incidentes_sla import (
    ATENCION_MINUTOS, CONSULTA, INCIDENTE, SEGURIDAD, SIN_CATEGORIA,
    agregar_evaluacion, cargar_horario, categoria_gestion, evaluar_incidente,
    objetivos, segundos_soporte, resumen_etapa, tabla_objetivos, vencimiento_soporte,
)
from services.servicenow import mapear_documento


AHORA = "2026-10-15 17:00:00"


def incidente(**cambios):
    return {"numero": "INC100001", "categoria": "Incidente", "prioridad": "1 - Crítica",
            "creado": "2026-10-05 08:00:00", "fecha_notificacion": "2026-10-05 08:00:00",
            "fecha_atencion": "2026-10-05 08:15:00", "fecha_resolucion": "2026-10-05 12:00:00",
            "estado": "Resuelto", "cerrado": "2026-10-07 17:00:00", **cambios}


def muestra_sin_fecha_atencion():
    """Datos ficticios con la distribución 34 finalizados / 3 abiertos."""
    filas = []
    inicio = "2026-10-01 08:00:00"
    for indice in range(34):
        categoria = INCIDENTE if indice < 26 else CONSULTA
        prioridad = "Alto" if indice < 9 else "Moderado"
        horas = objetivos(categoria, prioridad)[1]
        fin = vencimiento_soporte(inicio, (horas + (1 if indice in (1, 2, 33) else -1)) * 3600)
        filas.append(incidente(numero=f"INC{100001 + indice}", categoria=categoria, prioridad=prioridad,
                               creado=inicio, fecha_notificacion=inicio, fecha_atencion="", fecha_resolucion=fin.isoformat(),
                               cerrado=fin.isoformat() if indice != 33 else "", estado="Cerrado" if indice != 33 else "Resuelto"))
    for indice in range(3):
        fecha = inicio if indice < 2 else "2026-10-15 16:50:00"
        filas.append(incidente(numero=f"INC{100035 + indice}", categoria=CONSULTA, prioridad="Moderado",
                               creado=fecha, fecha_notificacion=fecha, fecha_atencion="", fecha_resolucion="",
                               cerrado="", estado="Activo"))
    return pd.DataFrame(filas)


class IncidentesSlaTest(unittest.TestCase):
    def test_las_nueve_combinaciones_de_la_tabla_y_sus_limites(self):
        esperado = {INCIDENTE: [(15, 4), (20, 8), (30, 36), (30, 72)],
                    SEGURIDAD: [(15, 4), (20, 8), (30, 24), (30, 48)],
                    CONSULTA: [(30, 36)]}
        for categoria, prioridades in esperado.items():
            nombres = ["Moderado"] if categoria == CONSULTA else list(ATENCION_MINUTOS)
            for prioridad, (minutos, horas) in zip(nombres, prioridades):
                with self.subTest(categoria=categoria, prioridad=prioridad):
                    self.assertEqual((minutos, horas), objetivos(categoria, prioridad))
                    inicio = "2026-10-01 08:00:00"
                    fin_atencion = vencimiento_soporte(inicio, minutos * 60)
                    fin_resolucion = vencimiento_soporte(inicio, horas * 3600)
                    row = incidente(categoria=categoria, prioridad=prioridad, creado=inicio,
                                    fecha_notificacion=inicio, fecha_atencion=fin_atencion, fecha_resolucion=fin_resolucion)
                    exacto = evaluar_incidente(row, AHORA)
                    self.assertEqual("Cumple", exacto["atencion_estado"])
                    self.assertEqual("Cumple", exacto["resolucion_estado"])
                    row.update(fecha_atencion=vencimiento_soporte(fin_atencion, 1),
                               fecha_resolucion=vencimiento_soporte(fin_resolucion, 1))
                    tarde = evaluar_incidente(row, AHORA)
                    self.assertEqual("No cumple", tarde["atencion_estado"])
                    self.assertEqual("No cumple", tarde["resolucion_estado"])
        self.assertEqual(9, len(tabla_objetivos()))

    def test_noches_fines_de_semana_y_nueve_horas_por_dia(self):
        self.assertEqual(9 * 3600, segundos_soporte("2026-10-09 08:00", "2026-10-09 17:00"))
        self.assertEqual(15 * 60, segundos_soporte("2026-10-09 16:50", "2026-10-12 08:05"))
        self.assertEqual(0, segundos_soporte("2026-10-09 17:00", "2026-10-12 08:00"))
        self.assertEqual("2026-10-08 17:00:00-05:00", str(vencimiento_soporte("2026-10-05 08:00", 36 * 3600)))
        self.assertEqual("2026-10-14 17:00:00-05:00", str(vencimiento_soporte("2026-10-05 08:00", 72 * 3600)))

    def test_festivos_configurados_y_zona_colombia(self):
        horario = {**cargar_horario(), "festivos": ["2026-10-12"]}
        self.assertEqual(15 * 60, segundos_soporte("2026-10-09 16:50", "2026-10-13 08:05", horario))
        self.assertEqual("2026-10-13 08:05:00-05:00", str(vencimiento_soporte("2026-10-09 16:50", 900, horario)))
        row = incidente(fecha_notificacion="2026-10-05T13:00:00Z", fecha_atencion="2026-10-05T13:15:01Z")
        self.assertEqual("No cumple", evaluar_incidente(row, AHORA)["atencion_estado"])

    def test_clasificacion_oficial_prevalece_y_no_inventa_prioridad(self):
        for valor, esperado in [("1. Incidente", INCIDENTE), ("2 - Incidente Seguridad", SEGURIDAD),
                                ("3. Consulta / Ayuda", CONSULTA), ("hardware", SIN_CATEGORIA)]:
            self.assertEqual(esperado, categoria_gestion(valor))
        consulta = incidente(categoria=CONSULTA, empresa="Cliente externo", prioridad="Alto")
        resultado = evaluar_incidente(consulta, AHORA)
        self.assertEqual("Moderado", resultado["prioridad_gestion"])
        self.assertEqual(36, resultado["objetivo_resolucion_horas"])
        self.assertTrue(resultado["validacion_clasificacion"])
        self.assertEqual("Alto", consulta["prioridad"])
        self.assertEqual(app_logic.TIPIFICACION_ATENCION_NOC, app_logic.tipificacion_incidente(consulta))
        for prioridad in (None, "", "5 - Planeado", "Dato erróneo"):
            sin_prioridad = evaluar_incidente(incidente(prioridad=prioridad), AHORA)
            self.assertEqual("Sin clasificación", sin_prioridad["atencion_estado"])
            self.assertIsNone(sin_prioridad["objetivo_resolucion_horas"])

    def test_no_usa_actualizacion_ni_cierre_como_evidencia(self):
        row = incidente(fecha_atencion="", fecha_resolucion="", actualizado="2026-10-05 08:01:00",
                        duracion_horas=1, duracion_segundos=3600)
        resultado = evaluar_incidente(row, AHORA)
        self.assertEqual("Sin fecha de atención", resultado["atencion_estado"])
        self.assertEqual("Sin fecha de resolución", resultado["resolucion_estado"])
        self.assertIsNone(app_logic.calcular_tiempo_respuesta_minutos(row))
        evaluado = app_logic.agregar_campos_sla_incidentes(pd.DataFrame([row])).iloc[0]
        self.assertTrue(pd.isna(evaluado["duracion_sla_horas"]))
        self.assertEqual("Sin fecha de resolución", evaluado["estado_sla"])

    def test_resolucion_precede_al_cierre_y_no_redondea_incumplimiento(self):
        row = incidente(fecha_resolucion="2026-10-05 12:00:01")
        resultado = evaluar_incidente(row, AHORA)
        self.assertEqual("No cumple", resultado["resolucion_estado"])
        self.assertEqual(14401, resultado["resolucion_segundos"])
        row["fecha_resolucion"] = "2026-10-05 12:00:00"
        self.assertEqual("Cumple", evaluar_incidente(row, AHORA)["resolucion_estado"])

    def test_pendientes_vencidos_y_fechas_inconsistentes(self):
        row = incidente(estado="Nuevo", cerrado="", fecha_atencion="", fecha_resolucion="")
        self.assertEqual("Sin fecha de atención", evaluar_incidente(row, "2026-10-05 08:15:00")["atencion_estado"])
        self.assertEqual("Sin fecha de atención", evaluar_incidente(row, "2026-10-05 08:15:01")["atencion_estado"])
        self.assertEqual("Dentro del plazo", evaluar_incidente(row, "2026-10-05 12:00:00")["resolucion_estado"])
        self.assertEqual("Vencido", evaluar_incidente(row, "2026-10-05 12:00:01")["resolucion_estado"])
        row.update(fecha_atencion="2026-10-05 07:59:59")
        self.assertEqual("Fechas inconsistentes", evaluar_incidente(row, AHORA)["atencion_estado"])
        row.update(fecha_atencion="2026-10-20 10:00:00")
        self.assertEqual("Fechas inconsistentes", evaluar_incidente(row, AHORA)["atencion_estado"])

    def test_historial_y_ausencia_de_fecha_no_prueban_que_no_se_atendio(self):
        row = incidente(estado="Activo", cerrado="", fecha_atencion="", fecha_resolucion="",
                        observaciones_trabajo="2026-10-05 08:10 | Nota de Trabajo | Se está revisando el servicio",
                        actualizado="2026-10-05 08:10:00")
        salida = evaluar_incidente(row, AHORA)
        self.assertEqual("Sin fecha de atención", salida["atencion_estado"])
        self.assertIn("primera atención", salida["atencion_motivo"])
        resumen = resumen_etapa(agregar_evaluacion(pd.DataFrame([row]), AHORA), "atencion")
        self.assertEqual(0, resumen["vencidos"])
        self.assertEqual(0, resumen["pendientes"])
        self.assertIsNone(resumen["porcentaje"])

    def test_37_registros_sin_atencion_no_alteran_cumplimiento_resolucion(self):
        base = agregar_evaluacion(muestra_sin_fecha_atencion(), AHORA)
        atencion, resolucion = resumen_etapa(base, "atencion"), resumen_etapa(base, "resolucion")
        self.assertEqual(37, atencion["sin_datos"])
        self.assertEqual(0, atencion["vencidos"])
        self.assertEqual((31, 3, 34, 91.18), tuple(resolucion[k] for k in ("cumple", "no_cumple", "evaluados", "porcentaje")))
        self.assertEqual((3, 2), (resolucion["pendientes"], resolucion["vencidos"]))

    def test_sin_datos_y_pendientes_no_se_ocultan_en_porcentaje(self):
        df = agregar_evaluacion(pd.DataFrame([
            incidente(), incidente(numero="INC2", fecha_resolucion=""),
            incidente(numero="INC3", fecha_resolucion="", fecha_atencion="", cerrado="", estado="Nuevo"),
        ]), AHORA)
        resumen = resumen_etapa(df, "resolucion")
        self.assertEqual({"cumple": 1, "no_cumple": 0, "evaluados": 1, "porcentaje": 100,
                          "pendientes": 1, "vencidos": 1, "sin_datos": 1}, resumen)
        self.assertIsNone(resumen_etapa(df.iloc[1:], "resolucion")["porcentaje"])
        self.assertIsNone(resumen_etapa(df.iloc[0:0], "resolucion")["porcentaje"])
        tabla = tabla_resumen_categorias(df)
        self.assertEqual(1, tabla.iloc[0]["No evaluables resolución"])

    def test_api_mapea_fechas_explicitas_y_respeta_mapa_personalizado(self):
        doc = {"campos_raw": {"number": "INC100001", "opened_at": "2026-10-05 13:00:00",
                               "resolved_at": "2026-10-05 17:00:00", "first_response_at": "2026-10-05 13:15:00"}}
        with patch("services.servicenow.config_value", return_value="{}"):
            mapped = mapear_documento(doc, "incidentes")
        self.assertEqual("2026-10-05 08:00:00", mapped["fecha_notificacion"])
        self.assertEqual("2026-10-05 08:15:00", mapped["fecha_atencion"])
        self.assertEqual("2026-10-05 12:00:00", mapped["fecha_resolucion"])
        doc["campos_raw"]["u_atendido"] = "2026-10-05 13:16:00"
        with patch("services.servicenow.config_value", return_value='{"fecha_atencion":"u_atendido"}'):
            self.assertEqual("2026-10-05 08:16:00", mapear_documento(doc, "incidentes")["fecha_atencion"])

    def test_recupera_fechas_ya_sincronizadas_sin_api_y_sin_sobrescribir(self):
        df = pd.DataFrame([incidente(fecha_atencion="", fecha_resolucion="")])
        conn = MagicMock()
        with patch.object(app_logic, "get_conn", return_value=conn), \
             patch.object(app_logic, "leer_fechas_incidentes", return_value=[("INC100001", {
                 "opened_at": "2026-10-05 15:00:00", "resolved_at": "2026-10-05 17:00:00",
                 "first_response_at": "2026-10-05 13:15:00"})]), \
             patch("services.servicenow.config_value", return_value="{}"), \
             patch.object(app_logic, "config_value", return_value="{}"):
            result = app_logic.completar_fechas_incidentes(df).iloc[0]
        self.assertEqual("2026-10-05 08:00:00", result["fecha_notificacion"])
        self.assertEqual("2026-10-05 08:15:00", result["fecha_atencion"])
        self.assertEqual("2026-10-05 12:00:00", result["fecha_resolucion"])
        conn.close.assert_called_once()


class PanelIncidentesTest(unittest.TestCase):
    def test_grafica_agrupa_tipos_y_conserva_totales_y_resultados_sla(self):
        base = agregar_evaluacion(muestra_sin_fecha_atencion(), AHORA)
        fig = grafica_cumplimiento(base, "resolucion")
        cantidades = {}
        resultados = {}
        for traza in fig.data:
            resultados[traza.name] = sum(traza.x)
            for tipo, cantidad in zip(traza.y, traza.x):
                cantidades[tipo] = cantidades.get(tipo, 0) + cantidad
        self.assertEqual({INCIDENTE: 26, CONSULTA: 11}, cantidades)
        self.assertEqual({"Cumple": 31, "No cumple": 3, "Dentro del plazo": 1, "Vencido": 2}, resultados)
        self.assertEqual([INCIDENTE, CONSULTA], list(fig.layout.yaxis.categoryarray))
        seguridad = agregar_evaluacion(pd.DataFrame([incidente(categoria=SEGURIDAD)]), AHORA)
        fig_seguridad = grafica_cumplimiento(seguridad, "resolucion")
        self.assertEqual([SEGURIDAD], list(fig_seguridad.layout.yaxis.categoryarray))

    def test_panel_muestra_solo_sla_resolucion_sin_perder_registros(self):
        codigo = '''
from unittest.mock import patch
from tests.test_incidentes_sla import muestra_sin_fecha_atencion, AHORA
from services.incidentes_sla import agregar_evaluacion
from components.incidentes_ans import render_panel_incidentes_ans
df = muestra_sin_fecha_atencion()
with patch("components.incidentes_ans.agregar_evaluacion", side_effect=lambda datos, horario=None: agregar_evaluacion(datos, AHORA, horario)):
 render_panel_incidentes_ans(df, "sin_fecha", "Dashboard Incidentes")
'''
        app = AppTest.from_string(codigo, default_timeout=20).run()
        self.assertFalse(app.exception)
        self.assertEqual(["37", "3", "91.18%"], [m.value for m in app.metric])
        self.assertEqual("SLA de resolución", app.metric[2].label)
        textos = [x.value for x in app.markdown] + [x.value for x in app.caption]
        self.assertFalse(any("atención" in texto.lower() or "no calculable" in texto.lower() for texto in textos))
        self.assertEqual(2, len(app.get("plotly_chart")))
        detalle = next(item.value for item in app.dataframe if "Incidente" in item.value.columns)
        self.assertEqual(37, len(detalle))
        self.assertIn("Resultado resolución", detalle)
        for tabla in app.dataframe:
            self.assertFalse(any("atención" in columna.lower() for columna in tabla.value.columns))

    def test_vista_incidentes_conserva_fechas_y_motivos_en_la_descarga(self):
        codigo = '''
import streamlit as st
from unittest.mock import patch
import app_ui as ui
from tests.test_incidentes_sla import muestra_sin_fecha_atencion
from app_logic import preparar_incidentes, agregar_campos_sla_incidentes
df = agregar_campos_sla_incidentes(preparar_incidentes(muestra_sin_fecha_atencion()))
with patch.object(ui, "render_consulta_traza_servicenow"), patch.object(ui, "selector_periodo_sql", return_value=(2026, 10, "Octubre 2026")), patch.object(ui, "cargar_incidentes_sla_filtrados_cache", return_value=df), patch.object(ui, "render_descarga_dataframe") as descarga:
 ui.vista_incidentes()
 st.session_state["columnas_descarga"] = descarga.call_args.args[0].columns.tolist()
'''
        app = AppTest.from_string(codigo, default_timeout=20).run()
        self.assertFalse(app.exception)
        for columna in ("fecha_atencion", "fecha_resolucion", "atencion_motivo", "resolucion_motivo"):
            self.assertIn(columna, app.session_state["columnas_descarga"])

    def test_filtra_categoria_y_muestra_objetivos_coordinador_y_datos_faltantes(self):
        app = AppTest.from_string('''
import streamlit as st
import pandas as pd
from components.incidentes_ans import render_panel_incidentes_ans
rows = [
 {"numero":"INC1", "categoria":"Incidente", "prioridad":"Crítico", "estado":"Cerrado", "creado":"2026-10-01 08:00", "fecha_atencion":"2026-10-01 08:15", "fecha_resolucion":"2026-10-01 12:00"},
 {"numero":"INC2", "categoria":"Incidente Seguridad", "prioridad":"Alto", "estado":"Cerrado", "creado":"2026-10-01 08:00"},
 {"numero":"INC3", "categoria":"Consulta / Ayuda", "prioridad":"Moderado", "estado":"Nuevo", "creado":"2026-10-01 08:00"},
]
render_panel_incidentes_ans(pd.DataFrame(rows), "prueba", "Dashboard Incidentes", "Octubre 2026")
'''.encode("ascii", "backslashreplace").decode("ascii"), default_timeout=15).run()
        self.assertFalse(app.exception)
        self.assertEqual("3", app.metric[0].value)
        app.selectbox(key="prueba_categoria").select(SEGURIDAD).run()
        self.assertFalse(app.exception)
        self.assertEqual("1", app.metric[0].value)
        self.assertEqual("No calculable", app.metric[2].value)
        self.assertTrue(any("Especialista en seguridad" in item.value for item in app.markdown))
        self.assertTrue(any("sin fecha de resolución" in item.value for item in app.caption))
        detalle = next(item.value for item in app.dataframe if "Incidente" in item.value.columns)
        self.assertEqual(["INC2"], detalle["Incidente"].tolist())
        app.selectbox(key="prueba_categoria").select(CONSULTA).run()
        self.assertFalse(app.exception)
        self.assertTrue(any("Moderado" in item.value and "4 días" in item.value for item in app.caption))
        self.assertEqual(["Moderado"], app.multiselect[0].options)

    def test_dashboard_y_kpi_usan_mismos_tiempos_y_conservan_presentacion(self):
        codigo = '''
import streamlit as st
import pandas as pd
from unittest.mock import patch
import app_ui as ui
from app_logic import preparar_incidentes, aplicar_clasificacion_incidentes
df = aplicar_clasificacion_incidentes(preparar_incidentes(pd.DataFrame([
 {"numero":"INC100001", "categoria":"Incidente", "prioridad":"Critico", "estado":"Resuelto", "creado":"2026-10-01 08:00:00", "fecha_atencion":"2026-10-01 08:15:01", "fecha_resolucion":"2026-10-01 12:00:00"}
])))
vista = st.radio("Vista", ["Dashboard", "KPI"])
if vista == "Dashboard":
 with patch.object(ui, "selector_periodo_sql", return_value=(2026, 10, "Octubre 2026")), patch.object(ui, "cargar_incidentes_filtrados_cache", return_value=df):
  ui.dashboard_incidentes()
else:
 ui.render_kpi_incidentes(df, "Octubre 2026")
'''
        app = AppTest.from_string(codigo, default_timeout=15).run()
        self.assertFalse(app.exception)
        self.assertEqual(["1", "0", "100.00%"], [m.value for m in app.metric])
        app.radio[0].set_value("KPI").run()
        self.assertFalse(app.exception)
        self.assertEqual(["1", "0", "100.00%"], [m.value for m in app.metric])
        app.toggle(key="kpi_incidentes_ans_diapositiva").set_value(True).run()
        self.assertFalse(app.exception)
        self.assertEqual(1, len(app.get("iframe")))


if __name__ == "__main__":
    unittest.main()
