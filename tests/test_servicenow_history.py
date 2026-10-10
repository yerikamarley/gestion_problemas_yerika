import json
import unittest
from unittest.mock import MagicMock, patch

import pandas as pd
from streamlit.testing.v1 import AppTest

import app_logic
import app_ui
from repositories import servicenow as repository
from repositories import tables
from services.servicenow import ServiceNowClient, mapear_documento
from services.casos_sla import tabla_casos_plataforma
from tests.test_servicenow import caso, documento, field, response


class ConversationHistoryTest(unittest.TestCase):
    def setUp(self):
        config = {"SERVICENOW_INSTANCE": "test.service-now.com", "SERVICENOW_USER": "api-test", "SERVICENOW_PASSWORD": "test"}
        patcher = patch("services.servicenow.config_value", side_effect=lambda k, d="": config.get(k, d))
        patcher.start()
        self.addCleanup(patcher.stop)

    def test_correo_html_y_asunto_entran_al_historial_y_al_analisis(self):
        email = {"sys_id": field("c" * 32), "type": field("received"), "sys_created_on": field("2026-10-10 10:00:00"),
                 "origemail": field("cliente@example.test"), "recipients": field("soporte@example.test"),
                 "subject": field("Error al firmar"), "body_text": field(""),
                 "body": field("<style>.noise{}</style><p>El cliente no puede firmar &amp; requiere ayuda.</p><script>noise()</script>")}
        with ServiceNowClient() as client, patch.object(client._session, "get", side_effect=[
            response([], 0), response([], 0), response([email], 1),
        ]):
            doc = client.consolidar(caso())
        event = doc["historial_conversacion"][0]
        self.assertEqual("Correo Entrante", event["tipo"])
        self.assertEqual("Error al firmar", event["asunto"])
        self.assertIn("no puede firmar & requiere ayuda", event["contenido"])
        self.assertNotIn("<p>", event["contenido"])
        self.assertNotIn("noise", event["contenido"])
        row = mapear_documento(doc)
        self.assertIn("cliente@example.test", row["historial_conversacion"])
        self.assertEqual("Fallas al firmar", app_ui.inferir_detalle_causa_comun(row))
        self.assertEqual("Firma digital", app_ui.inferir_causa_comun_caso(row))
        self.assertNotIn("observaciones_trabajo", row)

    def test_evidencia_solo_en_conversacion_cambia_tipificacion(self):
        row = {"descripcion": "Consulta del cliente", "historial_conversacion": "El correo recibido contiene phishing"}
        self.assertEqual("1 - phishing", app_logic.tipificar_caso(row))
        self.assertIn("phishing", app_ui.texto_caso_para_tipologia_soporte(row))
        self.assertIn("phishing", app_ui.texto_analisis_agendamiento(row))
        visible = tabla_casos_plataforma(pd.DataFrame([row]))
        self.assertEqual(row["historial_conversacion"], visible.iloc[0]["Historial de la conversación"])

    def test_cien_casos_se_consolidan_con_cuatro_llamadas(self):
        cases = []
        for i in range(1, 101):
            raw = caso()
            raw["sys_id"] = field(f"{i:032x}")
            raw["number"] = field(f"CS{i:07d}")
            cases.append(raw)
        emails = [{"sys_id": field("e" * 32), "instance": field("1".zfill(32)), "type": field("received"),
                   "sys_created_on": field("2026-10-10 10:00:00"), "body_text": field("Primer cliente")},
                  {"sys_id": field("f" * 32), "instance": field("2".zfill(32)), "type": field("sent"),
                   "sys_created_on": field("2026-10-10 11:00:00"), "body_text": field("Segundo cliente")}]
        with ServiceNowClient() as client, patch.object(client._session, "get", side_effect=[
            response(cases, 100), response([], 0), response([], 0), response(emails, 2),
        ]) as get:
            docs = [d for page in client.listar_mes("2026-10") for d in page]
        self.assertEqual(100, len(docs))
        self.assertEqual(4, get.call_count)
        self.assertEqual("Primer cliente", docs[0]["historial_conversacion"][0]["contenido"])
        self.assertEqual("Segundo cliente", docs[1]["historial_conversacion"][0]["contenido"])
        self.assertTrue(all(not d["historial_conversacion"] for d in docs[2:]))
        self.assertIn("instanceIN", get.call_args.kwargs["params"]["sysparm_query"])

    def test_mensajes_se_guardan_una_sola_vez_y_se_conserva_api_anterior(self):
        doc = documento()
        historia = [{"id": "e" * 32, "contenido": "Mensaje original"}]
        doc.update(historial_conversacion=historia, traza_unificada=historia)
        with patch.object(repository, "db_execute") as execute:
            repository.guardar_documento(MagicMock(), doc, "cases", {"numero": "CS0215905", "descripcion": "Caso"}, "admin")
        snapshot = json.loads(execute.call_args.args[2][-2])
        self.assertEqual(historia, snapshot["historial_conversacion"])
        self.assertNotIn("traza_unificada", snapshot)
        self.assertIn("traza_unificada", doc)
        with patch.object(repository, "db_execute", return_value=MagicMock(fetchone=lambda: (snapshot,))):
            recuperado = repository.leer_documento(MagicMock(), doc["tabla"], "CS0215905")
        self.assertEqual(historia, recuperado["traza_unificada"])

    def test_historial_local_no_hace_llamadas_a_servicenow(self):
        expected = [{"fecha": "2026-10-10 10:00:00", "contenido": "Cliente"}]
        conn = MagicMock()
        with patch.object(app_logic, "exigir_permiso_actor"), patch.object(app_logic, "get_conn", return_value=conn), \
             patch.object(repository, "db_execute", return_value=MagicMock(fetchone=lambda: (expected,))) as execute, \
             patch.object(app_logic, "ServiceNowClient") as client:
            self.assertEqual(expected, app_logic.consultar_historial_conversacion("CS0215905", actor_email="viewer"))
        client.assert_not_called()
        self.assertNotIn("SELECT documento FROM", execute.call_args.args[1])
        conn.close.assert_called_once()

    def test_sin_permiso_no_consulta_historial(self):
        with patch.object(app_logic, "exigir_permiso_actor", side_effect=app_logic.AutorizacionError), \
             patch.object(app_logic, "get_conn") as conn:
            with self.assertRaises(app_logic.AutorizacionError):
                app_logic.consultar_historial_conversacion("CS0215905", actor_email="unknown")
        conn.assert_not_called()

    def test_actualizacion_sin_conversacion_conserva_evidencia_y_duracion(self):
        anterior = {"numero": "CS0215905", "descripcion": "Consulta del cliente",
                    "creado": "2026-10-09 08:00:00", "cerrado": "2026-10-09 10:00:00",
                    "historial_conversacion": "El cliente reportó phishing en el correo"}
        doc = {"tabla": "sn_customerservice_case", "case_id": "a" * 32,
               "campos_caso": {"number": "CS0215905"},
               "campos_raw": {"number": field("CS0215905"), "state": field("3", "Cerrado")},
               "slas": [], "historial_conversacion": []}
        fila = tuple(anterior.get(c) for c in app_logic.CASE_READ_COLUMNS)
        conn = MagicMock()
        with patch.object(app_logic, "get_conn", return_value=conn), \
             patch.object(app_logic, "_bloquear_y_validar_admin"), \
             patch.object(app_logic, "bloquear_escritura_casos"), \
             patch.object(app_logic, "numeros_existentes", return_value={"CS0215905"}), \
             patch.object(app_logic, "db_execute", return_value=MagicMock(fetchall=lambda: [fila])), \
             patch.object(app_logic, "guardar_documentos_pagina") as save:
            self.assertEqual((1, 1), app_logic._guardar_pagina_servicenow([doc], "casos", "admin"))
        row = save.call_args.args[3][0]
        self.assertNotIn("historial_conversacion", row)
        self.assertEqual("1 - phishing", row["tipificacion"])
        self.assertEqual(2, row["tiempo_respuesta"])
        self.assertEqual("Cerrado", row["estado"])
        conn.commit.assert_called_once()
        conn.close.assert_called_once()

    def test_cache_solo_se_actualiza_cuando_cambian_los_datos(self):
        source = '''
from unittest.mock import patch
import pandas as pd
import streamlit as st
import app_ui as ui
st.session_state['user'] = 'viewer'
paso = st.number_input('Ciclo', min_value=0, max_value=2, value=0)
def limpiar():
    st.session_state['limpiezas'] = st.session_state.get('limpiezas', 0) + 1
estado = {'conectado': True, 'automatica': True, 'estado': 'correcto',
          'ultima_correcta': pd.Timestamp('2026-10-10 10:00:00Z') + pd.Timedelta(minutes=int(paso)),
          'ultima_datos': pd.Timestamp('2026-10-10 10:00:00Z') + pd.Timedelta(minutes=int(paso == 2))}
with patch.object(ui, 'estado_servicenow', return_value=estado), patch.object(ui, 'limpiar_cache_casos', side_effect=limpiar):
    ui.render_estado_actualizacion_casos()
'''
        app = AppTest.from_string(source).run(timeout=15)
        self.assertEqual(1, app.session_state['limpiezas'])
        app.number_input[0].set_value(1).run(timeout=15)
        self.assertEqual(1, app.session_state['limpiezas'])
        app.number_input[0].set_value(2).run(timeout=15)
        self.assertFalse(app.exception)
        self.assertEqual(2, app.session_state['limpiezas'])

    def test_presentacion_historial_es_legible_y_sin_documento_tecnico(self):
        source = '''
from unittest.mock import patch
import streamlit as st
import app_ui as ui
st.session_state['user'] = 'viewer'
with patch.object(ui, 'consultar_historial_conversacion', return_value=[{
    'id': 'remote-test', 'fecha': '2026-10-10 10:00:00', 'tipo': 'Correo Entrante',
    'usuario': 'cliente@example.test', 'asunto': 'Error al firmar', 'contenido': 'No puedo firmar'
}]):
    ui.render_consulta_traza_servicenow('casos')
'''
        app = AppTest.from_string(source).run(timeout=15)
        app.text_input[0].set_value("CS0215905")
        app.button[0].click().run(timeout=15)
        self.assertFalse(app.exception)
        self.assertFalse(app.error, [error.value for error in app.error])
        self.assertEqual("Historial de la conversación", app.expander[0].label)
        self.assertEqual("No puedo firmar", app.dataframe[0].value.iloc[0]["Mensaje"])
        self.assertEqual("10/10/2026 05:00:00", app.dataframe[0].value.iloc[0]["Fecha"])
        self.assertNotIn("id", app.dataframe[0].value.columns)

    def test_metadatos_de_columnas_se_consultan_una_vez(self):
        cursor = MagicMock()
        cursor.description = [("numero",), ("estado",), ("creado",)]
        cursor.fetchall.return_value = [("CS0215905", "Abierto", "2026-10-10 08:00:00")]
        with patch.object(tables, "exigir_contexto_consulta"), patch.object(tables, "get_conn", return_value=MagicMock()), \
             patch.object(tables, "columnas_existentes", return_value={"numero", "estado", "creado"}) as schema, \
             patch.object(tables, "db_execute", return_value=cursor):
            df = tables.read_table_filtered("cases", columns=["numero", "estado", "creado"], equals={"estado": "Abierto"})
        self.assertEqual(1, schema.call_count)
        self.assertEqual("CS0215905", df.iloc[0]["numero"])


if __name__ == "__main__":
    unittest.main()
