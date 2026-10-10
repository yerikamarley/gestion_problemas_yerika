import json
import unittest
from unittest.mock import MagicMock, patch

import requests

import app_logic
from repositories.servicenow import ensure_servicenow_schema
from services.servicenow import (
    ServiceNowClient, ServiceNowError, ServiceNowAuthenticationError,
    limpiar_texto, limites_mes_utc, mapear_documento, normalizar_instancia,
)


CASE_ID = "a" * 32


def field(value, display=None):
    return {"value": value, "display_value": value if display is None else display}


def caso(tipo="casos"):
    return {"sys_id": field(CASE_ID), "number": field("CS0215905" if tipo == "casos" else "INC0000001"),
            "short_description": field("Firma digital no funciona"), "state": field("1", "Abierto"),
            "sys_created_on": field("2026-10-01 04:59:00", "01/10/2026 04:59:00"),
            "sla_due": field("2026-10-01 15:00:00"), "assigned_to": field("b" * 32, "Agente"),
            "u_campo_extra": field("codigo", "Dato adicional")}


def response(rows, total=None, status=200):
    result = MagicMock()
    result.status_code = status
    result.json.return_value = {"result": rows}
    result.headers = {} if total is None else {"X-Total-Count": str(total)}
    return result


class ServiceNowClientTest(unittest.TestCase):
    def setUp(self):
        config = {"SERVICENOW_INSTANCE": "test.service-now.com", "SERVICENOW_USER": "api-test",
                  "SERVICENOW_PASSWORD": "password-test"}
        self.config_patch = patch("services.servicenow.config_value", side_effect=lambda k, d="": config.get(k, d))
        self.config_patch.start()
        self.addCleanup(self.config_patch.stop)
        self.client = ServiceNowClient()
        self.addCleanup(self.client._session.close)

    def test_todas_las_fuentes_valores_y_orden_cronologico(self):
        journal = {"sys_id": "c" * 32, "sys_created_on": "2026-10-01 11:00:00", "element": "work_notes",
                   "value": "primera\r\nsegunda\ttercera", "sys_created_by": "agente"}
        email = {"sys_id": "d" * 32, "sys_created_on": "2026-10-01 10:00:00", "type": "received",
                 "body_text": "hola\nmundo\tfin", "origemail": "cliente@example.test"}
        sla = {"sys_id": field("e" * 32), "sla": field("f" * 32, "SLA respuesta"),
               "stage": field("in_progress", "En curso"), "has_breached": field("false"),
               "start_time": field("2026-10-01 10:00:00"), "planned_end_time": field("2026-10-01 15:00:00")}
        with patch.object(self.client._session, "get", side_effect=[
            response([caso()], 1), response([sla], 1), response([journal], 1), response([email], 1),
        ]) as get:
            doc = self.client.obtener("CS0215905")
        self.assertEqual("codigo", doc["campos_raw"]["u_campo_extra"]["value"])
        self.assertEqual("Dato adicional", doc["campos_caso"]["u_campo_extra"])
        self.assertEqual("Abierto", doc["campos_caso"]["state"])
        self.assertFalse(doc["slas"][0]["has_breached"])
        self.assertEqual("SLA respuesta", doc["slas"][0]["sla"])
        self.assertEqual(["Email", "Journal"], [r["origen"] for r in doc["traza_unificada"]])
        self.assertEqual("primera --- segunda --- tercera", doc["traza_unificada"][1]["contenido"])
        for call in get.call_args_list:
            self.assertEqual("all", call.kwargs["params"]["sysparm_display_value"])
            self.assertNotIn("sysparm_fields", call.kwargs["params"])
            self.assertEqual((10, 60), call.kwargs["timeout"])
            self.assertFalse(call.kwargs["allow_redirects"])
        self.assertIn(f"target_table=sn_customerservice_case^instance={CASE_ID}",
                      get.call_args_list[-1].kwargs["params"]["sysparm_query"])

    def test_pagina_corta_por_acl_no_trunca_la_consulta(self):
        with patch.object(self.client._session, "get", side_effect=[
            response([caso()], 201), response([], 201), response([{"sys_id": "b" * 32}], 201),
        ]) as get:
            pages = list(self.client._paginas("sn_customerservice_case", "active=true"))
        self.assertEqual(2, len(pages))
        self.assertEqual([0, 100, 200], [c.kwargs["params"]["sysparm_offset"] for c in get.call_args_list])
        self.assertTrue(all(c.kwargs["params"]["sysparm_limit"] == 100 for c in get.call_args_list))

    def test_lote_largo_de_correos_conserva_todas_las_paginas_sin_header_link(self):
        query = "target_table=sn_customerservice_case^instanceIN" + ",".join(f"{i:032x}" for i in range(100))
        rows = [{"sys_id": f"{i + 1000:032x}"} for i in range(225)]

        def consultar(url, **kwargs):
            params = kwargs["params"]
            if params.get("sysparm_suppress_pagination_header") != "true":
                rejected = response([], status=400)
                rejected.json.return_value = {"error": {"message": "Pagination not supported"}}
                return rejected
            offset = params["sysparm_offset"]
            return response(rows[offset:offset + 100], len(rows))

        with patch.object(self.client._session, "get", side_effect=consultar) as get:
            result = [row for page in self.client._paginas("sys_email", query) for row in page]
        self.assertEqual(rows, result)
        self.assertEqual([0, 100, 200], [c.kwargs["params"]["sysparm_offset"] for c in get.call_args_list])

    def test_mes_con_664_casos_conserva_numero_campos_y_relaciones(self):
        rows = []
        for i in range(664):
            row = caso()
            row.update(sys_id=field(f"{i + 1:032x}"), number=field(f"CS{i + 1000000}"))
            rows.append(row)

        def consultar(url, **kwargs):
            params = kwargs["params"]
            if url.endswith("/sn_customerservice_case"):
                offset = params["sysparm_offset"]
                return response(rows[offset:offset + 100], len(rows))
            if not params.get("sysparm_suppress_pagination_header"):
                return response([], status=400)
            return response([], 0)

        with patch.object(self.client._session, "get", side_effect=consultar) as get:
            docs = [doc for page in self.client.listar_mes("2026-10") for doc in page]
        self.assertEqual(664, len(docs))
        self.assertEqual(664, len({doc["case_id"] for doc in docs}))
        self.assertEqual("CS1000663", docs[-1]["campos_raw"]["number"]["value"])
        self.assertTrue(all("u_campo_extra" in doc["campos_raw"] for doc in docs))
        self.assertEqual(28, get.call_count)  # Siete páginas, cuatro fuentes por página.

    def test_error_table_api_reconocido_no_expone_cuerpo_remoto(self):
        rejected = response([], status=400)
        rejected.json.return_value = {"error": {
            "message": "Pagination not supported", "detail": "password-test correo-privado",
        }}
        with patch.object(self.client._session, "get", return_value=rejected):
            with self.assertRaisesRegex(ServiceNowError, "enlaces de paginación") as error:
                list(self.client._paginas("sys_email", "instanceIN" + CASE_ID))
        self.assertNotIn("password-test", str(error.exception))
        self.assertNotIn("correo-privado", str(error.exception))

    def test_lote_email_rechazado_se_divide_sin_duplicar_resultados_parciales(self):
        rows = []
        emails = []
        for i in range(4):
            raw = caso()
            raw.update(sys_id=field(f"{i + 1:032x}"), number=field(f"CS{i + 1000000}"))
            rows.append(raw)
            for j in range(2):
                emails.append({"sys_id": field(f"{i * 2 + j + 4000:032x}"),
                               "instance": raw["sys_id"], "type": field("received"),
                               "body_text": field(f"Mensaje {i}-{j}"),
                               "sys_created_on": field("2026-10-09 23:20:38")})

        def consultar(url, **kwargs):
            params = kwargs["params"]
            if not url.endswith("/sys_email"):
                return response([], 0)
            ids = params["sysparm_query"].split("instanceIN")[1].split("^")[0].split(",")
            if len(ids) > 2:
                return response(emails[:2], 101) if params["sysparm_offset"] == 0 else response([], status=400)
            return response([r for r in emails if r["instance"]["value"] in ids], len(ids) * 2)

        with patch.object(self.client._session, "get", side_effect=consultar) as get:
            docs = self.client._consolidar_pagina(rows)
        self.assertEqual(4, len(docs))
        self.assertEqual([2] * 4, [len(d["historial_conversacion"]) for d in docs])
        self.assertEqual(8, len({r["id"] for d in docs for r in d["historial_conversacion"]}))
        self.assertEqual(6, get.call_count)  # SLA y notas una vez; solo correos se reintentan.

    def test_rechazo_email_individual_se_reporta_sin_bucle(self):
        with patch.object(self.client._session, "get", return_value=response([], status=400)) as get:
            with self.assertRaisesRegex(ServiceNowError, "sys_email.*400"):
                list(self.client._paginas_referencias("sys_email", "instance", [CASE_ID], "sn_customerservice_case"))
        get.assert_called_once()

    def test_sin_total_consulta_hasta_pagina_vacia(self):
        with patch.object(self.client._session, "get", side_effect=[response([caso()]), response([])]) as get:
            self.assertEqual(1, len(list(self.client._paginas("sn_customerservice_case", "active=true"))))
        self.assertEqual(2, get.call_count)

    def test_offset_ignorado_no_crea_bucle(self):
        with patch.object(self.client._session, "get", side_effect=[response([caso()]), response([caso()])]):
            with self.assertRaisesRegex(ServiceNowError, "Paginación inconsistente"):
                list(self.client._paginas("sn_customerservice_case", "active=true"))

    def test_caso_inexistente_no_consulta_actividades(self):
        with patch.object(self.client._session, "get", return_value=response([], 0)) as get:
            self.assertIsNone(self.client.obtener("CS0215905"))
        self.assertEqual(1, get.call_count)

    def test_error_de_actividad_impide_documento_incompleto(self):
        with patch.object(self.client._session, "get", side_effect=[response([caso()], 1), response([], status=403)]):
            with self.assertRaisesRegex(ServiceNowError, "task_sla.*403"):
                self.client.obtener("CS0215905")

    def test_respuesta_invalida_y_error_http_no_exponen_secretos(self):
        for simulated in [requests.ConnectionError("https://api-test:password-test@test.service-now.com"),
                          response([], status=401), response([], status=302)]:
            with self.subTest(simulated=simulated):
                kwargs = {"side_effect": simulated} if isinstance(simulated, Exception) else {"return_value": simulated}
                with patch.object(self.client._session, "get", **kwargs), self.assertRaises(ServiceNowError) as ctx:
                    list(self.client._paginas("task_sla", f"task={CASE_ID}"))
                self.assertNotIn("password-test", str(ctx.exception))
                self.assertNotIn("test.service-now.com", str(ctx.exception))
        invalid = response([])
        invalid.json.return_value = {"result": {"error": "secreto"}}
        with patch.object(self.client._session, "get", return_value=invalid), self.assertRaises(ServiceNowError):
            list(self.client._paginas("task_sla", f"task={CASE_ID}"))

    def test_numero_y_sysid_no_permiten_inyeccion_de_filtros(self):
        with patch.object(self.client._session, "get") as get:
            with self.assertRaises(ValueError):
                self.client.obtener("CS0215905^ORactive=true")
            invalid = caso()
            invalid["sys_id"] = "id^ORactive=true"
            with self.assertRaises(ServiceNowError):
                self.client.consolidar(invalid)
        get.assert_not_called()

    def test_incidente_usa_tabla_y_filtro_de_correo_correctos(self):
        with patch.object(self.client._session, "get", side_effect=[
            response([caso("incidentes")], 1), response([], 0), response([], 0), response([], 0),
        ]) as get:
            doc = self.client.obtener("INC0000001", "incidentes")
        self.assertEqual("incident", doc["tabla"])
        self.assertTrue(get.call_args_list[0].args[0].endswith("/incident"))
        self.assertIn("target_table=incident", get.call_args_list[-1].kwargs["params"]["sysparm_query"])

    def test_instancia_ausente_o_url_arbitraria_se_rechaza(self):
        for instance in ("", "http://test.service-now.com", "localhost", "test.service-now.com/path", "user@host"):
            with patch("services.servicenow.config_value", side_effect=lambda k, d="": instance if k == "SERVICENOW_INSTANCE" else "test"):
                with self.assertRaises(ServiceNowError):
                    ServiceNowClient()

    def test_enlace_del_portal_se_convierte_en_endpoint_api(self):
        with patch("services.servicenow.config_value", return_value=""), ServiceNowClient(
            instance=" https://TEST.service-now.com/navpage.do?sysparm_stack=no ",
            user="prueba", password="test",
        ) as client:
            self.assertEqual("https://test.service-now.com/api/now/table", client._base_url)
        for invalid in ("https://user:secret@test.service-now.com/navpage.do",
                        "https://test.service-now.com:8443/navpage.do",
                        "https://test.service-now.com.evil.test/navpage.do"):
            with self.subTest(invalid=invalid), self.assertRaises(ServiceNowError):
                normalizar_instancia(invalid)

    def test_credenciales_rechazadas_requieren_reconexion(self):
        with patch.object(self.client._session, "get", return_value=response([], status=401)):
            with self.assertRaisesRegex(ServiceNowAuthenticationError, "HTTP 401"):
                self.client.validar_acceso()

    def test_permiso_denegado_se_distingue_del_inicio_de_sesion(self):
        with patch.object(self.client._session, "get", return_value=response([], status=403)):
            with self.assertRaisesRegex(ServiceNowError, "permisos de lectura") as error:
                self.client.validar_acceso()
        self.assertNotIsInstance(error.exception, ServiceNowAuthenticationError)


def documento(tipo="casos"):
    raw = caso(tipo)
    return {"case_id": CASE_ID, "tabla": "sn_customerservice_case" if tipo == "casos" else "incident",
            "campos_raw": raw, "campos_caso": {k: v["display_value"] for k, v in raw.items()},
            "slas": [], "traza_unificada": []}


class ServiceNowMappingTest(unittest.TestCase):
    def test_campos_estandar_sla_y_notas_antes_omitidos(self):
        doc = documento()
        doc["campos_raw"].update({
            "made_sla": field("true", "verdadero"),
            "comments_and_work_notes": field("Observación completa"),
            "work_notes": field("Nota original"),
        })
        row = mapear_documento(doc)
        self.assertEqual("verdadero", row["sla_creado"])
        self.assertEqual("Observación completa", row["observaciones_trabajo"])
        self.assertEqual("Nota original", row["notas_trabajo"])

    def test_campo_personalizado_vacio_no_oculta_alternativa_estandar(self):
        doc = documento()
        doc["campos_raw"].update({"u_producto": field(""), "product": field("id", "Firma digital"),
                                  "close_notes": field(""), "resolution_notes": field("Resolución"),
                                  "u_sla_creado": field(""), "made_sla": False})
        row = mapear_documento(doc)
        self.assertEqual("Firma digital", row["producto"])
        self.assertEqual("Resolución", row["notas_resolucion"])
        self.assertIs(False, row["sla_creado"])

    def test_fechas_y_limites_mensuales_en_colombia(self):
        self.assertEqual(("2026-12-01 05:00:00", "2027-01-01 05:00:00"), limites_mes_utc("2026-12"))
        for mes in ("2026-13", "2026-1", "2026-10^ORactive=true", "9999-12"):
            with self.assertRaises(ValueError):
                limites_mes_utc(mes)
        row = mapear_documento(documento())
        self.assertEqual("2026-09-30 23:59:00", row["creado"])
        self.assertEqual("2026-10-01 10:00:00", row["fecha_vencimiento_sla"])
        self.assertEqual("Abierto", row["estado"])
        self.assertEqual("Agente", row["asignado"])
        self.assertNotIn("producto", row)

    def test_mapa_de_campos_personalizados(self):
        with patch("services.servicenow.config_value", return_value='{"producto":"u_campo_extra"}'):
            self.assertEqual("Dato adicional", mapear_documento(documento())["producto"])
        for invalid in ("[]", "error", '{"producto":[]}'):
            with patch("services.servicenow.config_value", return_value=invalid), self.assertRaises(ServiceNowError):
                mapear_documento(documento())

    def test_limpieza_de_tabulaciones_y_lineas(self):
        self.assertEqual("a --- b --- c --- d --- e", limpiar_texto("a\r\nb\nc\rd\te"))


class ServiceNowPersistenceTest(unittest.TestCase):
    def test_json_no_se_expone_por_rest_publico_de_supabase(self):
        with patch("repositories.servicenow.db_execute") as execute:
            ensure_servicenow_schema(MagicMock())
        self.assertIn("ENABLE ROW LEVEL SECURITY", execute.call_args_list[-1].args[1])

    def test_permiso_rechazado_antes_de_consultar_servicenow(self):
        with patch.object(app_logic, "exigir_permiso_actor", side_effect=app_logic.AutorizacionError), \
             patch.object(app_logic, "ServiceNowClient") as client:
            with self.assertRaises(app_logic.AutorizacionError):
                app_logic.sincronizar_servicenow(actor_email="viewer", mes="2026-10")
        client.assert_not_called()

    def test_fila_y_json_comparten_commit_y_rollback(self):
        for fail in (False, True):
            with self.subTest(fail=fail):
                conn = MagicMock()
                with patch.object(app_logic, "get_conn", return_value=conn), \
                     patch.object(app_logic, "_bloquear_y_validar_admin"), \
                     patch.object(app_logic, "bloquear_escritura_casos"), \
                     patch.object(app_logic, "numeros_existentes", return_value={"CS0215905"}), \
                     patch("repositories.servicenow.db_execute", side_effect=[None, RuntimeError("fallo JSON")] if fail else None) as execute:
                    if fail:
                        with self.assertRaises(RuntimeError):
                            app_logic._guardar_pagina_servicenow([documento()], "casos", "ADMIN@EXAMPLE.TEST")
                        conn.commit.assert_not_called()
                        conn.rollback.assert_called_once()
                    else:
                        self.assertEqual((1, 1), app_logic._guardar_pagina_servicenow([documento()], "casos", "ADMIN@EXAMPLE.TEST"))
                        conn.commit.assert_called_once()
                        params = execute.call_args_list[-1].args[2]
                        self.assertEqual("admin@example.test", params[-1])
                        self.assertEqual("codigo", json.loads(params[-2])["campos_raw"]["u_campo_extra"]["value"])
                    self.assertIn("INSERT INTO cases", execute.call_args_list[0].args[1])
                    self.assertIn("ON CONFLICT", execute.call_args_list[-1].args[1])
                    conn.close.assert_called_once()

    def test_admin_revocado_durante_consulta_no_guarda(self):
        conn = MagicMock()
        with patch.object(app_logic, "get_conn", return_value=conn), \
             patch.object(app_logic, "_bloquear_y_validar_admin", side_effect=app_logic.AutorizacionError), \
             patch.object(app_logic, "guardar_documento") as save:
            with self.assertRaises(app_logic.AutorizacionError):
                app_logic._guardar_pagina_servicenow([documento()], "casos", "admin")
        save.assert_not_called()
        conn.rollback.assert_called_once()
        conn.close.assert_called_once()

    def test_fallo_fuente_remota_no_invoca_persistencia(self):
        with patch.object(app_logic, "exigir_permiso_actor"), \
             patch.object(app_logic, "ServiceNowClient") as client, \
             patch.object(app_logic, "_guardar_pagina_servicenow") as save:
            client.return_value.__enter__.return_value.listar_mes.side_effect = ServiceNowError("fallo journal")
            with self.assertRaises(ServiceNowError):
                app_logic.sincronizar_servicenow(actor_email="admin", mes="2026-10")
        save.assert_not_called()

    def test_columna_no_admitida_no_llega_a_sql(self):
        conn = MagicMock()
        with patch.object(app_logic, "get_conn", return_value=conn), \
             patch.object(app_logic, "_bloquear_y_validar_admin"), \
             patch.object(app_logic, "bloquear_escritura_casos"), \
             patch.object(app_logic, "numeros_existentes", return_value=set()), \
             patch("services.servicenow.config_value", return_value='{"hack":"u_campo_extra"}'), \
             patch.object(app_logic, "guardar_documento") as save:
            with self.assertRaises(ValueError):
                app_logic._guardar_pagina_servicenow([documento()], "casos", "admin")
        save.assert_not_called()


if __name__ == "__main__":
    unittest.main()
