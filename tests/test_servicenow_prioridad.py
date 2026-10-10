"""La consulta manual puede continuar aunque el trabajador automático tenga el turno."""

from datetime import date, datetime, timezone
from threading import Event, RLock, Thread
from unittest import TestCase
from unittest.mock import MagicMock, patch

import app_logic
from services.servicenow import ServiceNowClient, ServiceNowError
from tests.test_servicenow import caso, response


class PrioridadManualTest(TestCase):
    def test_cliente_real_propaga_cesion_antes_y_despues_de_una_peticion_sin_guardar_pagina_incompleta(self):
        config = {"SERVICENOW_INSTANCE": "test.service-now.com", "SERVICENOW_USER": "test",
                  "SERVICENOW_PASSWORD": "test"}
        for llamada in (1, 2):
            with self.subTest(llamada=llamada), \
                 patch("services.servicenow.config_value", side_effect=lambda k, d="": config.get(k, d)):
                contador = 0

                def comprobar(*args):
                    nonlocal contador
                    contador += 1
                    if contador == llamada:
                        raise app_logic._TurnoManualSolicitado

                recibidas = []
                with ServiceNowClient(progress_callback=comprobar) as cliente, \
                     patch.object(cliente._session, "get", return_value=response([caso()], 1)) as consultar:
                    with self.assertRaises(app_logic._TurnoManualSolicitado):
                        recibidas.extend(cliente._paginas("sn_customerservice_case", "active=true"))
                self.assertEqual([], recibidas)
                self.assertEqual(llamada - 1, consultar.call_count)

    def test_automatico_cede_y_manual_sincroniza_en_la_misma_solicitud(self):
        # Dos conexiones independientes representan dos procesos del servidor.
        # Los bloqueos pertenecen a la transacción, no al objeto Python global.
        mutex = RLock()
        titulares = {}
        conexiones = []
        solicitud = Event()
        primera_guardada = Event()
        paginas_guardadas = []
        resultados = {}
        errores = []
        estado = []

        class Conexion:
            cerrada = False

            def __init__(self):
                self.rollbacks = 0
                conexiones.append(self)

            def rollback(self):
                self.rollbacks += 1
                with mutex:
                    for clave in list(titulares):
                        if titulares[clave] is self:
                            del titulares[clave]

            def close(self):
                self.cerrada = True

            def commit(self):
                raise AssertionError("El coordinador no debe confirmar antes de terminar")

        def ejecutar(conn, sql, params=()):
            with mutex:
                if "pg_try_advisory_xact_lock" in sql:
                    clave = params[0]
                    concedido = titulares.get(clave, conn) is conn
                    if concedido:
                        titulares[clave] = conn
                        if "solicitud_manual" in clave:
                            solicitud.set()
                    return MagicMock(fetchone=lambda: (concedido,))
                if "FROM pg_locks" in sql:
                    pendiente = "servicenow:casos:solicitud_manual:v1" in titulares
                    return MagicMock(fetchone=lambda: (pendiente,))
                return MagicMock()

        class Cliente:
            total_registros = 1

            def __init__(self, *, progress_callback):
                self.informar = progress_callback

            def __enter__(self):
                return self

            def __exit__(self, *args):
                pass

            def listar_casos_actualizados(self, *args):
                self.informar(0, "Consulta automática")
                yield [{"origen": "automatico"}]
                primera_guardada.set()
                if not solicitud.wait(3):
                    raise AssertionError("No se recibió la consulta manual")
                self.informar(0, "Siguiente consulta automática")
                yield [{"origen": "automatico_no_debe_guardarse"}]

            def listar_rango(self, *args):
                self.informar(0, "Consulta del período manual")
                yield [{"origen": "manual"}]

        def guardar(documentos, *_):
            paginas_guardadas.append(documentos[0]["origen"])
            return 1, 1

        def tarea_automatica():
            try:
                resultados["auto"] = app_logic.sincronizar_casos_automaticamente(actor_email="admin")
            except Exception as error:
                errores.append(error)

        mensajes = []
        with patch.object(app_logic, "exigir_permiso_actor"), \
             patch.object(app_logic, "get_conn", side_effect=Conexion), \
             patch.object(app_logic, "db_execute", side_effect=ejecutar), \
             patch.object(app_logic, "_consultar_estado_ingesta", return_value={
                 "ultima_correcta": datetime(2026, 10, 10, 5, tzinfo=timezone.utc)}), \
             patch.object(app_logic, "_registrar_estado_ingesta", side_effect=lambda **kw: estado.append(kw)), \
             patch.object(app_logic, "ServiceNowClient", Cliente), \
             patch.object(app_logic, "_guardar_pagina_servicenow", side_effect=guardar), \
             patch.object(app_logic, "sleep", side_effect=lambda _: Event().wait(0.005)):
            worker = Thread(target=tarea_automatica)
            worker.start()
            try:
                self.assertTrue(primera_guardada.wait(3))
                resultados["manual"] = app_logic.sincronizar_servicenow(
                    actor_email="admin", fecha_inicio=date(2026, 10, 9), fecha_fin=date(2026, 10, 10),
                    progress_callback=lambda valor, texto: mensajes.append((valor, texto)),
                )
            finally:
                solicitud.set()
                worker.join(3)
        self.assertFalse(worker.is_alive())
        self.assertFalse(errores)
        self.assertEqual(["automatico", "manual"], paginas_guardadas)
        self.assertTrue(resultados["auto"]["cedida"])
        self.assertEqual(1, resultados["auto"]["cargados"])
        self.assertEqual(1, resultados["manual"]["cargados"])
        self.assertNotIn("ocupado", resultados["manual"])
        self.assertTrue(all(s.get("correcta") is None for s in estado))
        self.assertEqual("correcto", estado[-1]["estado"])
        self.assertFalse(estado[-1]["automatica"])
        self.assertEqual(1, mensajes[-1][0])
        self.assertFalse(titulares)
        self.assertTrue(all(c.cerrada and c.rollbacks == 1 for c in conexiones))

    def test_pedido_manual_pendiente_impide_arrancar_otro_trabajador(self):
        conn = MagicMock()
        with patch.object(app_logic, "exigir_permiso_actor"), \
             patch.object(app_logic, "get_conn", return_value=conn), \
             patch.object(app_logic, "_tomar_bloqueo_sincronizacion", return_value=True), \
             patch.object(app_logic, "_hay_solicitud_manual", return_value=True), \
             patch.object(app_logic, "ServiceNowClient") as cliente, \
             patch.object(app_logic, "_registrar_estado_ingesta") as estado:
            resultado = app_logic.sincronizar_casos_automaticamente(actor_email="admin")
        self.assertTrue(resultado["cedida"])
        cliente.assert_not_called()
        estado.assert_not_called()
        conn.rollback.assert_called_once()
        conn.close.assert_called_once()

    def test_espera_limitada_informa_y_libera_solicitud_si_el_otro_proceso_no_responde(self):
        conn = MagicMock()
        mensajes = []
        with patch.object(app_logic, "exigir_permiso_actor"), \
             patch.object(app_logic, "get_conn", return_value=conn), \
             patch.object(app_logic, "_tomar_bloqueo_sincronizacion", return_value=False), \
             patch.object(app_logic, "db_execute", return_value=MagicMock(fetchone=lambda: (True,))), \
             patch.object(app_logic, "monotonic", side_effect=[0, 0, 91]), \
             patch.object(app_logic, "sleep"), \
             patch.object(app_logic, "ServiceNowClient") as cliente, \
             patch.object(app_logic, "_registrar_estado_ingesta") as estado:
            with self.assertRaisesRegex(ServiceNowError, "90 segundos"):
                app_logic.sincronizar_servicenow(actor_email="admin", fecha_inicio=date.today(), fecha_fin=date.today(),
                    progress_callback=lambda valor, texto: mensajes.append(texto))
        self.assertIn("prioridad", mensajes[0])
        cliente.assert_not_called()
        estado.assert_not_called()
        conn.rollback.assert_called_once()
        conn.close.assert_called_once()

    def test_error_manual_no_deja_reservado_el_turno(self):
        conn = MagicMock()
        with patch.object(app_logic, "exigir_permiso_actor"), \
             patch.object(app_logic, "get_conn", return_value=conn), \
             patch.object(app_logic, "_tomar_bloqueo_sincronizacion", side_effect=[False, True]), \
             patch.object(app_logic, "db_execute", return_value=MagicMock(fetchone=lambda: (True,))), \
             patch.object(app_logic, "ServiceNowClient") as cliente, \
             patch.object(app_logic, "_registrar_estado_ingesta") as estado:
            cliente.return_value.__enter__.return_value.listar_rango.side_effect = ServiceNowError("HTTP 400")
            with self.assertRaisesRegex(ServiceNowError, "HTTP 400"):
                app_logic.sincronizar_servicenow(actor_email="admin", fecha_inicio=date.today(), fecha_fin=date.today())
        self.assertEqual("error", estado.call_args.kwargs["estado"])
        conn.rollback.assert_called_once()
        conn.close.assert_called_once()
