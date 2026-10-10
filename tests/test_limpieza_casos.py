from unittest import TestCase
from unittest.mock import patch

import pandas as pd

import app_ui as ui
from services.casos_historial import (
    analizar_caso, agregar_analisis_historial, resumen_clasificacion_casos,
    COL_CLASIFICACION, COL_FISICO, COL_VIRTUAL, COL_MOTIVO, COL_CAUSA, COL_ESTADO_CAUSA,
    COL_ATENCION, COL_TIPOLOGIA, ENVIADO_OTROS_CANALES, SIN_CAUSA,
    COL_MOTIVO_INICIAL, COL_CLASIFICACION_INICIAL, COL_DESTINO,
    CLASIFICACIONES_PRINCIPALES, resumen_causas_historial, COL_EVIDENCIA_MOTIVO,
)
from tests.test_casos_historial import evento, DERIVACION_AGENDA, ACUSE
from tests.test_resumen_ejecutivo_casos import caso


DERIVACION_VENTAS = """Buen día,
Reciba un cordial saludo.
Entendemos que actualmente está presentando dificultades para la reactivación del servicio
y sabemos que este tipo de situaciones puede afectar la continuidad de sus actividades.
Para brindarle una atención más rápida y especializada, este tipo de requerimientos se atiende
a través de nuestro canal de ventas y será atendido directamente por uno de nuestros comerciales,
quien lo acompañará hasta solucionar el inconveniente presentado.
Tienda virtual: https://ventadigital.certicamara.com
(601) 744 2725 — opción 2
Este es el canal oficial habilitado para atender este tipo de solicitudes y garantizar un
acompañamiento personalizado durante todo el proceso.
Agradecemos su comprensión y quedamos atentos a cualquier inquietud adicional.
Cordialmente, Mesa de Ayuda Certicámara S.A."""


class LimpiezaCasosTest(TestCase):
    def test_respuesta_real_ventas_comerciales_se_cuenta_en_canales(self):
        from services.servicenow import texto_historial
        for historial in ([evento(DERIVACION_VENTAS, "Correo Saliente")],
                           texto_historial([evento(DERIVACION_VENTAS, "Correo Saliente")])):
            with self.subTest(formato=type(historial).__name__):
                row = caso("CS1", causa="", descripcion="Necesito reactivar el servicio",
                           historial_conversacion=historial)
                resultado = analizar_caso(row)
                self.assertEqual(ENVIADO_OTROS_CANALES, resultado[COL_CLASIFICACION])
                self.assertEqual(ENVIADO_OTROS_CANALES, resultado[COL_TIPOLOGIA])
                self.assertEqual("Canal de ventas, Comercial", resultado[COL_DESTINO])
                self.assertEqual("Redirigido sin atención técnica registrada", resultado[COL_ATENCION])
                self.assertEqual("Reactivación del servicio", resultado[COL_MOTIVO_INICIAL])
                self.assertEqual(SIN_CAUSA, resultado[COL_CAUSA])
                self.assertFalse(resultado[COL_FISICO])

    def test_promesas_de_acompanamiento_no_son_trabajo_tecnico_realizado(self):
        row = {"descripcion": "Necesito reactivar el servicio", "historial_conversacion": [
            evento(DERIVACION_VENTAS, "Correo Entrante", "cliente")]}
        self.assertEqual("", analizar_caso(row)[COL_DESTINO])
        row["historial_conversacion"] = [evento("Se verificó el servicio y se realizaron pruebas."),
                                         evento(DERIVACION_VENTAS, "Correo Saliente")]
        resultado = analizar_caso(row)
        self.assertNotEqual(ENVIADO_OTROS_CANALES, resultado[COL_CLASIFICACION])
        self.assertEqual("Redirigido con atención técnica registrada", resultado[COL_ATENCION])

    def test_catalogo_de_siete_grupos_concilia_dashboard_y_kpi(self):
        mensajes = ["Necesito firmar", "Instalar token físico", "Activar certificado",
                    "Token virtual no reconocido", "El sistema está caído", "Consultar facturas",
                    "Necesito reactivar el servicio"]
        filas = [caso(str(i), causa="", descripcion=m) for i, m in enumerate(mensajes)]
        filas[-1]["historial_conversacion"] = [evento(DERIVACION_VENTAS, "Correo Saliente")]
        base = agregar_analisis_historial(pd.DataFrame(filas + [filas[-1]]))
        self.assertEqual(7, len(CLASIFICACIONES_PRINCIPALES))
        tablero = resumen_clasificacion_casos(base).set_index(COL_CLASIFICACION)
        kpi, _ = ui.resumen_principales_causas_servicios_casos(base)
        kpi = kpi.set_index(COL_CLASIFICACION)
        self.assertEqual(set(CLASIFICACIONES_PRINCIPALES), set(tablero.index))
        self.assertEqual(tablero["Casos"].to_dict(), kpi["Cantidad"].to_dict())
        self.assertEqual(7, tablero["Casos"].sum())
        self.assertAlmostEqual(100, kpi["% casos"].sum())

    def test_motivo_inicial_se_conserva_despues_de_resolver_otro_sintoma(self):
        row = {"descripcion": "Necesito instalar mi token físico", "historial_conversacion": [
            {**evento("Necesito instalar mi token físico", "Correo Entrante", "cliente"), "fecha": "2026-10-01 10:00:00"},
            {**evento("Ahora el PIN está bloqueado", "Correo Entrante", "cliente"), "fecha": "2026-10-02 10:00:00"},
            evento("Se desbloqueó el PIN del token físico. Causa raíz: PIN bloqueado por intentos fallidos")],
            "causa": "PIN bloqueado por intentos fallidos"}
        for descripcion in (row["descripcion"], ""):
            resultado = analizar_caso({**row, "descripcion": descripcion})
            self.assertEqual("Instalación o configuración", resultado[COL_MOTIVO_INICIAL])
            self.assertEqual("Instalación o configuración", resultado[COL_CLASIFICACION_INICIAL])
            self.assertEqual(row["causa"], resultado[COL_CAUSA])

    def test_solicitud_clara_sin_regla_es_consulta_con_evidencia_y_no_por_revisar(self):
        solicitud = "Necesito actualizar los datos de nuestra empresa"
        resultado = analizar_caso({"descripcion": solicitud, "notas_resolucion": "Se instaló una actualización del software"})
        self.assertEqual("Solicitudes y consultas", resultado[COL_CLASIFICACION_INICIAL])
        self.assertEqual("Solicitud o consulta", resultado[COL_MOTIVO_INICIAL])
        self.assertIn(solicitud, resultado[COL_EVIDENCIA_MOTIVO])
        self.assertEqual(SIN_CAUSA, resultado[COL_CAUSA])

    def test_causas_documentadas_se_agrupan_y_conservan_explicacion_por_caso(self):
        filas = [caso(str(i), causa=f"Detalle técnico particular {i}", descripcion="Necesito instalar token físico")
                 for i in range(20)]
        filas.append(caso("AGENTE", causa="", notas_trabajo="Causa raíz: detalle técnico particular explicado por el agente"))
        base = agregar_analisis_historial(pd.DataFrame(filas))
        resumen = resumen_causas_historial(base)
        self.assertEqual(["Otras causas documentadas"], resumen["Grupo de causa"].tolist())
        self.assertEqual([21], resumen["Casos"].tolist())
        self.assertEqual(filas[0]["causa"], base.iloc[0][COL_CAUSA])
        self.assertEqual(filas[-2]["causa"], base.iloc[-2][COL_CAUSA])

    def test_agenda_usa_evidencia_actual_y_motivo_inicial_sin_duplicar_producto(self):
        filas = [
            caso("INSTALAR", causa="", descripcion="Instalar token físico", creado="2026-10-09", cuenta="Banco", canal="Web",
                 historial_conversacion=[evento(DERIVACION_AGENDA, "Correo Saliente")]),
            caso("DESCARGAR", causa="", descripcion="No puedo descargar token físico", creado="2026-10-10", cuenta="Banco", canal="Web",
                 historial_conversacion=[evento(DERIVACION_AGENDA, "Correo Saliente")]),
            caso("ANTIGUA", causa="", descripcion="Consultar factura", creado="2026-10-10", cuenta="Banco",
                 tipificacion="9 - Redireccionamiento Agenda"),
            caso("VENTAS", causa="", descripcion="Reactivar servicio", creado="2026-10-10", cuenta="Banco",
                 historial_conversacion=[evento(DERIVACION_VENTAS, "Correo Saliente")]),
        ]
        base = pd.DataFrame(filas + [filas[0]])
        agenda = ui.preparar_analisis_agendamiento(base, base, "2026-10").set_index("numero")
        self.assertEqual({"INSTALAR", "DESCARGAR"}, set(agenda.index))
        self.assertEqual("Instalación o configuración", agenda.loc["INSTALAR", ui.COL_MOTIVO_INFERIDO])
        self.assertEqual("Falla de descarga o activación", agenda.loc["DESCARGAR", ui.COL_MOTIVO_INFERIDO])
        self.assertEqual(2, agenda[COL_FISICO].sum())
        self.assertEqual("No", agenda.loc["INSTALAR", ui.COL_REINCIDENTE_AGENDA])
        self.assertEqual("Si", agenda.loc["DESCARGAR", ui.COL_REINCIDENTE_AGENDA])

    def test_operacion_agenda_sin_historico_muestra_motivo_y_token_separados(self):
        from streamlit.testing.v1 import AppTest
        app = AppTest.from_string('''
from unittest.mock import patch
import pandas as pd
import streamlit as st
import app_ui as ui
from tests.test_limpieza_casos import DERIVACION_VENTAS
from tests.test_casos_historial import evento, DERIVACION_AGENDA
from tests.test_resumen_ejecutivo_casos import caso
base = pd.DataFrame([
    caso("INSTALAR", causa="", descripcion="Instalar token físico", creado="2026-10-09", cuenta="Banco", canal="Web",
         historial_conversacion=[evento(DERIVACION_AGENDA, "Correo Saliente")]),
    caso("VENTAS", causa="", descripcion="Reactivar servicio", creado="2026-10-10", cuenta="Banco",
         historial_conversacion=[evento(DERIVACION_VENTAS, "Correo Saliente")]),
])
with patch("services.servicenow.ServiceNowClient", side_effect=AssertionError("Sin consulta a ServiceNow")):
    ui.render_analisis_agendamiento_mesa(base, pd.DataFrame(), "2026-10")
'''.encode("ascii", errors="backslashreplace").decode("ascii")).run()
        self.assertFalse(app.exception)
        detalle = app.dataframe[-1].value
        self.assertEqual(["INSTALAR"], detalle["numero"].tolist())
        self.assertEqual(["Instalación o configuración"], detalle[COL_MOTIVO_INICIAL].tolist())
        self.assertEqual(["Token físico"], detalle[ui.COL_TOKEN].tolist())

    def test_siete_grupos_se_conservan_en_resumen_ejecutivo(self):
        from services.resumen_ejecutivo_casos import construir_resumen_ejecutivo
        from components.resumen_ejecutivo_casos import lamina_resumen_casos
        from datetime import date
        mensajes = ["Necesito firmar", "Instalar token físico", "Activar certificado",
                    "Token virtual no reconocido", "El sistema está caído", "Consultar facturas",
                    "Necesito reactivar el servicio"]
        filas = [caso(str(i), causa="", descripcion=m, creado="2026-10-09") for i, m in enumerate(mensajes)]
        filas[-1]["historial_conversacion"] = [evento(DERIVACION_VENTAS, "Correo Saliente")]
        reporte = construir_resumen_ejecutivo(pd.DataFrame(), pd.DataFrame(filas), date(2026, 10, 10), ui.inferir_causa_comun_caso)
        self.assertEqual(7, len(reporte["causas_lamina"]))
        self.assertEqual(7, reporte["causas_lamina"]["Casos"].sum())
        html = lamina_resumen_casos(reporte)
        for grupo in CLASIFICACIONES_PRINCIPALES:
            self.assertIn(f'title="{grupo}"', html)

    def test_token_fisico_exige_mencion_explicita_y_no_marca_modelos_o_usb(self):
        for mensaje in ["Instalar ePass2003", "Ayuda con SafeNet", "Descarga de token USB",
                        "Configurar Feitian", "Instalar el driver", "Token gris no funciona",
                        "Necesito instalar token", "Configurar token virtual"]:
            with self.subTest(mensaje=mensaje):
                row = {"descripcion": mensaje}
                self.assertFalse(analizar_caso(row)[COL_FISICO])
                self.assertNotEqual(ui.AGENDA_MOTIVO_TOKEN_FISICO, ui.inferir_motivo_agendamiento(row))
                foco = ui.resumen_focos_destacados_kpi_casos(pd.DataFrame([row])).set_index("Foco")
                self.assertEqual(0, foco.loc["Token fisico", "Cantidad"])
        for mensaje in ["Necesito instalar mi token físico", "No puedo descargar mi token físico",
                        "Instalación de TOKEN FISICO USB", "Ayuda con el toquen fisico"]:
            with self.subTest(mensaje=mensaje):
                self.assertTrue(analizar_caso({"descripcion": mensaje})[COL_FISICO])

    def test_caidas_y_consumos_se_agrupan_por_problema_reportado(self):
        for mensaje, categoria in [
            ("Todo nuestro sistema está caído", "Caída o indisponibilidad"),
            ("Se cayó el servicio Certihuella", "Caída o indisponibilidad"),
            ("Se callo tal servicio", "Caída o indisponibilidad"),
            ("El portal no funciona", "Caída o indisponibilidad"),
            ("Tenemos fallas en la plataforma", "Falla del servicio"),
            ("No podemos ver consumos", "Falla de visualización de consumos"),
            ("Solicito reporte de consumo", "Consulta o visualización de consumos"),
        ]:
            with self.subTest(mensaje=mensaje):
                row = {"descripcion": mensaje, "causa": "Token físico", "historial_conversacion": [
                    evento(ACUSE, "Correo Saliente")]}
                resultado = analizar_caso(row)
                self.assertEqual(categoria, resultado[COL_MOTIVO_INICIAL])
                grupo = "Solicitudes y consultas" if categoria == "Consulta o visualización de consumos" else "Fallas del servicio"
                self.assertEqual(grupo, resultado[COL_CLASIFICACION])
                self.assertEqual(grupo, ui.inferir_causa_comun_caso(row))
                self.assertEqual(SIN_CAUSA, resultado[COL_CAUSA])
                self.assertEqual("Pendiente de validar", resultado[COL_ESTADO_CAUSA])
                if grupo == "Fallas del servicio":
                    self.assertEqual("Incidentes", resultado[COL_TIPOLOGIA])

    def test_instalacion_enviada_a_agenda_se_cuenta_una_vez(self):
        row = caso("CS1", descripcion="Necesito instalar token físico", causa="", historial_conversacion=[
            evento(DERIVACION_AGENDA, "Correo Saliente")])
        base = agregar_analisis_historial(pd.DataFrame([row, row]))
        self.assertTrue(base[COL_FISICO].all())
        self.assertFalse(base[COL_VIRTUAL].any())
        self.assertTrue(base[COL_MOTIVO].eq("Instalación o configuración").all())
        self.assertTrue(base[COL_CLASIFICACION].eq(ENVIADO_OTROS_CANALES).all())
        resumen = resumen_clasificacion_casos(base)
        self.assertEqual([ENVIADO_OTROS_CANALES], resumen[COL_CLASIFICACION].tolist())
        self.assertEqual([1], resumen["Casos"].tolist())
        self.assertEqual([100], resumen["% casos"].tolist())
        tipologias = ui.resumen_tipologias_soporte_casos(base)
        self.assertEqual(1, tipologias["Cantidad"].sum())
        self.assertNotIn(ui.ENVIO_AGENDA_MANUAL_USO, tipologias[COL_TIPOLOGIA].tolist())
        kpi, metricas = ui.preparar_kpi_casos_cliente_externo(base)
        self.assertEqual(1, len(kpi))
        self.assertEqual(1, metricas["total"])

    def test_envio_con_trabajo_tecnico_no_reemplaza_el_problema_por_agenda(self):
        row = {"descripcion": "No puedo descargar mi token físico", "tipificacion": "9 - Redireccionamiento Agenda",
               "historial_conversacion": [evento("Se revisó la instalación del token físico y se realizaron pruebas"),
                                            evento(DERIVACION_AGENDA, "Correo Saliente")]}
        resultado = analizar_caso(row)
        self.assertEqual("Descarga o activación", resultado[COL_CLASIFICACION])
        self.assertEqual("Incidentes", resultado[COL_TIPOLOGIA])
        self.assertEqual("Redirigido con atención técnica registrada", resultado[COL_ATENCION])
        self.assertEqual(SIN_CAUSA, resultado[COL_CAUSA])

    def test_tipificacion_antigua_no_fabrica_redireccion_o_atencion(self):
        row = {"descripcion": "El sistema está caído", "tipificacion": "9 - Redireccionamiento Agenda"}
        base = ui.agregar_tipologia_soporte_casos(pd.DataFrame([row]))
        self.assertEqual("Fallas del servicio", base.iloc[0][COL_CLASIFICACION])
        self.assertEqual("Incidentes", base.iloc[0][COL_TIPOLOGIA])
        self.assertFalse(ui.mascara_envio_agenda_casos(base).any())
        self.assertEqual(row["tipificacion"], base.iloc[0]["tipificacion"])

    def test_causa_real_se_conserva_aunque_clasificacion_sea_instalacion_o_redireccion(self):
        row = {"descripcion": "Necesito instalar token físico", "causa": "Driver incompatible con el sistema operativo",
               "historial_conversacion": [evento(DERIVACION_AGENDA, "Correo Saliente")]}
        resultado = analizar_caso(row)
        self.assertEqual(ENVIADO_OTROS_CANALES, resultado[COL_CLASIFICACION])
        self.assertEqual(row["causa"], resultado[COL_CAUSA])
        self.assertEqual("Controlador ausente, incompatible o corrupto", resultado["Grupo de causa"])
        self.assertNotEqual(ENVIADO_OTROS_CANALES, resultado["Grupo de causa"])

    def test_un_analisis_por_caso_se_reutiliza_en_graficas_y_reportes(self):
        from services import casos_historial as servicio
        with patch.object(servicio, "analizar_caso", wraps=servicio.analizar_caso) as clasificar:
            base = ui.agregar_tipologia_soporte_casos(pd.DataFrame([
                caso("CS1", descripcion="Instalar token físico"), caso("CS2", descripcion="El sistema está caído")]))
            inicial = clasificar.call_count
            resumen_clasificacion_casos(base)
            ui.resumen_tipologias_soporte_casos(base)
            ui.resumen_principales_causas_servicios_casos(base)
            self.assertEqual(inicial, clasificar.call_count)

    def test_version_anterior_se_recalcula_con_token_fisico_estricto(self):
        base = pd.DataFrame([{"numero": "CS1", "descripcion": "Instalar ePass", COL_FISICO: True,
                              "_version_analisis_historial": "historial-causas-atencion-v3"}])
        resultado = agregar_analisis_historial(base)
        self.assertFalse(resultado.iloc[0][COL_FISICO])
        self.assertEqual("Instalación o configuración", resultado.iloc[0][COL_CLASIFICACION])

    def test_gestion_tecnica_del_servicio_previa_a_redireccion_se_conserva(self):
        resultado = analizar_caso({"descripcion": "El sistema está caído", "historial_conversacion": [
            evento("Se verificó el sistema y se realizaron pruebas en el portal."),
            evento("Se remite a comercial para renovación")]})
        self.assertEqual("Fallas del servicio", resultado[COL_CLASIFICACION])
        self.assertEqual("Redirigido con atención técnica registrada", resultado[COL_ATENCION])

    def test_comparativo_usa_un_solo_indicador_de_canales(self):
        base = pd.DataFrame([
            caso("CS1", descripcion="Instalación de token físico", causa="", historial_conversacion=[
                evento(DERIVACION_AGENDA, "Correo Saliente")]),
            caso("CS2", descripcion="El sistema está caído", causa=""),
        ])
        tabla = ui.tabla_kpi_casos_comparativo_rangos(base, [
            ("Base", "2026-09-01", "2026-09-01"), ("Comparado", "2026-09-01", "2026-09-01")])
        self.assertNotIn("Envio agenda", tabla.columns)
        self.assertEqual([1, 1], tabla[ENVIADO_OTROS_CANALES].tolist())
        variacion = ui.tabla_variacion_kpi_casos(tabla).set_index("Metrica")
        self.assertNotIn("Envio agenda", variacion.index)
        self.assertEqual(1, variacion.loc[ENVIADO_OTROS_CANALES, "Base"])
        self.assertEqual(0, variacion.loc[ENVIADO_OTROS_CANALES, "Diferencia"])
