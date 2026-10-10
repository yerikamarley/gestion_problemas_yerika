from unittest import TestCase
from unittest.mock import patch

import pandas as pd

import app_ui as ui
from services.casos_historial import (
    analizar_caso, agregar_analisis_historial, resumen_clasificacion_casos,
    COL_CLASIFICACION, COL_FISICO, COL_VIRTUAL, COL_MOTIVO, COL_CAUSA, COL_ESTADO_CAUSA,
    COL_ATENCION, COL_TIPOLOGIA, ENVIADO_OTROS_CANALES, SIN_CAUSA,
)
from tests.test_casos_historial import evento, DERIVACION_AGENDA, ACUSE
from tests.test_resumen_ejecutivo_casos import caso


class LimpiezaCasosTest(TestCase):
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
            ("No podemos ver consumos", "Consulta o visualización de consumos"),
            ("Solicito reporte de consumo", "Consulta o visualización de consumos"),
        ]:
            with self.subTest(mensaje=mensaje):
                row = {"descripcion": mensaje, "causa": "Token físico", "historial_conversacion": [
                    evento(ACUSE, "Correo Saliente")]}
                resultado = analizar_caso(row)
                self.assertEqual(categoria, resultado[COL_CLASIFICACION])
                self.assertEqual(categoria, ui.inferir_causa_comun_caso(row))
                self.assertEqual(SIN_CAUSA, resultado[COL_CAUSA])
                self.assertEqual("Pendiente de validar", resultado[COL_ESTADO_CAUSA])
                if "consumos" not in categoria:
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
        self.assertEqual("Caída o indisponibilidad", base.iloc[0][COL_CLASIFICACION])
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
        self.assertEqual("Caída o indisponibilidad", resultado[COL_CLASIFICACION])
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
