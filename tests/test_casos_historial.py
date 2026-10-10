from unittest import TestCase
from unittest.mock import MagicMock, patch

import pandas as pd

import app_logic
import app_ui
from repositories.servicenow import leer_evidencia_casos
from services.casos_historial import (
    analizar_caso, agregar_analisis_historial, analizar_evidencias_completas,
    eventos_historial, resumen_causas_historial, texto_util_caso,
    COL_TOKEN, COL_FISICO, COL_VIRTUAL, COL_MOTIVO, COL_CAUSA, COL_ESTADO_CAUSA,
    COL_ATENCION, COL_DESTINO, COL_EVIDENCIA_TOKEN, COL_EVIDENCIA_MOTIVO,
    COL_EVIDENCIA_CAUSA, COL_EVIDENCIA_ATENCION, REDIRIGIDO_SIN_ATENCION, SIN_ATENCION,
    ENVIADO_OTROS_CANALES,
)
from services.servicenow import texto_historial


def evento(contenido, tipo="Nota de Trabajo", usuario="orlando.mendez", fecha="2026-10-09 18:00:00"):
    return {"id": "id", "fecha": fecha, "tipo": tipo, "usuario": usuario, "asunto": "", "contenido": contenido}


ACUSE = ("Hola, en Certicámara hemos recibido tu solicitud y ya estamos trabajando en ella. "
         "Tu número de caso es CS123. Conoce nuestros canales oficiales de atención: "
         "¿Necesitas instalar token físico o token virtual? Agenda tu asesoría. "
         "Visita PQRF o canal de ventas.")

DERIVACION_AGENDA = (
    'Hola BUZONPERFORMANCEART, --- Atendiendo su reporte CS0215184 ("Primer uso firma"): --- '
    'Este canal es exclusivo para reportes de errores en certificados digitales. --- '
    'Para procesos de instalación, descarga o reinstalación de token físico USB, token virtual, '
    'Pjee o pfx, y para brindarle una mejor atención personalizada, por favor separe su cita '
    'para una sesión vía Meet en la siguiente URL: --- '
    'https://zcal.co/t/agendate/certicamara --- Gracias por comunicarse con Certicámara. --- '
    'Esperamos que el servicio haya sido de tu total agrado. Si tienes alguna duda adicional '
    'sobre este tema o consideras que falta algo, quedamos muy atentos a tus comentarios'
)


class EvidenciaTokenTest(TestCase):
    def test_instalacion_descarga_uso_y_fallas_de_token_fisico_en_historial(self):
        for contenido, motivo in [
            ("Necesito instalación de mi token físico", "Instalación o configuración"),
            ("No puedo descargar el certificado en mi token físico", "Falla de descarga o activación"),
            ("Tengo dificultades al usar mi token físico", "Uso del token"),
            ("Mi token físico no permite firmar", "Falla al firmar"),
            ("El equipo no reconoce el token físico", "Token no reconocido"),
            ("Ayuda para configurar ePass2003", "Instalación o configuración"),
            ("Necesito instalar el toquen fisico", "Instalación o configuración"),
        ]:
            with self.subTest(contenido=contenido):
                row = {"historial_conversacion": texto_historial([evento(contenido, "Correo Entrante", "cliente")])}
                resultado = analizar_caso(row)
                self.assertTrue(resultado[COL_FISICO])
                self.assertFalse(resultado[COL_VIRTUAL])
                self.assertEqual(motivo, resultado[COL_MOTIVO])
                self.assertIn("Historial de la conversación", resultado[COL_EVIDENCIA_TOKEN])
                self.assertEqual(SIN_ATENCION, resultado[COL_ATENCION])
                self.assertEqual("Pendiente de validar", resultado[COL_ESTADO_CAUSA])

    def test_token_virtual_no_suma_fisico_y_otp_solo_no_confirma_token(self):
        for contenido in ["Error al usar token virtual", "Necesito activar Certitoken", "Token móvil no funciona"]:
            resultado = analizar_caso({"descripcion": contenido})
            self.assertTrue(resultado[COL_VIRTUAL])
            self.assertFalse(resultado[COL_FISICO])
        for contenido in ["El portal rechaza el OTP", "Error del driver de impresora", "No puedo descargar el reporte", "Necesito ayuda con el token"]:
            resultado = analizar_caso({"descripcion": contenido})
            self.assertFalse(resultado[COL_FISICO])
            self.assertFalse(resultado[COL_VIRTUAL])

    def test_pie_y_correo_automatico_no_inflan_token_agenda_ni_atencion(self):
        row = {"descripcion": "Consulta sobre factura", "historial_conversacion": texto_historial([
            evento(ACUSE, "Correo Saliente", "system")])}
        resultado = analizar_caso(row)
        self.assertFalse(resultado[COL_FISICO])
        self.assertFalse(resultado[COL_VIRTUAL])
        self.assertEqual("", resultado[COL_DESTINO])
        self.assertEqual(SIN_ATENCION, resultado[COL_ATENCION])
        self.assertNotIn("agenda", texto_util_caso(row))
        self.assertNotIn("Agenda", app_logic.tipificar_caso(row))
        focos = app_ui.resumen_focos_destacados_kpi_casos(pd.DataFrame([row])).set_index("Foco")
        self.assertEqual(0, focos.loc["Token fisico", "Cantidad"])
        self.assertEqual(0, focos.loc["Token virtual", "Cantidad"])
        self.assertEqual(0, focos.loc["Envio agenda", "Cantidad"])

    def test_mensaje_real_con_pie_se_analiza_y_cita_original_queda_excluida(self):
        row = {"historial_conversacion": [evento(
            "Mi token físico no permite firmar. --- De: mesa de ayuda --- " + ACUSE,
            "Correo Entrante", "cliente")]}
        resultado = analizar_caso(row)
        self.assertTrue(resultado[COL_FISICO])
        self.assertFalse(resultado[COL_VIRTUAL])
        self.assertEqual("Falla al firmar", resultado[COL_MOTIVO])

    def test_negacion_no_cuenta_dispositivo_que_cliente_no_utiliza(self):
        resultado = analizar_caso({"descripcion": "No utiliza token físico, tiene un token virtual"})
        self.assertFalse(resultado[COL_FISICO])
        self.assertTrue(resultado[COL_VIRTUAL])

    def test_ambos_tokens_y_eventos_flat_conservan_autor_fecha_y_texto(self):
        eventos = [evento("Necesito instalación del token físico", "Correo Entrante", "cliente"),
                   evento("También necesito activar mi token virtual", "Correo Entrante", "cliente")]
        recuperados = eventos_historial(texto_historial(eventos))
        self.assertEqual(2, len(recuperados))
        self.assertEqual("cliente", recuperados[0]["usuario"])
        self.assertEqual(eventos[1]["contenido"], recuperados[1]["contenido"])
        a = analizar_caso({"historial_conversacion": eventos})
        b = analizar_caso({"historial_conversacion": texto_historial(eventos)})
        self.assertEqual("Físico y virtual", a[COL_TOKEN])
        self.assertEqual(a, b)


class AtencionCausaTest(TestCase):
    def test_respuesta_real_separe_cita_es_redireccion_sin_atencion_ni_producto_confirmado(self):
        for historial in ([evento(DERIVACION_AGENDA, "Correo Saliente")],
                          texto_historial([evento(DERIVACION_AGENDA, "Correo Saliente")])):
            with self.subTest(formato=type(historial).__name__):
                row = {"descripcion": "Primer uso firma", "historial_conversacion": historial}
                resultado = analizar_caso(row)
                self.assertEqual("Agenda", resultado[COL_DESTINO])
                self.assertEqual(REDIRIGIDO_SIN_ATENCION, resultado[COL_ATENCION])
                self.assertIn("separe su cita", resultado[COL_EVIDENCIA_ATENCION])
                self.assertFalse(resultado[COL_FISICO])
                self.assertFalse(resultado[COL_VIRTUAL])
                self.assertEqual("Sin motivo específico", resultado[COL_MOTIVO])
                self.assertEqual("Sin causa documentada", resultado[COL_CAUSA])
                self.assertNotIn("token fisico", texto_util_caso(row))
                focos = app_ui.resumen_focos_destacados_kpi_casos(pd.DataFrame([row])).set_index("Foco")
                self.assertEqual(1, focos.loc["Envio agenda", "Cantidad"])
                self.assertEqual(0, focos.loc["Token fisico", "Cantidad"])
                self.assertEqual(0, focos.loc["Token virtual", "Cantidad"])

    def test_catalogo_de_derivacion_no_oculta_producto_reportado_ni_atencion_previa(self):
        resultado = analizar_caso({"descripcion": "No puedo descargar mi token físico",
                                  "historial_conversacion": [
                                      evento("Se revisó el token físico en sesión remota."),
                                      evento(DERIVACION_AGENDA, "Correo Saliente")]})
        self.assertTrue(resultado[COL_FISICO])
        self.assertFalse(resultado[COL_VIRTUAL])
        self.assertEqual("Falla de descarga o activación", resultado[COL_MOTIVO])
        self.assertEqual("Redirigido con atención técnica registrada", resultado[COL_ATENCION])

    def test_catalogo_de_productos_en_solicitud_cliente_se_conserva(self):
        resultado = analizar_caso({"historial_conversacion": [
            evento("Para procesos de instalación de token físico y token virtual, por favor "
                   "ayúdenme con mis dos dispositivos.", "Correo Entrante", "cliente")]})
        self.assertEqual("Físico y virtual", resultado[COL_TOKEN])
        self.assertEqual("Instalación o configuración", resultado[COL_MOTIVO])
        self.assertEqual(SIN_ATENCION, resultado[COL_ATENCION])

    def test_redirecciones_a_los_cuatro_destinos_sin_atencion_tecnica(self):
        for mensaje, destino in [
            ("Se redirecciona al canal de ventas para realizar la compra", "Canal de ventas"),
            ("Se remite a PQRS para registrar la reclamación", "PQRS"),
            ("Se envió enlace de agenda para que agende su cita", "Agenda"),
            ("Caso remitido a comercial para cotización", "Comercial"),
        ]:
            with self.subTest(destino=destino):
                resultado = analizar_caso({"historial_conversacion": [evento(mensaje)]})
                self.assertEqual(REDIRIGIDO_SIN_ATENCION, resultado[COL_ATENCION])
                self.assertEqual(destino, resultado[COL_DESTINO])
                self.assertIn(mensaje, resultado[COL_EVIDENCIA_ATENCION])

    def test_trabajo_tecnico_antes_de_redireccion_no_se_marca_sin_atencion(self):
        resultado = analizar_caso({"historial_conversacion": [
            evento("Se revisó la instalación del token físico y se realizaron pruebas."),
            evento("Se remite a comercial por la renovación del certificado.")]})
        self.assertEqual("Redirigido con atención técnica registrada", resultado[COL_ATENCION])
        self.assertEqual("Comercial", resultado[COL_DESTINO])
        self.assertIn("Se revisó", resultado[COL_EVIDENCIA_ATENCION])

    def test_pedir_ayuda_o_redireccion_cliente_no_demuestra_trabajo_de_soporte(self):
        for mensaje in ["Necesito que se instale el driver", "Se revisó el token en mi banco", "Se remite a comercial"]:
            resultado = analizar_caso({"historial_conversacion": [evento(mensaje, "Correo Entrante", "cliente")]})
            self.assertEqual(SIN_ATENCION, resultado[COL_ATENCION])
            self.assertEqual("", resultado[COL_DESTINO])

    def test_gestion_administrativa_o_agenda_no_se_cuenta_como_trabajo_tecnico(self):
        for mensaje in ["Se realizó la redirección a comercial", "Se revisó la solicitud de compra y se remite a ventas", "Se realizó el envío del enlace de agenda", "Se revisó la solicitud de compra de token físico y se remite a ventas"]:
            resultado = analizar_caso({"historial_conversacion": [evento(mensaje)]})
            self.assertEqual(REDIRIGIDO_SIN_ATENCION, resultado[COL_ATENCION])

    def test_causa_no_identificada_no_se_convierte_en_confirmada(self):
        resultado = analizar_caso({"notas_trabajo": "No se identificó que el driver estuviera corrupto; pendiente de investigación."})
        self.assertEqual("Pendiente de validar", resultado[COL_ESTADO_CAUSA])

    def test_problema_de_descarga_pesa_antes_que_instrucciones_de_pin_en_guia(self):
        resultado = analizar_caso({"descripcion": "No puedo descargar mi certificado en token físico",
                                  "historial_conversacion": [evento("Se comparte el manual: instalación y PIN del token físico.", "Correo Saliente")]})
        self.assertEqual("Falla de descarga o activación", resultado[COL_MOTIVO])

    def test_nombre_producto_y_accion_exitosa_no_se_inventan_como_causa_raiz(self):
        for causa in ["Token físico", "Token físico / ePass", "Firma digital", "Pendiente de revisión"]:
            resultado = analizar_caso({"causa": causa, "historial_conversacion": [evento("Se instaló el controlador y funciona.")]})
            self.assertEqual("Pendiente de validar", resultado[COL_ESTADO_CAUSA])
            self.assertEqual("Sin causa documentada", resultado[COL_CAUSA])
            self.assertEqual("Atención técnica registrada", resultado[COL_ATENCION])

    def test_causa_explicada_en_nota_interna_con_fragmento_y_grupo(self):
        mensaje = "Se revisó el token físico. Causa raíz: el driver es incompatible con el sistema operativo."
        resultado = analizar_caso({"historial_conversacion": [evento(mensaje)]})
        self.assertEqual("Explicada en la atención", resultado[COL_ESTADO_CAUSA])
        self.assertEqual("Controlador ausente, incompatible o corrupto", resultado["Grupo de causa"])
        self.assertIn("driver es incompatible", resultado[COL_CAUSA])
        self.assertIn("orlando.mendez", resultado[COL_EVIDENCIA_CAUSA])
        self.assertIn("Causa raíz:", resultado[COL_EVIDENCIA_CAUSA])

    def test_causa_registrada_prevalece_y_cliente_no_confirma_causa_tecnica(self):
        eventos = [evento("Se debe a incompatibilidad de Java", "Correo Entrante", "cliente")]
        resultado = analizar_caso({"historial_conversacion": eventos})
        self.assertEqual("Pendiente de validar", resultado[COL_ESTADO_CAUSA])
        resultado = analizar_caso({"causa": "Certificado vencido", "historial_conversacion": eventos})
        self.assertEqual("Registrada en el caso", resultado[COL_ESTADO_CAUSA])
        self.assertEqual("Certificado vencido", resultado[COL_CAUSA])

    def test_no_confunde_menciones_de_destino_con_redireccion_ni_frase_negada(self):
        for mensaje in ["El cliente compró su token en ventas", "No se remitió a comercial; se revisó el token", "Se realizó una prueba; no se redireccionó a agenda"]:
            resultado = analizar_caso({"historial_conversacion": [evento(mensaje)]})
            self.assertEqual("", resultado[COL_DESTINO])

    def test_evidencia_larga_conserva_fragmento_original_cerca_del_hallazgo(self):
        original = "Párrafo    repetido \n" * 1000 + "El token físico no permite firmar."
        resultado = analizar_caso({"descripcion": original})
        self.assertIn("token físico", resultado[COL_EVIDENCIA_TOKEN])
        self.assertIn("no permite firmar", resultado[COL_EVIDENCIA_MOTIVO])
        self.assertLess(len(resultado[COL_EVIDENCIA_TOKEN]), 500)


class IntegracionHistorialTest(TestCase):
    def test_enviado_a_otros_canales_concilia_causas_tipologias_y_historial(self):
        mensajes = [DERIVACION_AGENDA, "Se remite a PQRS", "Se envía al canal de ventas",
                    "Caso remitido a comercial"]
        causa_original = "Firma digital: certificado vencido"
        filas = [{"numero": f"CS{i}", "causa": causa_original, "descripcion": "Primer uso firma",
                  "historial_conversacion": texto_historial([evento(mensaje, "Correo Saliente")])}
                 for i, mensaje in enumerate(mensajes)]
        filas.extend([
            {"numero": "CS4", "causa": causa_original, "historial_conversacion": texto_historial([
                evento("Se revisó el token físico y se realizaron pruebas"),
                evento(DERIVACION_AGENDA, "Correo Saliente")])},
            {"numero": "CS5", "causa": causa_original, "historial_conversacion": texto_historial([
                evento(ACUSE, "Correo Saliente")])},
        ])
        base = app_ui.agregar_tipologia_soporte_casos(pd.DataFrame(filas))
        for fila in filas[:4]:
            self.assertEqual(ENVIADO_OTROS_CANALES, app_ui.inferir_causa_comun_caso(fila))
            self.assertEqual(ENVIADO_OTROS_CANALES, app_ui.clasificar_tipologia_soporte_caso(fila))
        self.assertEqual([ENVIADO_OTROS_CANALES] * 4,
                         base.iloc[:4][app_ui.TEXT_TIPOLOGIA_SOPORTE].tolist())
        self.assertTrue(base["causa"].eq(causa_original).all())
        self.assertTrue(base[COL_CAUSA].eq(causa_original).all())
        self.assertFalse(base.iloc[4:]["Grupo de causa"].eq(ENVIADO_OTROS_CANALES).any())
        tipologias = app_ui.resumen_tipologias_soporte_casos(base).set_index(app_ui.TEXT_TIPOLOGIA_SOPORTE)
        self.assertEqual(4, tipologias.loc[ENVIADO_OTROS_CANALES, app_ui.TEXT_CANTIDAD])
        self.assertEqual(6, tipologias[app_ui.TEXT_CANTIDAD].sum())
        self.assertAlmostEqual(66.67, tipologias.loc[ENVIADO_OTROS_CANALES, "% casos"], places=2)
        causas, _ = app_ui.resumen_principales_causas_servicios_casos(base)
        self.assertEqual(4, causas.set_index("Causa raíz").loc[ENVIADO_OTROS_CANALES, "Cantidad"])
        resumen = resumen_causas_historial(base).set_index("Grupo de causa")
        self.assertEqual(4, resumen.loc[ENVIADO_OTROS_CANALES, "Casos"])
        self.assertEqual("Clasificación de gestión", resumen.loc[ENVIADO_OTROS_CANALES, COL_ESTADO_CAUSA])
        self.assertEqual(6, resumen["Casos"].sum())

    def test_categoria_manual_enviado_a_otros_canales_tambien_aparece_en_tipologia(self):
        row = {"numero": "CS1", "causa": ENVIADO_OTROS_CANALES}
        self.assertEqual(ENVIADO_OTROS_CANALES, app_ui.inferir_causa_comun_caso(row))
        self.assertEqual(ENVIADO_OTROS_CANALES, app_ui.clasificar_tipologia_soporte_caso(row))

    def test_tabla_indices_vacios_duplicados_y_resumen_concilian(self):
        df = pd.DataFrame([{"numero": "CS1", "causa": "Driver incompatible"},
                           {"numero": "CS2", "descripcion": "Token virtual"},
                           {"numero": "CS2", "descripcion": "Token virtual"}], index=[8, 9, 10])
        copia = df.copy()
        resultado = agregar_analisis_historial(df)
        pd.testing.assert_frame_equal(df, copia)
        self.assertEqual([8, 9, 10], resultado.index.tolist())
        resumen = resumen_causas_historial(resultado)
        self.assertEqual(2, resumen["Casos"].sum())
        self.assertEqual(100, resumen["% casos"].sum())
        vacio = agregar_analisis_historial(pd.DataFrame())
        self.assertTrue(vacio.empty)
        self.assertIn(COL_TOKEN, vacio)
        pd.testing.assert_frame_equal(resultado, agregar_analisis_historial(resultado))

    def test_evidencia_completa_se_lee_local_en_un_lote_y_exige_contexto(self):
        conn = MagicMock()
        leer_evidencia_casos(conn, ["CS1", "CS1", "CS2"])
        sql, params = conn.execute.call_args.args
        self.assertIn("ANY(%s::text[])", sql)
        self.assertNotIn("CS1", sql)
        self.assertEqual(["CS1", "CS2"], params[0])
        leer_evidencia_casos(conn, [])
        self.assertEqual(1, conn.execute.call_count)
        with patch.object(app_logic, "get_conn") as conectar, self.assertRaises(app_logic.AutorizacionError):
            app_logic.load_evidencia_casos(["CS1"])
        conectar.assert_not_called()

    def test_enriquecer_snapshot_no_borra_excel_o_historial_y_no_duplica_filas(self):
        df = pd.DataFrame([{"numero": "CS1", "historial_conversacion": "Nota histórica"},
                           {"numero": "CS2", "notas_trabajo": "Se revisó la instalación de token físico"}], index=[20, 21])
        registros = [("CS1", {"value": "No puedo descargar el certificado en el token virtual"}, [])]
        resultado = analizar_evidencias_completas(df, registros)
        self.assertEqual([20, 21], resultado.index.tolist())
        self.assertTrue(resultado.loc[20, COL_VIRTUAL])
        self.assertTrue(resultado.loc[21, COL_FISICO])
        self.assertEqual("Nota histórica", resultado.loc[20, "historial_conversacion"])

    def test_grafica_diaria_y_focos_cuentan_token_por_historial_aunque_causa_sea_firma(self):
        df = pd.DataFrame([{"numero": "CS1", "causa": "Firma digital", "descripcion": "Instalación de token físico", "creado_dt_dashboard": pd.Timestamp("2026-10-09")},
                           {"numero": "CS2", "descripcion": "Instalación de token virtual", "creado_dt_dashboard": pd.Timestamp("2026-10-09")},
                           {"numero": "CS3", "historial_conversacion": texto_historial([evento(ACUSE, "Correo Saliente")]), "creado_dt_dashboard": pd.Timestamp("2026-10-09")}])
        df[app_ui.TEXT_CREADO_DT_DASHBOARD] = df.pop("creado_dt_dashboard")
        diario = app_ui.resumen_entradas_diarias_casos(df, True, "2026-10-09", "2026-10-10")
        self.assertEqual([1, 0], diario[app_ui.TEXT_CASOS_2].tolist())
        focos = app_ui.resumen_focos_destacados_kpi_casos(df).set_index("Foco")
        self.assertEqual(1, focos.loc["Token fisico", "Cantidad"])
        self.assertEqual(1, focos.loc["Token virtual", "Cantidad"])

    def test_interfaz_incluye_otras_areas_filtra_y_descarga_evidencia_sin_api(self):
        from streamlit.testing.v1 import AppTest
        app = AppTest.from_string('''
from unittest.mock import patch
import pandas as pd
import streamlit as st
import app_ui as ui
from tests.test_casos_historial import evento, ACUSE
from services.servicenow import texto_historial
datos = pd.DataFrame([
    {"numero": "CS1", "asignado": "Yader Neira", "descripcion": "Instalación de token físico", "creado": "2026-10-09"},
    {"numero": "CS2", "asignado": "Otra área", "descripcion": "Falla al descargar token virtual", "creado": "2026-10-09"},
    {"numero": "CS3", "asignado": "Otra área", "historial_conversacion": texto_historial([evento("Se remite a PQRS", "Correo Saliente")]), "creado": "2026-10-09"},
    {"numero": "CS4", "asignado": "Yader Neira", "historial_conversacion": texto_historial([evento(ACUSE, "Correo Saliente")]), "creado": "2026-10-09"},
])
def evidencia(*args):
    st.session_state["consultas_evidencia"] = st.session_state.get("consultas_evidencia", 0) + 1
    return [("CS4", {"value": "Instalación de token físico"}, [])]
def descarga(df):
    st.session_state["descarga_analisis"] = df
    return b"csv"
with patch.object(ui, "cargar_evidencia_casos_cache", side_effect=evidencia), \\
     patch.object(ui, "obtener_contexto_autorizacion", return_value=("admin", "dashboard_casos_soporte")), \\
     patch.object(ui, "dataframe_csv_bytes", side_effect=descarga), \\
     patch("services.servicenow.ServiceNowClient", side_effect=AssertionError("No debe consultar ServiceNow")):
    ui.render_analisis_historial_casos(datos, "Octubre 2026")
'''.encode("ascii", errors="backslashreplace").decode("ascii")).run()
        self.assertFalse(app.exception)
        detalle = app.session_state["descarga_analisis"]
        self.assertEqual({"CS1", "CS2", "CS3", "CS4"}, set(detalle["Número"]))
        self.assertEqual(0, app.session_state.filtered_state.get("consultas_evidencia", 0))
        app.selectbox(key="historial_casos_destino").set_value("PQRS").run()
        self.assertFalse(app.exception)
        self.assertEqual(["CS3"], app.session_state["descarga_analisis"]["Número"].tolist())
        app.selectbox(key="historial_casos_destino").set_value("Todos").run()
        app.button(key="revisar_evidencias_casos").click().run()
        self.assertFalse(app.exception)
        self.assertEqual(1, app.session_state["consultas_evidencia"])
        app.selectbox(key="historial_casos_token").set_value("Token físico").run()
        self.assertFalse(app.exception)
        self.assertEqual({"CS1", "CS4"}, set(app.session_state["descarga_analisis"]["Número"]))
        self.assertEqual(1, app.session_state["consultas_evidencia"])
        self.assertIn("Evidencia del token", app.session_state["descarga_analisis"])
