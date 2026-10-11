from unittest import TestCase

import pandas as pd

from services.casos_historial import (
    analizar_caso, agregar_analisis_historial, resumen_clasificacion_casos,
    COL_ATENCION, COL_CAUSA, COL_CLASIFICACION, COL_DESTINO, COL_EVIDENCIA_ATENCION,
    COL_FISICO, COL_MOTIVO_INICIAL, COL_TIPOLOGIA, COL_VIRTUAL,
    ENVIADO_OTROS_CANALES, REDIRIGIDO_SIN_ATENCION, VERSION_ANALISIS,
)
from services.servicenow import texto_historial
from tests.test_casos_historial import evento, ACUSE


AGENDA = (
    "Entendemos que actualmente presentas dificultades para instalar, activar o utilizar tu firma digital. "
    "Este tipo de requerimientos lo manejamos exclusivamente a través de nuestro canal de agendamiento. "
    "Un técnico especializado te contactará para hacer el proceso contigo paso a paso. "
    "Programa tu cita oficial aquí: zcal.co/t/agendate/certicamara"
)
PRIMERA_DESCARGA = (
    "Al validar en el sistema, se verifica que cuentan con un token físico emitido el cual se encuentra "
    "en estado pendiente por descargar. Para llevar a cabo el proceso de descarga e instalación por primera vez, "
    "es necesario agendar una atención con uno de nuestros ingenieros de soporte. "
    "Les solicitamos amablemente realizar el agendamiento a través del siguiente enlace oficial: "
    "Programa tu cita oficial aquí: zcal.co/t/agendate/certicamara. "
    "Una vez programada la cita, un técnico especializado se conectará en el horario seleccionado "
    "para asistirles de manera remota en la configuración y descarga del certificado en el dispositivo. "
    "Procedemos con el cierre formal del presente caso."
)
PETICIONES = (
    "Entendemos que actualmente está presentando dificultades con su servicio. "
    "Este tipo de requerimientos se atiende a través de nuestro canal peticiones y será atendido "
    "directamente por uno de nuestros personales, quien lo acompañará hasta solucionar el inconveniente. "
    "https://pqrf.certicamara.com/registrar. Cordialmente, Mesa de Ayuda."
)
SERVICIO_CLIENTE = (
    "Al validar en el sistema, se verifica que la fecha en la cual se activó el enlace para la descarga "
    "del certificado digital fue el 24/06/2026. Los enlaces de descarga cuentan con un tiempo límite "
    "de vigencia de tres meses; al no realizarse la descarga dentro de dicho periodo, el sistema genera un bloqueo. "
    "Esta gestión no corresponde al área de Soporte Técnico. Le solicitamos amablemente comunicarse "
    "con nuestra Línea de Servicio al Cliente al (601) 744 27 25, Opción 1, donde le brindarán "
    "la información necesaria sobre el estado de su firma y el procedimiento para la habilitación o reactivación. "
    "Al haberle indicado el canal correspondiente, procedemos con el cierre formal del presente caso."
)
ENTIDAD = (
    "Este tipo de solicitudes lo debe escalar directamente la entidad bancaria que realizó el proceso "
    "de validación de identidad y generó la novedad, ya que son los únicos autorizados y consumen "
    "el Servicio de Certihuella. Para que la entidad realice el proceso lo debe realizar bajo el canal "
    "Mesadeayuda@certicamara.com; se procede con el cierre del presente PQRs."
)
AUTOGESTION = (
    "Confirmamos que su solicitud fue atendida y damos el caso por resuelto. "
    "Pueden descargar cualquiera de los Certificados solicitados en la Url: "
    "https://certititulo-certificados.web.app/ la consulta se debe hacer por cédula de ciudadanía "
    "o descargar el masivo de la Url: https://drive.google.com/drive/folders/ejemplo. "
    "Gracias por comunicarse con Certicámara."
)
COMERCIAL = (
    "Este canal es exclusivo para soporte técnico. Para temas comerciales (compras o renovaciones), "
    "por favor comunícate con nuestros canales especializados: "
    "Autogestión: https://ventadigital.certicamara.com. Tel: (601) 7442725."
)


class CanalesDocumentadosTest(TestCase):
    def test_ejemplos_del_documento_en_historial_y_notas(self):
        ejemplos = [
            (AGENDA, "Agenda"), (PRIMERA_DESCARGA, "Agenda"),
            (PETICIONES, "PQRS"), (SERVICIO_CLIENTE, "Servicio al cliente"),
            (ENTIDAD, "Entidad bancaria"), (AUTOGESTION, "Autogestión"),
            (COMERCIAL, "Comercial"),
        ]
        for mensaje, destino in ejemplos:
            for row in [
                {"historial_conversacion": [evento(mensaje, "Correo Saliente")]},
                {"historial_conversacion": texto_historial([evento(mensaje, "Correo Saliente")])},
                {"notas_trabajo": mensaje}, {"notas_resolucion": mensaje},
            ]:
                with self.subTest(destino=destino, campo=next(iter(row))):
                    analisis = analizar_caso({"descripcion": "Necesito reactivar el servicio", **row})
                    self.assertEqual(ENVIADO_OTROS_CANALES, analisis[COL_CLASIFICACION])
                    self.assertEqual(ENVIADO_OTROS_CANALES, analisis[COL_TIPOLOGIA])
                    self.assertIn(destino, analisis[COL_DESTINO].split(", "))
                    self.assertEqual(REDIRIGIDO_SIN_ATENCION, analisis[COL_ATENCION])
                    self.assertTrue(analisis[COL_EVIDENCIA_ATENCION])
                    self.assertEqual("Reactivación del servicio", analisis[COL_MOTIVO_INICIAL])

    def test_observaciones_con_autor_soporte_del_documento(self):
        for campo in ("observaciones_adicionales", "observaciones_trabajo"):
            mensaje = "6/10/2026 02:56:46 PM - Jhon Montes (Observaciones adicionales)\n" + PETICIONES
            analisis = analizar_caso({campo: mensaje, "descripcion": "Error con mi servicio"})
            self.assertEqual("PQRS", analisis[COL_DESTINO])
            self.assertEqual(ENVIADO_OTROS_CANALES, analisis[COL_CLASIFICACION])
            self.assertIn("Jhon Montes", analisis[COL_EVIDENCIA_ATENCION])
            sin_autor_soporte = analizar_caso({campo: mensaje.replace("Jhon Montes", "Cliente Ejemplo")})
            self.assertEqual("", sin_autor_soporte[COL_DESTINO])

    def test_cierre_y_causa_original_se_conservan_en_autogestion(self):
        row = {"numero": "CS1", "estado": "Cerrado", "causa": "Enlace de descarga vencido",
               "historial_conversacion": [evento(AUTOGESTION, "Correo Saliente")]}
        base = agregar_analisis_historial(pd.DataFrame([row, row]))
        self.assertEqual("Cerrado", base.iloc[0]["estado"])
        self.assertEqual(row["causa"], base.iloc[0][COL_CAUSA])
        self.assertEqual(row["causa"], base.iloc[0]["causa"])
        self.assertEqual(ENVIADO_OTROS_CANALES, base.iloc[0][COL_CLASIFICACION])
        self.assertEqual(1, resumen_clasificacion_casos(base)["Casos"].sum())

    def test_actualiza_analisis_previo_sin_sincronizar(self):
        row = {"numero": "CS1", "notas_trabajo": AGENDA,
               COL_CLASIFICACION: "Firma digital", "_version_analisis_historial": "historial-causas-detalle-v6"}
        base = agregar_analisis_historial(pd.DataFrame([row]))
        self.assertEqual(VERSION_ANALISIS, base.iloc[0]["_version_analisis_historial"])
        self.assertEqual(ENVIADO_OTROS_CANALES, base.iloc[0][COL_CLASIFICACION])

    def test_redireccion_no_borra_trabajo_tecnico_real_realizado(self):
        for mensaje in (AGENDA, PETICIONES, SERVICIO_CLIENTE, AUTOGESTION):
            analisis = analizar_caso({"descripcion": "Mi token físico no permite firmar", "historial_conversacion": [
                evento("Se configuró el token físico y se realizaron pruebas de firma."),
                evento(mensaje, "Correo Saliente")]})
            self.assertNotEqual(ENVIADO_OTROS_CANALES, analisis[COL_CLASIFICACION])
            self.assertEqual("Redirigido con atención técnica registrada", analisis[COL_ATENCION])

    def test_incoming_acuses_y_citas_no_son_redireccion_del_soporte(self):
        for mensaje in (AGENDA, PETICIONES, ENTIDAD, SERVICIO_CLIENTE, AUTOGESTION):
            for contenido, tipo, actor in [
                (mensaje, "Correo Entrante", "cliente"),
                ("Hola, consulta de estado.\nDe: Mesa de Ayuda\n" + mensaje, "Correo Saliente", "orlando.mendez"),
                (ACUSE + "\n" + mensaje, "Correo Saliente", "system"),
            ]:
                analisis = analizar_caso({"historial_conversacion": [evento(contenido, tipo, actor)]})
                self.assertEqual("", analisis[COL_DESTINO])
                self.assertNotEqual(ENVIADO_OTROS_CANALES, analisis[COL_CLASIFICACION])

    def test_menciones_sueltas_y_negaciones_no_suman_canales(self):
        for mensaje in [
            "El cliente ya llamó a Servicio al Cliente.", "La entidad bancaria consume Certihuella.",
            "Referencia: https://certititulo-certificados.web.app/", "No debe escalar a la entidad bancaria.",
            "No es necesario agendar una atención.", "No se atiende a través del canal de peticiones.",
            "No es un soporte, es una instalación.",
        ]:
            with self.subTest(mensaje=mensaje):
                analisis = analizar_caso({"notas_trabajo": mensaje})
                self.assertEqual("", analisis[COL_DESTINO])
                self.assertNotEqual(ENVIADO_OTROS_CANALES, analisis[COL_CLASIFICACION])

    def test_catalogo_generico_no_inventa_token_fisico(self):
        analisis = analizar_caso({"descripcion": "Primer uso de mi token virtual", "historial_conversacion": [
            evento("Para procesos de instalación, descarga o reinstalación de token físico USB, token virtual, "
                   "Pjee o pfx, por favor separe su cita: https://zcal.co/t/agendate/certicamara", "Correo Saliente")]})
        self.assertEqual(ENVIADO_OTROS_CANALES, analisis[COL_CLASIFICACION])
        self.assertFalse(analisis[COL_FISICO])
        self.assertTrue(analisis[COL_VIRTUAL])
