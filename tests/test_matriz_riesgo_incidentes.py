import unittest

import pandas as pd

from app_logic import (
    construir_analisis_anual_reincidencias_incidentes,
    construir_matriz_riesgo_incidentes,
)


class MatrizRiesgoIncidentesTest(unittest.TestCase):
    def test_usa_solo_incidentes_y_problemas_en_plan_de_trabajo(self):
        incidentes = pd.DataFrame([
            {"numero": "INC1", "servicio_negocio": "Portal", "tipificacion_auto": "Caída", "causa_raiz_auto": "Base de datos", "empresa": "TI", "breve_descripcion": "Portal no responde"},
            {"numero": "INC2", "servicio_negocio": "Portal", "tipificacion_auto": "Caída", "causa_raiz_auto": "Base de datos", "empresa": "TI", "breve_descripcion": "Nueva caída del portal"},
        ])
        problemas = pd.DataFrame([
            {"numero": "PRB1", "estado": "Plan de trabajo", "declaracion_problema": "Falla de base de datos del Portal"},
            {"numero": "PRB2", "estado": "Cerrado", "declaracion_problema": "Falla de base de datos del Portal"},
        ])

        matriz = construir_matriz_riesgo_incidentes(incidentes, problemas)

        self.assertEqual(1, len(matriz))
        self.assertEqual(2, matriz.iloc[0]["cantidad_incidentes"])
        self.assertEqual("PRB1", matriz.iloc[0]["problemas_plan_trabajo"])

    def test_superpone_campos_editables_guardados(self):
        incidentes = pd.DataFrame([
            {"numero": "INC1", "servicio_negocio": "Portal", "tipificacion_auto": "Caída", "causa_raiz_auto": "Red", "empresa": "TI"},
        ])
        inicial = construir_matriz_riesgo_incidentes(incidentes)
        ediciones = pd.DataFrame([{
            "id_matriz": inicial.iloc[0]["id_matriz"],
            "estado_mejora": "En ejecución",
            "causa_raiz": "Configuración de red validada",
            "mejoras": "Automatizar recuperación",
        }])

        matriz = construir_matriz_riesgo_incidentes(incidentes, ediciones_df=ediciones)

        self.assertEqual("En ejecución", matriz.iloc[0]["estado_mejora"])
        self.assertEqual("Configuración de red validada", matriz.iloc[0]["causa_raiz"])
        self.assertEqual("Automatizar recuperación", matriz.iloc[0]["mejoras"])

    def test_analisis_anual_segmenta_la_misma_causa_por_mes(self):
        incidentes = pd.DataFrame([
            {"numero": "INC1", "creado": "2026-01-05", "tipificacion_auto": "Caída", "causa_raiz_auto": "Base de datos"},
            {"numero": "INC2", "creado": "2026-01-20", "tipificacion_auto": "Caída", "causa_raiz_auto": "Base de datos"},
            {"numero": "INC3", "creado": "2026-03-02", "tipificacion_auto": "Caída", "causa_raiz_auto": "Base de datos"},
            {"numero": "INC4", "creado": "2026-02-01", "tipificacion_auto": "Caída", "causa_raiz_auto": "Red"},
        ])

        resumen, mensual = construir_analisis_anual_reincidencias_incidentes(incidentes)

        self.assertEqual(1, len(resumen))
        self.assertEqual(3, resumen.iloc[0]["total_anual"])
        self.assertEqual(2, resumen.iloc[0]["meses_con_eventos"])
        self.assertEqual([2, 1], mensual["incidentes"].tolist())

    def test_observaciones_prevalecen_sobre_servicio_ssps(self):
        incidentes = pd.DataFrame([
            {
                "numero": "INC1", "servicio_negocio": "SSPS", "creado": "2026-08-03",
                "breve_descripcion": "Novedad SSPS",
                "observaciones_trabajo": "Al validar corresponde al despliegue del nuevo canal de ventas y la BD de cuentas.",
            },
            {
                "numero": "INC2", "servicio_negocio": "SSPS", "creado": "2026-08-05",
                "descripcion": "Se ajusta tabla de homologación de cuentas bancarias del canal de ventas.",
            },
        ])

        matriz = construir_matriz_riesgo_incidentes(incidentes)

        self.assertTrue(matriz.empty, "Un despliegue o ajuste sin evidencia de falla no debe materializar un riesgo.")

    def test_clasifica_componentes_comunes_y_problemas_de_firma(self):
        incidentes = pd.DataFrame([
            {"numero": "INC1", "descripcion": "Certitoken no permite firmar el documento"},
            {"numero": "INC2", "observaciones_trabajo": "Se presenta error al firmar con Certi Token"},
            {"numero": "INC3", "descripcion": "Caída del servicio OCSP"},
            {"numero": "INC4", "descripcion": "OCSP no responde"},
        ])

        matriz = construir_matriz_riesgo_incidentes(incidentes)

        temas = set(matriz["criterio_similitud"])
        self.assertIn("Certitoken · problema de firma", temas)
        self.assertIn("OCSP · caída o indisponibilidad", temas)
        fila_token = matriz[matriz["componente_detectado"] == "Certitoken"].iloc[0]
        self.assertIn("INC1", fila_token["incidentes_asociados"])
        self.assertIn("Certitoken", fila_token["comentario_analisis"])

    def test_excluye_instalacion_y_alerta_sin_afectacion(self):
        incidentes = pd.DataFrame([
            {"numero": "INC1", "descripcion": "Solicitud de instalación y configuración de middleware en nuevo PC"},
            {"numero": "INC2", "descripcion": "Alerta NOC normalizada sin afectación del servicio"},
        ])
        self.assertTrue(construir_matriz_riesgo_incidentes(incidentes).empty)

    def test_separa_rpost_como_proveedor_y_noc_como_alertamiento(self):
        incidentes = pd.DataFrame([
            {"numero": "INC1", "descripcion": "Portal RPOST caído y no responde"},
            {"numero": "INC2", "descripcion": "Monitoreo NOC no generó alerta durante la caída"},
        ])
        matriz = construir_matriz_riesgo_incidentes(incidentes)
        self.assertEqual({"R140", "R-MON"}, set(matriz["id_riesgo"]))
        self.assertEqual("Proveedor externo", matriz[matriz["id_riesgo"] == "R140"].iloc[0]["dominio_evento"])

    def test_prioriza_servicio_sobre_canal_noc(self):
        incidentes = pd.DataFrame([
            {"numero": "INC1", "descripcion": "Alerta NOC: RPOST caído y no responde"},
            {"numero": "INC2", "descripcion": "Alerta de SSPS relacionada con problema de firma"},
            {"numero": "INC3", "descripcion": "Monitoreo reporta caída de OCSP"},
        ])
        matriz = construir_matriz_riesgo_incidentes(incidentes)
        componentes = set(matriz["componente_detectado"])
        self.assertIn("RPOST", componentes)
        self.assertIn("SSPS", componentes)
        self.assertIn("OCSP", componentes)
        self.assertNotIn("Monitoreo / NOC", componentes)

    def test_reconoce_autentic_tokens_crl_y_canal_ventas(self):
        incidentes = pd.DataFrame([
            {"numero": "INC1", "descripcion": "Autentic presenta falla de autenticación y acceso"},
            {"numero": "INC2", "descripcion": "Token físico no permite firmar"},
            {"numero": "INC3", "descripcion": "Token virtual no permite firmar"},
            {"numero": "INC4", "descripcion": "CRL presenta error de sincronización y falla"},
            {"numero": "INC5", "descripcion": "Canal de ventas caído y no responde"},
        ])
        matriz = construir_matriz_riesgo_incidentes(incidentes)
        componentes = set(matriz["componente_detectado"])
        self.assertTrue({"Autentic", "Token físico", "Token virtual", "CLR / PKI", "Canal de ventas"}.issubset(componentes))


if __name__ == "__main__":
    unittest.main()
