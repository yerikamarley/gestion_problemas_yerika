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
            {"numero": "INC1", "creado": "2026-01-05", "causa_raiz_auto": "Base de datos"},
            {"numero": "INC2", "creado": "2026-01-20", "causa_raiz_auto": "Base de datos"},
            {"numero": "INC3", "creado": "2026-03-02", "causa_raiz_auto": "Base de datos"},
            {"numero": "INC4", "creado": "2026-02-01", "causa_raiz_auto": "Red"},
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

        self.assertEqual(1, len(matriz))
        self.assertEqual("Canal de ventas", matriz.iloc[0]["criterio_similitud"])
        self.assertIn("INC1", matriz.iloc[0]["evidencia_analizada"])

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


if __name__ == "__main__":
    unittest.main()
