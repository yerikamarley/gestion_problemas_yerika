import unittest

import pandas as pd

from app_logic import construir_matriz_riesgo_incidentes


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


if __name__ == "__main__":
    unittest.main()
