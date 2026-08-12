import unittest

import pandas as pd

from app_logic import preparar_problemas


class ProblemasTest(unittest.TestCase):
    def test_normaliza_columnas_del_excel(self):
        entrada = pd.DataFrame({
            "Número": ["PRB001"],
            "Declaración de problema": ["Problema de prueba"],
            "Creado": [pd.Timestamp("2026-08-01")],
            "Estado": ["Abierto"],
            "Prioridad": ["Media"],
            "Incidentes relacionados": [2],
        })
        resultado = preparar_problemas(entrada)
        self.assertEqual("PRB001", resultado.iloc[0]["numero"])
        self.assertEqual("Problema de prueba", resultado.iloc[0]["declaracion_problema"])
        self.assertEqual(2, resultado.iloc[0]["incidentes_relacionados"])

    def test_descarta_filas_sin_numero_y_duplicados(self):
        entrada = pd.DataFrame({
            "Número": ["PRB001", "", "PRB001"],
            "Estado": ["Nuevo", "Nuevo", "Análisis"],
        })
        resultado = preparar_problemas(entrada)
        self.assertEqual(1, len(resultado))
        self.assertEqual("Análisis", resultado.iloc[0]["estado"])


if __name__ == "__main__":
    unittest.main()
