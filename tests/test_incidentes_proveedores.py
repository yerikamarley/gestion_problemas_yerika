import unittest

import pandas as pd

from app_logic import preparar_incidentes


class IncidentesProveedoresTest(unittest.TestCase):
    def test_conserva_nombre_proveedor_del_excel(self):
        entrada = pd.DataFrame(
            {
                "Número": ["INC0017298"],
                "Escalado a proveedor": [True],
                "Nombre del proveedor": ["Autentic"],
            }
        )

        resultado = preparar_incidentes(entrada)

        self.assertEqual("Autentic", resultado.iloc[0]["nombre_proveedor"])
        self.assertTrue(resultado.iloc[0]["escalado_proveedor"])

    def test_prefiere_columna_de_proveedor_con_datos_si_hay_encabezado_duplicado(self):
        entrada = pd.DataFrame(
            [["INC0017298", None, "Autentic"]],
            columns=["Número", "Nombre del Proveedor", "Nombre del proveedor"],
        )

        resultado = preparar_incidentes(entrada)

        self.assertEqual("Autentic", resultado.iloc[0]["nombre_proveedor"])


if __name__ == "__main__":
    unittest.main()
