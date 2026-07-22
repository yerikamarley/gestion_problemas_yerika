import unittest

import pandas as pd

from services.casos import top_categorias


class CasosServiceTest(unittest.TestCase):
    def test_retorna_top_cinco_con_cantidades(self):
        df = pd.DataFrame({"causa": ["Firma", "Token", "Firma", "Agenda", "Pago", "Token", "Otro"]})
        resumen = top_categorias(df, "causa", "Causa", top_n=5)
        self.assertEqual(5, len(resumen))
        self.assertEqual("Firma", resumen.iloc[0]["Causa"])
        self.assertEqual(2, resumen.iloc[0]["Cantidad"])

    def test_controla_columna_ausente(self):
        resumen = top_categorias(pd.DataFrame({"otra": [1]}), "causa", "Causa")
        self.assertEqual(["Causa", "Cantidad"], resumen.columns.tolist())
        self.assertTrue(resumen.empty)

    def test_agrupa_valores_vacios(self):
        df = pd.DataFrame({"servicio": [None, "", "Firma"]})
        resumen = top_categorias(df, "servicio", "Servicio", valor_vacio="Sin servicio")
        fila = resumen[resumen["Servicio"] == "Sin servicio"].iloc[0]
        self.assertEqual(2, fila["Cantidad"])


if __name__ == "__main__":
    unittest.main()

