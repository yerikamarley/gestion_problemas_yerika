import unittest

import pandas as pd

from config.equipo_soporte import (
    SEGMENTO_EQUIPO_SOPORTE,
    SEGMENTO_OTROS_RESPONSABLES,
    SEGMENTO_SIN_ASIGNACION,
)
from services.casos import COL_SEGMENTO_ASIGNACION, segmentar_casos_por_asignacion, top_categorias


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

    def test_separa_equipo_otros_y_sin_asignacion(self):
        df = pd.DataFrame(
            {
                "asignado": [
                    "Alejandra Preciado",
                    "PAULA PÁEZ",
                    "Otra Persona",
                    "",
                    None,
                    pd.NA,
                ]
            }
        )
        segmentos = segmentar_casos_por_asignacion(df)
        self.assertEqual(2, len(segmentos["equipo"]))
        self.assertEqual(1, len(segmentos["otros"]))
        self.assertEqual(3, len(segmentos["sin_asignacion"]))
        self.assertEqual(
            [SEGMENTO_EQUIPO_SOPORTE, SEGMENTO_OTROS_RESPONSABLES, SEGMENTO_SIN_ASIGNACION],
            segmentos["todos"][COL_SEGMENTO_ASIGNACION].drop_duplicates().tolist(),
        )


if __name__ == "__main__":
    unittest.main()
