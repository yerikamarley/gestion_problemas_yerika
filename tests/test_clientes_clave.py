import unittest

import pandas as pd

from config.clientes_clave import CLIENTES_CLAVE, COOPCENTRAL, GRUPOS_CLIENTES_CLAVE
from services.clientes_clave import (
    detectar_cliente_clave,
    detectar_cliente_en_fila,
    detectar_grupo_cliente_clave,
    filtrar_por_cliente_o_texto,
    filtrar_por_grupo_cliente_clave,
    serie_grupo_cliente_clave,
)


class ClientesClaveConfigTest(unittest.TestCase):
    def test_catalogo_no_tiene_duplicados(self):
        self.assertEqual(139, len(CLIENTES_CLAVE))
        self.assertEqual(len(CLIENTES_CLAVE), len(set(CLIENTES_CLAVE)))

    def test_grupos_tienen_las_cantidades_configuradas(self):
        self.assertEqual([17, 57, 59, 6], [len(clientes) for clientes in GRUPOS_CLIENTES_CLAVE.values()])

    def test_coopcentral_tiene_el_catalogo_completo(self):
        self.assertEqual(59, len(COOPCENTRAL))
        self.assertIn("Caja Unión", COOPCENTRAL)
        self.assertIn("Coopcentral", COOPCENTRAL)
        self.assertIn("Coodin", COOPCENTRAL)


class DeteccionClientesClaveTest(unittest.TestCase):
    def test_detecta_alias_frecuentes(self):
        casos = {
            "BANCO DAVIVIENDA S.A.": "Davivienda",
            "SUFI BANCOLOMBIA": "Sufi",
            "COMCEL": "Claro",
            "RCI COLOMBIA S.A COMPAÑÍA DE FINANCIAMIENTO": "RCI",
            "Cámara de Comercio de Bogotá": "Cámara de Comercio de Bogotá",
            "COOPSURAMERICA": "Coopsuramérica",
            "Cooperativa Leon XIII Ltda Guatape": "Cooperativa León XIII Ltda Guatapé",
        }
        for texto, esperado in casos.items():
            with self.subTest(texto=texto):
                self.assertEqual(esperado, detectar_cliente_clave(texto))

    def test_no_detecta_alias_dentro_de_otra_palabra(self):
        self.assertEqual("", detectar_cliente_clave("claroscuro"))

    def test_informa_el_campo_que_identifico_al_cliente(self):
        fila = pd.Series({"empresa": "Sin coincidencia", "descripcion": "Caso reportado por COMCEL"})
        self.assertEqual(("Claro", "descripcion"), detectar_cliente_en_fila(fila, ["empresa", "descripcion"]))

    def test_detecta_y_filtra_grupo_cliente_clave(self):
        self.assertEqual("Asobancaria", detectar_grupo_cliente_clave("BANCO DAVIVIENDA S.A."))
        self.assertEqual("Coopcentral", detectar_grupo_cliente_clave("Caso de COOPSURAMERICA"))
        df = pd.DataFrame(
            {
                "cuenta": [
                    "Banco Davivienda S.A.",
                    "Cámara de Comercio de Bogotá",
                    "Coopsuramerica",
                    "Cliente no configurado",
                ]
            }
        )
        filtrado = filtrar_por_grupo_cliente_clave(df, "cuenta", "Coopcentral")
        self.assertEqual(["Coopsuramerica"], filtrado["cuenta"].tolist())

    def test_detecta_cliente_por_correo_cuando_no_hay_cuenta(self):
        self.assertEqual("Codema", detectar_cliente_clave("mespinosa@codema.com.co"))
        df = pd.DataFrame(
            {
                "cuenta": ["", "Davivienda"],
                "creado_por": ["mespinosa@codema.com.co", "usuario@empresa.com"],
            }
        )
        grupos = serie_grupo_cliente_clave(df, ["cuenta", "creado_por"])
        self.assertEqual(["Coopcentral", "Asobancaria"], grupos.tolist())
        filtrado = filtrar_por_grupo_cliente_clave(df, ["cuenta", "creado_por"], "Coopcentral")
        self.assertEqual(["mespinosa@codema.com.co"], filtrado["creado_por"].tolist())

    def test_filtro_casos_usa_alias_y_campos_de_clientes_vip(self):
        df = pd.DataFrame(
            {
                "numero": ["1", "2", "3", "4"],
                "cuenta": ["COMCEL", "", "Bancolombia S.A", "Otro cliente"],
                "creado_por": ["", "usuario@claro.com", "", "usuario@empresa.com"],
            }
        )

        claro = filtrar_por_cliente_o_texto(df, ["cuenta", "creado_por"], "Claro")
        bancolombia = filtrar_por_cliente_o_texto(df, ["cuenta", "creado_por"], "Bancolombia")

        self.assertEqual(["1", "2"], claro["numero"].tolist())
        self.assertEqual(["3"], bancolombia["numero"].tolist())

    def test_filtro_casos_conserva_busqueda_de_texto_libre(self):
        df = pd.DataFrame(
            {
                "numero": ["1", "2"],
                "cuenta": ["Cliente Especial Norte", "Otro cliente"],
                "creado_por": ["usuario@empresa.com", "contacto@especial.com"],
            }
        )

        filtrado = filtrar_por_cliente_o_texto(df, ["cuenta", "creado_por"], "especial")

        self.assertEqual(["1", "2"], filtrado["numero"].tolist())


if __name__ == "__main__":
    unittest.main()
