import unittest

import pandas as pd

from config.clientes_clave import CLIENTES_CLAVE, GRUPOS_CLIENTES_CLAVE
from services.clientes_clave import detectar_cliente_clave, detectar_cliente_en_fila


class ClientesClaveConfigTest(unittest.TestCase):
    def test_catalogo_no_tiene_duplicados(self):
        self.assertEqual(80, len(CLIENTES_CLAVE))
        self.assertEqual(len(CLIENTES_CLAVE), len(set(CLIENTES_CLAVE)))

    def test_grupos_tienen_las_cantidades_configuradas(self):
        self.assertEqual([17, 57, 6], [len(clientes) for clientes in GRUPOS_CLIENTES_CLAVE.values()])


class DeteccionClientesClaveTest(unittest.TestCase):
    def test_detecta_alias_frecuentes(self):
        casos = {
            "BANCO DAVIVIENDA S.A.": "Davivienda",
            "SUFI BANCOLOMBIA": "Sufi",
            "COMCEL": "Claro",
            "RCI COLOMBIA S.A COMPAÑÍA DE FINANCIAMIENTO": "RCI",
            "Cámara de Comercio de Bogotá": "Cámara de Comercio de Bogotá",
        }
        for texto, esperado in casos.items():
            with self.subTest(texto=texto):
                self.assertEqual(esperado, detectar_cliente_clave(texto))

    def test_no_detecta_alias_dentro_de_otra_palabra(self):
        self.assertEqual("", detectar_cliente_clave("claroscuro"))

    def test_informa_el_campo_que_identifico_al_cliente(self):
        fila = pd.Series({"empresa": "Sin coincidencia", "descripcion": "Caso reportado por COMCEL"})
        self.assertEqual(("Claro", "descripcion"), detectar_cliente_en_fila(fila, ["empresa", "descripcion"]))


if __name__ == "__main__":
    unittest.main()

