import io
import unittest

import pandas as pd

from services.migracion_pki import (
    GRUPO_PKI_CON_COMPONENTES,
    GRUPO_PKI_SIN_COMPONENTES,
    agregar_grupo_migracion,
    leer_lista_migracion,
)


def excel_bytes(base, componentes):
    salida = io.BytesIO()
    with pd.ExcelWriter(salida, engine="openpyxl") as escritor:
        base.to_excel(escritor, sheet_name="mesa de ayuda bdf", index=False)
        componentes.to_excel(escritor, sheet_name="ClientesconComponentes", index=False)
    return salida.getvalue()


class MigracionPkiTest(unittest.TestCase):
    def lista(self):
        return leer_lista_migracion(io.BytesIO(excel_bytes(
            pd.DataFrame({"CORREO": ["persona@empresa.com", "otro@empresa.com"], "razon social": ["Empresa Uno", "Empresa Dos"], "fecha vencimiento": ["2026-09-01", "2026-09-02"]}),
            pd.DataFrame({"empresa": ["Empresa Uno"], "cliente": ["Persona Uno"], "correo": ["persona@empresa.com"], "fecha de vencimiento": ["2026-09-01"]}),
        )))

    def test_lee_las_dos_hojas_y_reconoce_subgrupo(self):
        lista = self.lista()
        self.assertEqual(2, lista["filas_base"])
        self.assertEqual(1, lista["filas_componentes"])
        casos = pd.DataFrame([
            {"numero": "1", "cuenta": "Empresa Uno", "contacto": "", "creado_por": ""},
            {"numero": "2", "cuenta": "Empresa Dos", "contacto": "", "creado_por": ""},
            {"numero": "3", "cuenta": "Fuera", "contacto": "persona@empresa.com", "creado_por": ""},
            {"numero": "4", "cuenta": "Fuera", "contacto": "", "creado_por": "No relacionado"},
        ])
        resultado = agregar_grupo_migracion(casos, lista)
        self.assertEqual([GRUPO_PKI_CON_COMPONENTES, GRUPO_PKI_SIN_COMPONENTES, GRUPO_PKI_CON_COMPONENTES, ""], resultado["grupo_migracion_pki"].tolist())

    def test_no_confunde_sin_asignacion_con_sin_componentes(self):
        lista = self.lista()
        casos = pd.DataFrame([{"numero": "1", "cuenta": "", "contacto": "", "creado_por": ""}])
        self.assertEqual("", agregar_grupo_migracion(casos, lista).iloc[0]["grupo_migracion_pki"])

    def test_rechaza_hojas_incorrectas(self):
        salida = io.BytesIO()
        with pd.ExcelWriter(salida, engine="openpyxl") as escritor:
            pd.DataFrame({"x": [1]}).to_excel(escritor, sheet_name="Otra", index=False)
        with self.assertRaises(ValueError):
            leer_lista_migracion(io.BytesIO(salida.getvalue()))


if __name__ == "__main__":
    unittest.main()
