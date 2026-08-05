import unittest

import pandas as pd

from services.casos import resumen_diario_soporte


class ControlDiarioSoporteTest(unittest.TestCase):
    def test_calcula_movimiento_y_pendientes_por_dia(self):
        casos = pd.DataFrame([
            {"asignado": "Yader Neira", "estado": "Cerrado", "creado": "2026-08-01 08:00", "cerrado": "2026-08-02 10:00"},
            {"asignado": "", "estado": "Esperando por cliente", "creado": "2026-08-01 09:00", "cerrado": ""},
            {"asignado": "Otro responsable", "estado": "Abierto", "creado": "2026-08-01 10:00", "cerrado": ""},
        ])
        resultado = resumen_diario_soporte(casos, 2026, 8)
        self.assertEqual(2, resultado.loc[0, "Total del día"])
        self.assertEqual(1, resultado.loc[0, "Cerrados"])
        self.assertEqual(1, resultado.loc[0, "Esperando cliente"])
        self.assertEqual(1, resultado.loc[0, "Sin asignación"])

    def test_permite_solo_sin_asignacion(self):
        casos = pd.DataFrame([
            {"asignado": "Paula Paez", "estado": "Abierto", "creado": "2026-08-01", "cerrado": ""},
            {"asignado": None, "estado": "Abierto", "creado": "2026-08-01", "cerrado": ""},
        ])
        resultado = resumen_diario_soporte(casos, 2026, 8, "sin_asignacion")
        self.assertEqual(1, resultado.loc[0, "Total del día"])


if __name__ == "__main__":
    unittest.main()
