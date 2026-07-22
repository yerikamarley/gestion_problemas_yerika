import unittest

import app_logic
from repositories.database import (
    db_placeholders,
    db_sql,
    es_error_db_transitorio,
    lotes,
    validar_identificador_sql,
)
from repositories.tables import agregar_condicion_periodo, limites_periodo, valor_filtro_activo


class DatabaseHelpersTest(unittest.TestCase):
    def test_adapta_placeholders_para_psycopg(self):
        self.assertEqual("SELECT * FROM cases WHERE numero = %s", db_sql("SELECT * FROM cases WHERE numero = ?"))
        self.assertEqual("%s, %s, %s", db_placeholders(3))

    def test_divide_valores_en_lotes(self):
        self.assertEqual([[1, 2], [3, 4], [5]], list(lotes([1, 2, 3, 4, 5], 2)))

    def test_valida_identificadores_sql(self):
        self.assertEqual("incidents_2026", validar_identificador_sql("incidents_2026"))
        with self.assertRaises(ValueError):
            validar_identificador_sql("cases; DROP TABLE cases")

    def test_reconoce_error_transitorio(self):
        error = RuntimeError("conflicto")
        error.sqlstate = "40001"
        self.assertTrue(es_error_db_transitorio(error))

    def test_app_logic_conserva_api_compatible(self):
        self.assertIs(app_logic.db_sql, db_sql)
        self.assertIs(app_logic.validar_identificador_sql, validar_identificador_sql)
        self.assertIs(app_logic.limites_periodo, limites_periodo)

    def test_calcula_limites_de_periodo(self):
        self.assertEqual(("2026-07", "2026-08"), limites_periodo(2026, 7))
        self.assertEqual(("2026-12", "2027-01"), limites_periodo(2026, 12))
        self.assertEqual(("2025", "2026"), limites_periodo(2025))
        self.assertIsNone(limites_periodo())

    def test_agrega_condicion_y_parametros_de_periodo(self):
        where, params = [], []
        agregar_condicion_periodo(where, params, 2026, 7)
        self.assertEqual(["(creado >= ? AND creado < ?)"], where)
        self.assertEqual(["2026-07", "2026-08"], params)

    def test_identifica_filtros_activos(self):
        self.assertTrue(valor_filtro_activo("Abierto"))
        self.assertFalse(valor_filtro_activo("Todos"))


if __name__ == "__main__":
    unittest.main()
