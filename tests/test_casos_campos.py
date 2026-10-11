import unittest
from unittest.mock import MagicMock, patch

import pandas as pd

import app_logic
from repositories.servicenow import leer_campos_casos
from services.casos_campos import ampliar_columnas_casos, COLUMNAS_CASOS_PORTAL
from services.casos_sla import tabla_casos_plataforma
from tests.test_servicenow import field


def campos_ejemplo():
    return {
        "number": field("CS0215905"), "case": field("id", "Solicitud de revisión CS0215905"),
        "category": field("1", "Problema"), "subcategory": field("2", "Pregunta"),
        "assignment_group": field("id-grupo", "Soporte Mesa de Ayuda"),
        "assigned_to": field("id-agente", "Orlando Mendez"),
        "sys_created_on": field("2026-10-09 23:20:48", "09/10/2026 06:20:48 PM"),
        "assigned_on": field("2026-10-10 00:12:08"), "opened_at": field("2026-10-09 23:20:48"),
        "made_sla": field("true", "verdadero"), "u_escalado_a_proveedor": field("false", "falso"),
        "u_boolean_3": field(False), "reassignment_count": field(0),
        "knowledge": field(False), "activity_due": field("UNKNOWN"),
        "product": field("id-producto", "Certihuella"), "u_producto": field(""),
        "u_campo_nuevo": field("id-desconocido", "Campo futuro"),
    }


def slas_ejemplo():
    return [{"sys_id": "sla-1", "sla": "Casos_Prioridad_Moderada", "stage": "En curso",
             "start_time": "2026-10-09 23:20:48", "planned_end_time": "2026-10-15 18:00:00",
             "has_breached": False},
            {"sys_id": "sla-2", "sla": "SLA respuesta", "stage": "Completado",
             "start_time": "2026-10-09 23:20:48", "planned_end_time": "2026-10-10 18:00:00",
             "has_breached": True}]


class ColumnasCompletasTest(unittest.TestCase):
    def test_campos_legibles_fechas_booleanos_y_slas_sin_duplicar_casos(self):
        historial = "Correo Entrante: " + "Detalle completo del cliente " * 2000 + " FIN"
        base = pd.DataFrame([
            {"numero": "CS0215905", "historial_conversacion": historial, "notas_trabajo": "Nota histórica"},
            {"numero": "CS0199999", "descripcion": "Caso de Excel antiguo"},
        ], index=[15, 42])
        raw = campos_ejemplo()
        raw["work_notes"] = field("")
        tabla = tabla_casos_plataforma(ampliar_columnas_casos(base, [("CS0215905", raw, slas_ejemplo())]))
        self.assertEqual([15, 42], tabla.index.tolist())
        self.assertEqual(2, len(tabla))
        row = tabla.loc[15]
        self.assertEqual("Solicitud de revisión CS0215905", row["Caso"])
        self.assertEqual("Soporte Mesa de Ayuda", row["Grupo de asignación"])
        self.assertEqual("Orlando Mendez", row["Asignado a"])
        self.assertEqual("Problema", row["Categoría"])
        self.assertEqual("Pregunta", row["Subcategoría"])
        self.assertEqual("2026-10-09 18:20:48", row["Creado"])
        self.assertEqual("2026-10-09 19:12:08", row["Asignado el"])
        self.assertEqual("verdadero", row["SLA creado"])
        self.assertEqual("falso", row["Conocimiento"])
        self.assertEqual("falso", row["Escalado a proveedor"])
        self.assertEqual("falso", row["Escalado a proveedor (u_boolean_3)"])
        self.assertEqual("0", row["Número de reasignación"])
        self.assertEqual("UNKNOWN", row["Vencimiento de la actividad"])
        self.assertEqual("Certihuella", row["Producto"])
        self.assertEqual("", row["Producto (u_producto)"])
        self.assertEqual("Campo futuro", row["u_campo_nuevo"])
        self.assertEqual("2", row["Cantidad de SLAs"])
        self.assertEqual("Casos_Prioridad_Moderada", row["SLA 1 · Nombre"])
        self.assertEqual("2026-10-15 13:00:00", row["SLA 1 · Vencimiento"])
        self.assertEqual("falso", row["SLA 1 · Incumplido"])
        self.assertEqual("verdadero", row["SLA 2 · Incumplido"])
        self.assertEqual(historial, row["Historial de la conversación"])
        self.assertEqual("Nota histórica", row["Notas del trabajo"])
        self.assertEqual("Caso de Excel antiguo", tabla.loc[42, "Breve descripción"])
        self.assertFalse(tabla.columns.duplicated().any())
        self.assertNotIn("sn__category", base.columns)

    def test_misma_etiqueta_no_oculta_campos_distintos_y_campos_en_casos_tardios(self):
        base = pd.DataFrame({"numero": ["CS0215905", "CS0215906"]})
        raw = campos_ejemplo()
        raw["opened_by"] = field("id", "Solicitante")
        raw["sys_created_by"] = field("cliente@example.test")
        raw["u_producto"] = field("id", "Producto personalizado")
        otro = {"number": field("CS0215906"), "u_otro": field("Nuevo campo")}
        tabla = tabla_casos_plataforma(ampliar_columnas_casos(base, [
            ("CS0215905", raw, []), ("CS0215906", otro, [])]))
        self.assertEqual("cliente@example.test", tabla.iloc[0]["Creado por"])
        self.assertEqual("Solicitante", tabla.iloc[0]["Abierto por"])
        self.assertEqual("Producto personalizado", tabla.iloc[0]["Producto"])
        self.assertEqual("Certihuella", tabla.iloc[0]["Producto (product)"])
        self.assertEqual("Nuevo campo", tabla.iloc[1]["u_otro"])
        self.assertFalse(tabla.columns.duplicated().any())

    def test_lectura_664_casos_una_consulta_proyectada_y_parametrizada(self):
        conn = MagicMock()
        conn.execute.return_value.fetchall.return_value = [("CS0215905", campos_ejemplo(), [])]
        numeros = [f"CS{i + 1000000}" for i in range(664)]
        registros = leer_campos_casos(conn, numeros + [numeros[0]])
        self.assertEqual(1, len(registros))
        conn.execute.assert_called_once()
        sql, params = conn.execute.call_args.args
        self.assertIn("documento->'campos_raw'", sql)
        self.assertIn("documento->'slas'", sql)
        self.assertNotIn("SELECT documento FROM", sql)
        self.assertNotIn("historial_conversacion", sql)
        self.assertNotIn(numeros[0], sql)
        self.assertEqual(numeros, params[0])
        leer_campos_casos(conn, [])
        self.assertEqual(1, conn.execute.call_count)

    def test_fachada_exige_autorizacion_y_cierra_conexion_incluso_al_fallar(self):
        with self.assertRaises(app_logic.AutorizacionError), patch.object(app_logic, "get_conn") as conectar:
            app_logic.load_campos_casos(["CS0215905"])
        conectar.assert_not_called()
        conn = MagicMock()
        with patch.object(app_logic, "exigir_contexto_lectura"), \
             patch.object(app_logic, "get_conn", return_value=conn), \
             patch.object(app_logic, "leer_campos_casos", side_effect=RuntimeError("fallo local")), \
             self.assertRaises(RuntimeError):
            app_logic.load_campos_casos(["CS0215905"])
        conn.close.assert_called_once()

    def test_vista_y_descarga_muestran_las_mismas_columnas_completas(self):
        import app_ui  # Inicializar el módulo fuera del contexto de AppTest.
        from streamlit.testing.v1 import AppTest
        app = AppTest.from_string('''
from datetime import date
from unittest.mock import patch
import pandas as pd
import streamlit as st
import app_ui as ui
from tests.test_casos_sla import IntegracionSlaCasosTest
from tests.test_casos_campos import campos_ejemplo, slas_ejemplo
df = ui.agregar_tipologia_soporte_casos(IntegracionSlaCasosTest().base())
df.loc[0, "numero"] = "CS0215905"
df["historial_conversacion"] = "Preocupacion del cliente por fallas recurrentes del portal"
df["Grupo migraci\\u00f3n PKI"] = "Clasificacion temporal anterior"
df["grupo_migracion_pki"] = "Clasificacion temporal anterior"
st.session_state["grupo_cliente_clave_casos"] = "Migraci\\u00f3n PKI"
def captura(df, *args):
    st.session_state["descarga_completa"] = df
with patch.object(ui, "selector_fechas_casos", return_value=(date(2026, 8, 1), date(2026, 8, 31), "Agosto 2026")), \\
     patch.object(ui, "cargar_casos_soporte_filtrados_cache", return_value=df), \\
     patch.object(ui, "cargar_casos_campos_cache", return_value=[("CS0215905", campos_ejemplo(), slas_ejemplo())]), \\
     patch.object(ui, "render_descarga_dataframe", side_effect=captura):
    ui.vista_casos()
''').run()
        self.assertEqual(0, len(app.exception))
        tabla = app.dataframe[0].value
        self.assertEqual(2, len(tabla))
        self.assertEqual(list(COLUMNAS_CASOS_PORTAL), tabla.columns.tolist())
        self.assertNotIn("Grupo migración PKI", tabla.columns)
        self.assertNotIn("grupo_migracion_pki", tabla.columns)
        filtro = app.selectbox(key="grupo_cliente_clave_casos")
        self.assertNotIn("Migración PKI", filtro.options)
        self.assertEqual("Todos", filtro.value)
        self.assertIn("Asobancaria", filtro.options)
        for columna in ["Caso", "Categoría", "Grupo de asignación", "SLA 2 · Nombre", "u_campo_nuevo"]:
            self.assertNotIn(columna, tabla.columns)
        self.assertEqual("En curso", tabla.iloc[0]["SLA 1 · Estado"])
        self.assertEqual("2026-10-09 18:20:48", tabla.iloc[0]["SLA 1 · Inicio"])
        self.assertEqual("2026-10-15 13:00:00", tabla.iloc[0]["SLA 1 · Vencimiento"])
        pd.testing.assert_frame_equal(
            tabla.fillna("").astype(str),
            app.session_state["descarga_completa"].reset_index(drop=True).fillna("").astype(str),
        )


if __name__ == "__main__":
    unittest.main()
