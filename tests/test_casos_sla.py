import unittest
from unittest.mock import MagicMock, patch

import pandas as pd

import app_logic
from services.casos_sla import (
    COL_ESTADO_SLA, COLUMNAS_CASOS_PLATAFORMA, agregar_sla_casos,
    guardar_fecha_plataforma, resumen_sla_casos, tabla_casos_plataforma,
)


class CasosSlaTest(unittest.TestCase):
    def caso(self, **cambios):
        return {"numero": "CS1", "estado": "Nuevo", "creado": "2026-08-10 08:00:00",
                "fecha_vencimiento_sla": "2026-08-13 14:00:00", "cerrado": "", **cambios}

    def test_abiertos_limite_y_un_segundo_de_atraso(self):
        df = pd.DataFrame([self.caso()])
        al_limite = agregar_sla_casos(df, "2026-08-13 14:00:00").iloc[0]
        vencido = agregar_sla_casos(df, "2026-08-13 14:00:01").iloc[0]
        self.assertEqual("Dentro del plazo", al_limite[COL_ESTADO_SLA])
        self.assertEqual(0, al_limite["horas_para_vencer"])
        self.assertEqual("Vencido", vencido[COL_ESTADO_SLA])
        self.assertLess(vencido["horas_para_vencer"], 0)

    def test_cerrados_se_evalua_fecha_y_no_objetivo_fijo(self):
        df = pd.DataFrame([
            self.caso(estado="Cerrado", cerrado="2026-08-20 14:00:00", fecha_vencimiento_sla="2026-08-20 14:00:00"),
            self.caso(estado="Cerrado", cerrado="2026-08-10 10:00:00", fecha_vencimiento_sla="2026-08-10 09:00:00"),
        ])
        result = agregar_sla_casos(df)
        self.assertEqual(["Cumple", "No cumple"], result[COL_ESTADO_SLA].tolist())
        self.assertEqual(50, resumen_sla_casos(df)["porcentaje"])

    def test_no_evaluables_no_entran_en_porcentaje(self):
        df = pd.DataFrame([
            self.caso(estado="Cerrado", cerrado="2026-08-13 12:00:00"),
            self.caso(),
            self.caso(fecha_vencimiento_sla=None),
            self.caso(estado="Cerrado"),
            self.caso(estado="Cerrado", cerrado="2026-08-01 12:00:00"),
            self.caso(fecha_vencimiento_sla="fecha inválida"),
        ])
        resumen = resumen_sla_casos(df)
        self.assertEqual({"cumple": 1, "no_cumple": 0, "evaluados": 1, "porcentaje": 100}, resumen)
        self.assertEqual(
            ["Sin fecha de vencimiento", "Sin fecha de cierre", "Fechas inconsistentes", "Fechas inconsistentes"],
            agregar_sla_casos(df).iloc[2:][COL_ESTADO_SLA].tolist(),
        )
        self.assertIsNone(resumen_sla_casos(pd.DataFrame([self.caso()]))["porcentaje"])
        self.assertIsNone(resumen_sla_casos(pd.DataFrame())["porcentaje"])

    def test_zona_horaria_no_depende_del_servidor(self):
        df = pd.DataFrame([self.caso()])
        result = agregar_sla_casos(df, "2026-08-13T19:00:00Z")
        self.assertEqual("Dentro del plazo", result.iloc[0][COL_ESTADO_SLA])
        self.assertEqual(0, result.iloc[0]["horas_para_vencer"])
        self.assertEqual("2026-08-13 14:00:00", guardar_fecha_plataforma("2026-08-13T19:00:00Z"))

    def test_formato_dia_mes_y_ausencia_columna(self):
        df = pd.DataFrame([self.caso(fecha_vencimiento_sla="13/08/2026 14:00:00")])
        self.assertEqual("Vencido", agregar_sla_casos(df, "2026-08-14").iloc[0][COL_ESTADO_SLA])
        self.assertEqual("Sin fecha de vencimiento", agregar_sla_casos(df.drop(columns="fecha_vencimiento_sla")).iloc[0][COL_ESTADO_SLA])

    def test_importacion_y_orden_independientes_de_posicion(self):
        origen = {etiqueta: f"dato-{campo}" for campo, etiqueta in COLUMNAS_CASOS_PLATAFORMA.items()}
        origen["Número"] = "CS1"
        origen["Caso"] = "Referencia CS1"
        df = pd.DataFrame([origen])
        preparado = app_logic.preparar_casos(df[df.columns[::-1]])
        for campo, etiqueta in COLUMNAS_CASOS_PLATAFORMA.items():
            self.assertEqual(origen[etiqueta], preparado.iloc[0][campo], campo)
        visible = tabla_casos_plataforma(agregar_sla_casos(preparado), [COL_ESTADO_SLA])
        self.assertEqual([*COLUMNAS_CASOS_PLATAFORMA.values(), "Estado SLA"], visible.columns.tolist())
        self.assertEqual("Referencia CS1", visible.iloc[0]["Caso"])

    def test_alias_caso_antiguo(self):
        preparado = app_logic.preparar_casos(pd.DataFrame({"Caso": ["CS1"]}))
        self.assertEqual("CS1", preparado.iloc[0]["numero"])

    def test_consulta_historica_conserva_vencimiento(self):
        with patch.object(app_logic, "exigir_contexto_lectura"), \
             patch.object(app_logic, "read_table_years", return_value=pd.DataFrame()) as leer:
            app_logic.load_casos_anios([2025, 2026])
        self.assertIn("fecha_vencimiento_sla", leer.call_args.kwargs["columns"])

    def test_guardado_incluye_campos_nuevos_sin_conexion_real(self):
        df = app_logic.preparar_casos(pd.DataFrame([{
            "Número": "CS1", "Creado": "2026-08-10 08:00:00",
            "Fecha de vencimiento del SLA": "2026-08-13T19:00:00Z",
            "Causa Raiz*": "Causa de origen", "Nombre del proveedor": "Proveedor",
            "Numero de caso externo": "EXT1", "Notas del trabajo": "Seguimiento",
        }]))
        conn = MagicMock()
        with patch.object(app_logic, "get_conn", return_value=conn), \
             patch.object(app_logic, "_bloquear_y_validar_admin"), \
             patch.object(app_logic, "bloquear_escritura_casos"), \
             patch.object(app_logic, "numeros_existentes", return_value=set()), \
             patch.object(app_logic, "ejecutar_upserts_lote") as guardar:
            app_logic._guardar_casos_preparados(df, False, 0)
        args = guardar.call_args.args
        registro = dict(zip(args[2], args[4][0]))
        self.assertEqual(len(args[2]), len(args[4][0]))
        self.assertEqual("2026-08-13 14:00:00", registro["fecha_vencimiento_sla"])
        self.assertEqual("Causa de origen", registro["causa"])
        self.assertEqual("Proveedor", registro["nombre_proveedor"])
        self.assertEqual("EXT1", registro["numero_caso_externo"])
        self.assertEqual("Seguimiento", registro["notas_trabajo"])
        conn.commit.assert_called_once()


class IntegracionSlaCasosTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        import app_ui
        cls.ui = app_ui

    def base(self, con_vencimiento=True):
        df = app_logic.preparar_casos(pd.DataFrame([
            {"Número": "CS1", "Estado": "Cerrado", "Creado": "2026-08-10 08:00:00",
             "Cerrado": "2026-08-20 14:00:00", "Cuenta": "Davivienda", "Asignado a": "",
             "Fecha de vencimiento del SLA": "2026-08-20 14:00:00" if con_vencimiento else ""},
            {"Número": "CS2", "Estado": "Nuevo", "Creado": "2026-08-10 08:00:00",
             "Cuenta": "Davivienda", "Asignado a": "", "Fecha de vencimiento del SLA": "2026-08-11 08:00:00"},
        ]))
        for col in app_logic.CASE_ALIASES:
            df[col] = df[col].fillna("")
        df["tiempo_respuesta"] = [70, "En proceso"]
        return app_logic.aplicar_tipificaciones_casos(df)

    def test_dashboard_comparativos_y_clientes_comparten_regla(self):
        df = self.base()
        ui = self.ui
        _, kpi = ui.preparar_kpi_casos_cliente_externo(df)
        self.assertEqual(100, kpi["cumplimiento_sla"])
        self.assertEqual(1, kpi["evaluados_sla"])
        ligero = ui.preparar_casos_kpi_comparativo_ligero(df)
        self.assertEqual(100, ui.metricas_casos_comparativo(ligero, 2026)["SLA %"])
        self.assertEqual(100, ui.metricas_casos_comparativo_rango(ligero, "Base", "2026-08-01", "2026-08-31")["SLA %"])
        self.assertEqual(100, ui.calcular_sla_casos_clientes(df))
        clientes = ui.preparar_casos_clientes_clave_comparativo(df)
        resumen = ui.resumen_casos_clientes_clave_periodo(clientes, "Agosto")
        self.assertEqual("100%", resumen.iloc[0]["SLA %"])

    def test_sin_vencimiento_no_se_convierte_en_cero_por_ciento(self):
        df = self.base(False)
        ui = self.ui
        _, kpi = ui.preparar_kpi_casos_cliente_externo(df)
        self.assertIsNone(kpi["cumplimiento_sla"])
        self.assertEqual(1, kpi["sin_evaluacion_sla"])
        ligero = ui.preparar_casos_kpi_comparativo_ligero(df)
        self.assertIsNone(ui.metricas_casos_comparativo(ligero, 2026)["SLA %"])
        self.assertEqual("Sin dato", ui.formato_porcentaje_presentacion(ui.calcular_sla_casos_clientes(df)))
        tabla = pd.DataFrame([
            {"Periodo": "Base", "SLA %": None}, {"Periodo": "Comparado", "SLA %": None},
        ])
        variacion = ui.tabla_variacion_kpi_casos(tabla)
        self.assertTrue(pd.isna(variacion[variacion["Metrica"] == "SLA %"].iloc[0]["Diferencia"]))

    def test_seguimiento_usa_horas_calendario_hasta_vencimiento(self):
        ui = self.ui
        df = self.base()
        seguimiento = ui.preparar_seguimiento_casos(df, ahora="2026-08-11 09:00:00")
        abierto = seguimiento[seguimiento["numero"] == "CS2"].iloc[0]
        self.assertEqual("Vencido", abierto[COL_ESTADO_SLA])
        self.assertTrue(abierto[ui.TEXT_VENCIDO])
        self.assertEqual(-1, abierto[ui.TEXT_HORAS_PARA_VENCER])

    def test_vista_renderiza_orden_exportacion_y_filtra_sla(self):
        from streamlit.testing.v1 import AppTest
        app = AppTest.from_string('''
from datetime import date
from unittest.mock import patch
import app_ui as ui
from tests.test_casos_sla import IntegracionSlaCasosTest
df = ui.agregar_tipologia_soporte_casos(IntegracionSlaCasosTest().base())
with patch.object(ui, "selector_fechas_casos", return_value=(date(2026, 8, 1), date(2026, 8, 31), "Agosto 2026")), \\
     patch.object(ui, "cargar_casos_soporte_filtrados_cache", return_value=df):
    ui.vista_casos()
''').run()
        self.assertEqual(0, len(app.exception))
        self.assertEqual(list(COLUMNAS_CASOS_PLATAFORMA.values()), app.dataframe[0].value.columns[:26].tolist())
        self.assertIn("Horas restantes / margen al cierre", app.dataframe[0].value.columns)
        app.selectbox(key="estado_sla_casos").select("Vencido").run()
        self.assertEqual(0, len(app.exception))
        self.assertEqual(["CS2"], app.dataframe[0].value["Número"].tolist())
        self.assertEqual(["Vencido"], app.dataframe[0].value["Estado SLA"].tolist())


if __name__ == "__main__":
    unittest.main()
