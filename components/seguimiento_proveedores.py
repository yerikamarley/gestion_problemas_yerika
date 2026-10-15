"""Vista única y exportación mensual de proveedores, clientes y ANS."""

from inspect import signature

import pandas as pd
import streamlit as st

from services.proveedores import (
    CONFIRMADO, ETAPAS, MENCION, SIN_CLIENTE, catalogo_ans,
    evaluar_ans_proveedores, filtrar_periodo_proveedores, resumen_ans_proveedores,
)
from utils.proveedores_excel_export import exportar_informe_proveedores, informe_proveedores


MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto",
         "Septiembre", "Octubre", "Noviembre", "Diciembre"]


def mostrar_tabla(df):
    # Streamlit local puede ser anterior al desplegado; soporta ambos contratos.
    ancho = {"width": "stretch"} if signature(st.dataframe).parameters["width"].default == "stretch" else {"use_container_width": True}
    st.dataframe(df, hide_index=True, **ancho)


def exportar_csv(df):
    salida = df.copy()
    for columna in salida.columns:
        # Las fechas del informe son locales y legibles también en Excel.
        if isinstance(salida[columna].dtype, pd.DatetimeTZDtype):
            salida[columna] = salida[columna].dt.strftime("%d/%m/%Y %H:%M:%S")
        if salida[columna].dtype == object:
            # Evita que el contenido de una nota se ejecute como fórmula al abrirlo.
            salida[columna] = salida[columna].map(lambda v: "'" + v if isinstance(v, str) and v.lstrip().startswith(("=", "+", "-", "@")) else v)
    return salida.to_csv(index=False, sep=";").encode("utf-8-sig")


def tabla_fechas(df):
    salida = df.copy()
    for columna in salida.columns:
        if pd.api.types.is_datetime64_any_dtype(salida[columna]):
            salida[columna] = salida[columna].dt.strftime("%d/%m/%Y %H:%M").fillna("")
    return salida


def tabla_objetivos(proveedor):
    datos = catalogo_ans()["proveedores"].get(proveedor)
    if not datos:
        return pd.DataFrame()
    return pd.DataFrame([{"Prioridad": "P" + p, **{ETAPAS.get(etapa, "Seguimiento periódico"): texto for etapa, texto in objetivos["texto"].items()},
                          "Fila del Excel": objetivos["fila"]} for p, objetivos in datos["prioridades"].items()])


def render_seguimiento_proveedores(cargar_datos):
    st.subheader("Seguimiento de proveedores")
    ahora = pd.Timestamp.now(tz="America/Bogota")
    c1, c2 = st.columns(2)
    anio = int(c1.number_input("Año", min_value=2000, max_value=2100, value=ahora.year, key="proveedores_anio"))
    mes = c2.selectbox("Mes", list(range(1, 13)), index=ahora.month - 1, format_func=lambda m: MESES[m - 1], key="proveedores_mes")
    inicio = pd.Timestamp(year=anio, month=mes, day=1, tz="America/Bogota")
    fin = inicio + pd.DateOffset(months=1)
    if inicio > ahora:
        st.info("Selecciona el mes actual o un mes anterior para consultar el informe.")
        return
    corte = min(ahora, fin - pd.Timedelta(microseconds=1))
    st.caption(f"{MESES[mes - 1]} {anio} · Corte: {corte:%d/%m/%Y %H:%M} (Colombia). "
               "Incluye escalados del período y pendientes anteriores. Las menciones sin escalamiento se muestran aparte.")
    with st.spinner("Cargando casos e incidentes de proveedores..."):
        datos = filtrar_periodo_proveedores(cargar_datos(), inicio, fin)
    proveedores = sorted(set(catalogo_ans()["proveedores"]) | set(datos.get("Proveedor", [])))
    c1, c2, c3 = st.columns(3)
    proveedor = c1.selectbox("Proveedor", ["Todos"] + proveedores, key="proveedores_filtro")
    cliente = c2.selectbox("Cliente", ["Todos"] + sorted(datos.get("Cliente", pd.Series(dtype=str)).unique().tolist()), key="proveedores_cliente")
    relacion = c3.selectbox("Relación", ["Todos", CONFIRMADO, MENCION], key="proveedores_relacion")
    c1, c2 = st.columns(2)
    tipo = c1.selectbox("Tipo de registro", ["Todos", "Caso", "Incidente"], key="proveedores_tipo")
    estado = c2.selectbox("Estado", ["Todos"] + sorted(datos.get("Estado", pd.Series(dtype=str)).unique().tolist()), key="proveedores_estado")
    for columna, valor in (("Proveedor", proveedor), ("Cliente", cliente), ("Relación", relacion), ("Tipo", tipo), ("Estado", estado)):
        if valor != "Todos":
            datos = datos[datos[columna] == valor]
    if datos.empty:
        st.info("No hay registros de proveedores para este período y estos filtros.")
    else:
        datos = evaluar_ans_proveedores(datos, corte)
        unicos = datos.drop_duplicates(["Tipo", "Número"])
        tarjetas = st.columns(4)
        tarjetas[0].metric("Casos", int((unicos["Tipo"] == "Caso").sum()))
        tarjetas[1].metric("Incidentes", int((unicos["Tipo"] == "Incidente").sum()))
        tarjetas[2].metric("Clientes identificados", unicos.loc[unicos["Cliente"] != SIN_CLIENTE, "Cliente"].nunique())
        tarjetas[3].metric("Escalados", len(datos[datos["Relación"] == CONFIRMADO].drop_duplicates(["Tipo", "Número"])))
        st.download_button("Descargar informe de proveedores (Excel)", exportar_informe_proveedores(datos),
                           file_name=f"informe_proveedores_{anio:04d}_{mes:02d}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           key="proveedores_informe_excel")
        st.caption("El informe incluye proveedor, fecha de creación en Colombia, caso o incidente, historial y estado registrado. "
                   "Respeta los filtros seleccionados. Los historiales largos continúan en filas adicionales del mismo registro.")
        resumen = resumen_ans_proveedores(datos)
        tabs = st.tabs(["Resumen mensual ANS", "Clientes", "Casos e incidentes"])
        with tabs[0]:
            principales = ["Proveedor", "Casos", "Incidentes", "Clientes", "Escalados", "Menciones",
                           "Atención % cumplimiento", "Solución % cumplimiento", "Solución vencidos", "Solución sin datos"]
            mostrar_tabla(resumen[principales])
            st.caption("Cumplimiento: cumplen / evaluados con fechas completas. Los abiertos vencidos y los registros sin datos "
                       "se informan por separado. Un caso relacionado con dos proveedores cuenta una vez en el total y una vez para cada proveedor.")
            with st.expander("Diagnóstico, mitigación y cobertura de los ANS"):
                mostrar_tabla(resumen)
            reporte = resumen.copy()
            reporte.insert(0, "Período", f"{anio:04d}-{mes:02d}")
            reporte.insert(1, "Corte Colombia", corte.strftime("%d/%m/%Y %H:%M:%S"))
            reporte["Criterio"] = "Desde escalamiento; horario contractual; pendientes anteriores incluidos; solo fechas completas en cumplimiento"
            reporte["Fuente ANS"] = catalogo_ans()["fuente"]
            reporte["Filtros"] = f"Proveedor: {proveedor}; Cliente: {cliente}; Relación: {relacion}; Tipo: {tipo}; Estado: {estado}"
            reporte["Horario de atención"] = reporte["Proveedor"].map(lambda p: catalogo_ans()["proveedores"].get(p, {}).get("horario", "Sin ANS definido"))
            reporte["Disponibilidad objetivo %"] = reporte["Proveedor"].map(lambda p: catalogo_ans()["proveedores"].get(p, {}).get("disponibilidad_objetivo"))
            reporte["Disponibilidad observada"] = "Sin medición de indisponibilidad y cobertura"
            st.download_button("Descargar informe mensual ANS", exportar_csv(reporte),
                               file_name=f"ans_proveedores_{anio:04d}_{mes:02d}.csv", mime="text/csv", key="proveedores_reporte")
        with tabs[1]:
            clientes = []
            for (prov, cli), grupo in datos.groupby(["Proveedor", "Cliente"], sort=True):
                registros = grupo.drop_duplicates(["Tipo", "Número"])
                clientes.append({"Proveedor": prov, "Cliente": cli, "Casos": int((registros["Tipo"] == "Caso").sum()),
                                 "Incidentes": int((registros["Tipo"] == "Incidente").sum()),
                                 "Escalados": len(grupo[grupo["Relación"] == CONFIRMADO].drop_duplicates(["Tipo", "Número"]))})
            tabla = pd.DataFrame(clientes)
            mostrar_tabla(tabla)
            st.caption("Cliente tomado de Cuenta en casos y Empresa en incidentes. Una mención muestra relación con el proveedor; la evidencia permite verificar la afectación.")
            st.download_button("Descargar clientes por proveedor", exportar_csv(tabla),
                               file_name=f"clientes_proveedores_{anio:04d}_{mes:02d}.csv", mime="text/csv", key="proveedores_clientes_csv")
        with tabs[2]:
            informe = informe_proveedores(datos)
            mostrar_tabla(tabla_fechas(informe))
            with st.expander("Ver historial completo"):
                elegido = st.selectbox("Caso o incidente", list(range(len(informe))),
                                       format_func=lambda i: f"{informe.iloc[i]['Caso o incidente']} · {informe.iloc[i]['Nombre del proveedor']}",
                                       key="proveedores_historial_registro")
                st.text_area("Historial disponible", informe.iloc[elegido]["Historial del caso o incidente"],
                             height=300, disabled=True, key="proveedores_historial_texto")
            columnas = ["Tipo", "Número", "Proveedor", "Cliente", "Relación", "Estado", "Prioridad", "Descripción",
                        "Escalamiento", "Primera respuesta", "Solución", "ANS Atención", "ANS Solución", "Período del registro", "Fuente", "Evidencia"]
            with st.expander("Ver detalle de ANS y evidencia"):
                mostrar_tabla(tabla_fechas(datos[columnas]))
            st.download_button("Descargar detalle y evidencia", exportar_csv(datos),
                               file_name=f"detalle_proveedores_{anio:04d}_{mes:02d}.csv", mime="text/csv", key="proveedores_detalle_csv")
    with st.expander("ANS y horarios por proveedor"):
        seleccion = proveedor if proveedor in catalogo_ans()["proveedores"] else st.selectbox("Ver acuerdo", list(catalogo_ans()["proveedores"]), key="proveedores_acuerdo")
        acuerdo = catalogo_ans()["proveedores"][seleccion]
        st.write(acuerdo["servicios"])
        st.write("Horario: " + acuerdo["horario"])
        if seleccion == "RPost / eSoluciona":
            st.caption("El Excel no especifica una excepción por festivos para RPost; se aplica el horario del día de semana.")
        mostrar_tabla(tabla_objetivos(seleccion))
        st.caption(f"Fuente: {catalogo_ans()['fuente']}, matriz detallada por prioridad. "
                   "Seguimiento periódico es un intervalo entre actualizaciones; no se calcula con las fechas de primera respuesta y solución. "
                   "Grafólogos P3/P4: solución en tres días hábiles, lunes a viernes sin festivos.")
        st.write(f"Disponibilidad contractual: {acuerdo['disponibilidad_objetivo']:.2f} %. "
                 "Su cumplimiento requiere períodos de indisponibilidad del servicio y cobertura de medición; no se deduce del número de casos.")
        st.write("Procesamiento: " + acuerdo["procesamiento"])
        st.write("Deducciones de referencia: " + acuerdo["deducciones_resumen"])
        st.caption("AutenTIC: la matriz detallada registra una hora de atención en P1 y diagnóstico N/A; se usa esa matriz para el cálculo. "
                   "Las deducciones requieren validación contractual y no se aplican automáticamente.")
    with st.expander("Cómo registrar las fechas en las notas de trabajo"):
        st.write("Registra estas líneas con fechas reales en hora de Colombia. Especifica el proveedor en cada bloque si hay más de uno.")
        st.code("Proveedor: RPost\nEscalamiento al proveedor: DD/MM/AAAA HH:MM\nPrimera respuesta del proveedor: DD/MM/AAAA HH:MM\n"
                "Diagnóstico del proveedor: DD/MM/AAAA HH:MM\nMitigación del proveedor: DD/MM/AAAA HH:MM\nSolución del proveedor: DD/MM/AAAA HH:MM", language=None)
        st.caption("La creación y el cierre del caso no sustituyen el escalamiento y la solución del proveedor. "
                   "Los datos se consultan desde la base compartida; no se vuelve a llamar a ServiceNow al cambiar filtros.")
