"""Panel compartido de SLA de resolución para Dashboard y KPI Incidentes."""

import html
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

from services.incidentes_sla import (
    CATEGORIAS, CONSULTA, PRIORIDADES, SIN_CATEGORIA,
    agregar_evaluacion, cargar_horario, resumen_etapa, tabla_objetivos,
)


COLORES_ESTADO = {
    "Cumple": "#5d16a6", "No cumple": "#f35b04", "Dentro del plazo": "#9683ec",
    "Vencido": "#d6a21a", "No evaluable": "#c5c0cb",
}


def porcentaje_visible(valor):
    return "No calculable" if valor is None else f"{valor:.2f}%"


def motivo_no_evaluacion(base, etapa):
    estados = base[f"{etapa}_estado"]
    sin_fecha = int(estados.eq("Sin fecha de atención" if etapa == "atencion" else "Sin fecha de resolución").sum())
    nombre = "primera atención" if etapa == "atencion" else "resolución"
    motivos = []
    if sin_fecha:
        motivos.append(f"{sin_fecha} sin fecha de {nombre}")
    otros = int((~estados.isin(["Cumple", "No cumple", "Dentro del plazo", "Vencido", "Sin fecha de atención", "Sin fecha de resolución"])).sum())
    if otros:
        motivos.append(f"{otros} con fechas o clasificación por revisar")
    return " · ".join(motivos)


def grafica_cumplimiento(base, etapa):
    columna = f"{etapa}_estado"
    datos = base[["categoria_gestion", columna]].copy()
    datos["Resultado"] = datos[columna].where(datos[columna].isin(COLORES_ESTADO), "No evaluable")
    datos = datos.groupby(["categoria_gestion", "Resultado"], dropna=False).size().reset_index(name="Registros")
    tipos = [tipo for tipo in (*CATEGORIAS, SIN_CATEGORIA) if datos["categoria_gestion"].eq(tipo).any()]
    fig = px.bar(datos, x="Registros", y="categoria_gestion", color="Resultado",
                 orientation="h", text="Registros", barmode="stack",
                 category_orders={"categoria_gestion": tipos, "Resultado": list(COLORES_ESTADO)},
                 color_discrete_map=COLORES_ESTADO,
                 labels={"categoria_gestion": "Tipo de incidente"})
    fig.update_traces(textposition="inside", insidetextanchor="middle", textfont=dict(size=13, color="white"),
                      marker_line_width=0, width=.45,
                      hovertemplate="%{y}<br>%{x} registros<extra>%{fullData.name}</extra>")
    for traza in fig.data:
        if traza.name != "Cumple":
            traza.textfont.color = "#302936"
    fig.update_layout(height=max(210, len(tipos) * 64 + 46), bargap=.44,
                      margin=dict(l=0, r=8, t=10, b=32), showlegend=False,
                      font=dict(family="Arial, sans-serif", size=12, color="#6c6570"),
                      uniformtext=dict(minsize=11, mode="hide"), hoverlabel=dict(bgcolor="white", font_size=13),
                      plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)" )
    fig.update_xaxes(title=dict(text="Registros", font=dict(size=11)), nticks=6, rangemode="tozero",
                     gridcolor="#eee9f0", showgrid=True, zeroline=False, fixedrange=True)
    fig.update_yaxes(title=None, categoryorder="array", categoryarray=tipos,
                     autorange="reversed", tickmode="array", tickvals=tipos,
                     ticktext=[tipo.replace("Incidente Seguridad", "Incidente<br>Seguridad")
                               .replace("Consulta / Ayuda", "Consulta /<br>Ayuda") for tipo in tipos],
                     showgrid=False, zeroline=False, fixedrange=True, automargin=True)
    return fig


def tabla_resumen_categorias(base):
    filas = []
    for categoria, grupo in base.groupby("categoria_gestion", sort=False):
        resolucion = resumen_etapa(grupo, "resolucion")
        filas.append({"Categoría": categoria, "Registros": len(grupo),
                      "Resolución %": resolucion["porcentaje"], "Evaluados resolución": resolucion["evaluados"],
                      "No evaluables resolución": resolucion["sin_datos"], "Pendientes resolución": resolucion["pendientes"],
                      "Vencidos resolución": resolucion["vencidos"]})
    return pd.DataFrame(filas)


def tabla_tendencia(base):
    datos = base.copy()
    datos["Mes"] = pd.to_datetime(datos.get("creado"), errors="coerce").dt.strftime("%Y-%m")
    filas = []
    for (mes, categoria), grupo in datos.dropna(subset=["Mes"]).groupby(["Mes", "categoria_gestion"]):
        for etapa, etiqueta in (("resolucion", "Resolución"),):
            resumen = resumen_etapa(grupo, etapa)
            filas.append({"Mes": mes, "Categoría": categoria, "Etapa": etiqueta,
                          "Cumplimiento %": resumen["porcentaje"], "Evaluados": resumen["evaluados"],
                          "Cumplen": resumen["cumple"], "No cumplen": resumen["no_cumple"],
                          "Sin datos": resumen["sin_datos"], "Pendientes": resumen["pendientes"]})
    return pd.DataFrame(filas)


def tabla_detalle(base):
    detalle = base.copy()
    detalle["resolucion_horas_soporte"] = pd.to_numeric(detalle["resolucion_segundos"], errors="coerce") / 3600
    columnas = {
        "numero": "Incidente", "estado": "Estado", "categoria_gestion": "Categoría de gestión",
        "resolucion_estado": "Resultado resolución", "resolucion_motivo": "Por qué no se evalúa resolución",
        "categoria": "Categoría de origen",
        "prioridad": "Prioridad de origen", "prioridad_gestion": "Prioridad de la guía",
        "validacion_clasificacion": "Validación pendiente", "coordinador_gestion": "Cargo que coordina",
        "servicio_negocio": "Servicio", "asignado_a": "Asignado a",
        "inicio_sla": "Inicio del reloj", "fuente_inicio_sla": "Fuente del inicio",
        "fecha_resolucion": "Resolución",
        "objetivo_resolucion_horas": "Objetivo resolución (h soporte)",
        "resolucion_horas_soporte": "Resolución (h soporte)", "resolucion_vencimiento": "Vencimiento resolución",
        "cerrado": "Cierre administrativo",
    }
    return detalle[[col for col in columnas if col in detalle]].rename(columns=columnas)


def render_panel_incidentes_ans(df, key_prefix, titulo, periodo=None, presentacion=None):
    css = (Path(__file__).resolve().parents[1] / "assets" / "incidentes.css").read_text(encoding="utf-8")
    st.markdown(f"<style>{css}</style>", unsafe_allow_html=True)
    with st.container(key=f"incidentes_panel_{key_prefix}"):
        return _render_panel(df, key_prefix, titulo, periodo, presentacion)


def _metas_html(metas, categoria):
    def corto(valor):
        return html.escape(str(valor).replace("minutos", "min"))

    if categoria in CATEGORIAS:
        tarjetas = []
        for fila in metas.to_dict("records"):
            tarjetas.append('<div class="inc-ans-priority">'
                            f'<div class="inc-ans-priority-name">{corto(fila["Prioridad"])}</div>'
                            '<div class="inc-ans-priority-values">'
                            f'<div><b>{corto(fila["Resolución máxima"])}</b><small>Resolución</small></div>'
                            '</div></div>')
        clase = "inc-ans-priority-goals single" if len(tarjetas) == 1 else "inc-ans-priority-goals"
        return f'<div class="{clase}">' + "".join(tarjetas) + '</div>'
    tarjetas = []
    for indice, (nombre, grupo) in enumerate(metas.groupby("Categoría", sort=False), 1):
        filas = ['<div class="inc-ans-goal-row header"><span>Prioridad</span><span>Resolución</span></div>']
        for fila in grupo.to_dict("records"):
            filas.append('<div class="inc-ans-goal-row">'
                         f'<span>{corto(fila["Prioridad"])}</span>'
                         f'<span>{corto(fila["Resolución máxima"])}</span></div>')
        tarjetas.append('<div class="inc-ans-goal-category"><div class="inc-ans-goal-name">'
                        f'<span class="inc-ans-goal-index">{indice:02}</span>{html.escape(nombre)}</div>'
                        + "".join(filas) + '</div>')
    return '<div class="inc-ans-goals">' + "".join(tarjetas) + '</div>'


def _render_indicadores(base, resolucion, key_prefix):
    indicadores = [
        ("registros", "Registros", len(base), "En la categoría y período seleccionados"),
        ("abiertos", "Abiertos", int((~base["finalizado_gestion"].astype(bool)).sum()),
         f"{resolucion['vencidos']} con plazo de resolución vencido"),
        ("resolucion", "SLA de resolución", porcentaje_visible(resolucion["porcentaje"]),
         f"{resolucion['evaluados']} evaluados" + (" · " + motivo_no_evaluacion(base, "resolucion") if resolucion["sin_datos"] else "")),
    ]
    for col, (clave, etiqueta, valor, nota) in zip(st.columns(len(indicadores), gap="small"), indicadores):
        with col.container(border=True, key=f"incidentes_stat_{key_prefix}_{clave}"):
            st.metric(etiqueta, valor)
            st.caption(nota)


def _render_grafica(base, etapa, key_prefix):
    resumen = resumen_etapa(base, etapa)
    titulo, subtitulo = (("Atención por tipo de incidente", "Primer contacto por categoría") if etapa == "atencion"
                         else ("Resolución por tipo de incidente", "Cumplimiento del SLA y pendientes por categoría"))
    with st.container(border=True, key=f"incidentes_chart_{key_prefix}_{etapa}"):
        st.markdown('<div class="inc-ans-chart-heading"><div>'
                    f'<strong>{titulo}</strong><small>{subtitulo}</small></div>'
                    f'<span class="inc-ans-chart-result">{porcentaje_visible(resumen["porcentaje"])}</span></div>',
                    unsafe_allow_html=True)
        if resumen["evaluados"] == 0 and resumen["sin_datos"] == len(base):
            faltantes = html.escape(motivo_no_evaluacion(base, etapa))
            explicacion = ("Se necesita saber cuándo comenzó la primera atención. La creación, el cierre y los correos automáticos no indican esa fecha."
                           if etapa == "atencion" else "Se necesita la fecha de resolución para comparar el tiempo con el plazo de la guía.")
            if not base[f"{etapa}_estado"].isin(["Sin fecha de atención", "Sin fecha de resolución"]).all():
                explicacion = "Hay fechas o clasificaciones que deben revisarse. El detalle indica el motivo específico por caso."
            st.markdown('<div class="inc-ans-unavailable">'
                        '<strong>Este porcentaje todavía no se puede calcular</strong>'
                        f'<p>{faltantes}.</p><p>{explicacion}</p>'
                        '<small>Los incidentes sí tienen información. En «Detalle de registros» puedes ver el motivo por caso.</small></div>',
                        unsafe_allow_html=True)
            return
        else:
            st.plotly_chart(grafica_cumplimiento(base, etapa), use_container_width=True,
                            theme=None, config={"displayModeBar": False, "scrollZoom": False},
                            key=f"{key_prefix}_grafica_{etapa}")
        causa = motivo_no_evaluacion(base, etapa)
        pendientes = (f'<br>{resumen["pendientes"]} pendientes, de los cuales {resumen["vencidos"]} están vencidos'
                      if etapa == "resolucion" else '')
        st.markdown(f'<div class="inc-ans-chart-footer">{resumen["cumple"]} cumplen · '
                    f'{resumen["no_cumple"]} no cumplen' + (f' · {html.escape(causa)}' if causa else '') + pendientes + '</div>',
                    unsafe_allow_html=True)


def _render_leyenda(base):
    estados = set(base["resolucion_estado"])
    resultados = {e if e in COLORES_ESTADO else "No evaluable" for e in estados}
    items = [f'<span class="inc-ans-legend-item"><i style="background:{color}"></i>{estado}</span>'
             for estado, color in COLORES_ESTADO.items() if estado in resultados]
    st.markdown('<div class="inc-ans-legend">' + "".join(items) + '</div>', unsafe_allow_html=True)


MESES = ("enero", "febrero", "marzo", "abril", "mayo", "junio",
         "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre")


def nombre_mes(mes):
    fecha = pd.Timestamp(f"{mes}-01")
    return f"{MESES[fecha.month - 1].capitalize()} {fecha.year}"


def grafica_mensual(tendencia):
    datos = tendencia.sort_values(["Mes", "Categoría"]).copy()
    meses = sorted(datos["Mes"].unique())
    un_mes = len(meses) == 1
    tipos = [tipo for tipo in (*CATEGORIAS, SIN_CATEGORIA) if datos["Categoría"].eq(tipo).any()]
    datos["Período"] = datos["Mes"].map(nombre_mes)
    datos["Porcentaje"] = datos["Cumplimiento %"].map(lambda valor: f"{valor:.2f}%" if pd.notna(valor) else "")
    colores = {CATEGORIAS[0]: "#f35b04", CATEGORIAS[1]: "#9683ec", CATEGORIAS[2]: "#d6a21a", SIN_CATEGORIA: "#c5c0cb"}
    fig = px.bar(datos, x="Cumplimiento %" if un_mes else "Período",
                 y="Categoría" if un_mes else "Cumplimiento %", color="Categoría",
                 orientation="h" if un_mes else "v", barmode="relative" if un_mes else "group", text="Porcentaje",
                 custom_data=["Período", "Evaluados", "Cumplen", "No cumplen", "Pendientes", "Sin datos"],
                 category_orders={"Categoría": tipos,
                                  "Período": [nombre_mes(mes) for mes in meses]},
                 color_discrete_map=colores)
    fig.update_traces(textposition="outside", cliponaxis=False, marker_line_width=0,
                      textfont=dict(size=12, color="#302936"),
                      hovertemplate="%{customdata[0]}<br>%{fullData.name}<br>%{text} dentro del plazo"
                                    "<br>%{customdata[2]} de %{customdata[1]} evaluados cumplen"
                                    "<br>%{customdata[3]} fuera del plazo · %{customdata[4]} abiertos"
                                    "<br>%{customdata[5]} no evaluables<extra></extra>")
    fig.update_layout(height=max(240, datos["Categoría"].nunique() * 62 + 50) if un_mes else 330,
                      bargap=.45, showlegend=not un_mes,
                      font=dict(family="Arial, sans-serif", size=12, color="#6c6570"),
                      paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                      margin=dict(l=0, r=28, t=24, b=35),
                      legend=dict(title=None, orientation="h", y=-.25, font=dict(size=11)))
    porcentaje = dict(range=[0, 115], title="Resueltos dentro del plazo (%)", ticksuffix="%",
                      tickvals=[0, 25, 50, 75, 100], gridcolor="#eee9f0", zeroline=False,
                      fixedrange=True, automargin=True)
    if un_mes:
        fig.update_xaxes(**porcentaje)
        fig.update_yaxes(title=None, autorange="reversed", categoryorder="array", categoryarray=tipos,
                         showgrid=False, fixedrange=True,
                         automargin=True, tickmode="array", tickvals=tipos,
                         ticktext=[tipo.replace("Incidente Seguridad", "Incidente<br>Seguridad")
                                   .replace("Consulta / Ayuda", "Consulta /<br>Ayuda") for tipo in tipos])
    else:
        fig.update_yaxes(**porcentaje)
        fig.update_xaxes(title=None, type="category", showgrid=False, fixedrange=True)
    return fig


def _render_tendencia(tendencia, key_prefix):
    if tendencia.empty:
        st.info("No hay registros con fecha de creación para mostrar el resumen mensual.")
        return
    meses = sorted(tendencia["Mes"].unique())
    un_mes = len(meses) == 1
    with st.container(border=True, key=f"incidentes_chart_{key_prefix}_tendencia_resolucion"):
        st.markdown("**" + (f"SLA de resolución · {nombre_mes(meses[0])}" if un_mes
                             else "Comparación mensual del SLA de resolución") + "**")
        st.caption("Cada barra muestra qué porcentaje de los registros evaluados se resolvió dentro del plazo, por tipo.")
        if un_mes:
            st.caption("Hay un solo mes seleccionado: se comparan los tipos. Selecciona todos los meses para ver los cambios mes a mes.")
        if tendencia["Cumplimiento %"].notna().any():
            st.plotly_chart(grafica_mensual(tendencia), use_container_width=True, theme=None,
                            config={"displayModeBar": False}, key=f"{key_prefix}_tendencia_resolucion")
        else:
            st.info("No hay resoluciones evaluables para calcular porcentajes. La tabla muestra los registros disponibles.")
        if un_mes:
            for fila in tendencia.to_dict("records"):
                tipo = html.escape(str(fila["Categoría"]))
                st.markdown(f"**{tipo}: {porcentaje_visible(fila['Cumplimiento %']) if pd.notna(fila['Cumplimiento %']) else 'No calculable'}** — "
                            f"{fila['Cumplen']} de {fila['Evaluados']} evaluados dentro del plazo; "
                            f"{fila['No cumplen']} fuera del plazo; {fila['Pendientes']} abiertos.")
    detalle = tendencia.drop(columns=["Etapa"]).copy()
    detalle["Mes"] = detalle["Mes"].map(nombre_mes)
    st.dataframe(detalle, hide_index=True, use_container_width=True,
                 column_config={"Categoría": st.column_config.TextColumn("Tipo de incidente"),
                                "Cumplimiento %": st.column_config.NumberColumn("SLA resolución", format="%.2f%%"),
                                "Cumplen": st.column_config.NumberColumn("Dentro del plazo"),
                                "No cumplen": st.column_config.NumberColumn("Fuera del plazo"),
                                "Sin datos": st.column_config.NumberColumn("No evaluables"),
                                "Pendientes": st.column_config.NumberColumn("Abiertos")})
    st.caption("El porcentaje excluye abiertos y registros sin resolución evaluable. Un porcentaje ausente no equivale a 0 %.")


def _render_panel(df, key_prefix, titulo, periodo, presentacion):
    horario = cargar_horario()
    base = agregar_evaluacion(df, horario=horario)
    cabecera, horario_col = st.columns([2, 1], vertical_alignment="center")
    with cabecera:
        st.markdown('<div class="inc-ans-eyebrow">GESTIÓN DE INCIDENTES</div>', unsafe_allow_html=True)
        st.subheader(titulo)
        if periodo:
            st.markdown(f'<div class="inc-ans-period">Período por creación · {html.escape(str(periodo))}</div>', unsafe_allow_html=True)
    with horario_col:
        st.markdown('<div class="inc-ans-schedule"><span class="inc-ans-badge">L–V · 08:00–17:00</span>'
                    '<span class="inc-ans-badge soft">1 día = 9 h de soporte</span></div>', unsafe_allow_html=True)
    categorias = ["Todas las categorías", *CATEGORIAS]
    if base["categoria_gestion"].eq(SIN_CATEGORIA).any():
        categorias.append(SIN_CATEGORIA)
    diapositiva = False
    with st.container(border=True, key=f"incidentes_filters_{key_prefix}"):
        filtros = st.columns([1.45, 1, .75] if presentacion else [1.45, 1], vertical_alignment="center")
        with filtros[0]:
            categoria = st.selectbox("Categoría de gestión", categorias, key=f"{key_prefix}_categoria")
        if categoria != categorias[0]:
            base = base[base["categoria_gestion"].eq(categoria)].copy()
        prioridades = [p for p in (*PRIORIDADES, "Sin prioridad") if base["prioridad_gestion"].eq(p).any()]
        with filtros[1]:
            elegidas = st.multiselect("Prioridad", prioridades, key=f"{key_prefix}_prioridades", placeholder="Todas las prioridades")
        if elegidas:
            base = base[base["prioridad_gestion"].isin(elegidas)].copy()
        if presentacion:
            with filtros[2]:
                diapositiva = st.toggle("Formato diapositiva 16:9", key=f"{key_prefix}_diapositiva")
    metas = tabla_objetivos(horario).drop(columns=["Atención máxima"])
    if categoria != categorias[0]:
        metas = metas[metas["Categoría"].eq(categoria)]
    if elegidas:
        metas = metas[metas["Prioridad"].replace({"Crítico": "Critico"}).isin(elegidas)]
    if categoria in CATEGORIAS:
        coordinador = metas["Cargo que coordina"].iloc[0] if not metas.empty else "Consultar tabla de tiempos"
        st.markdown(f'<div class="inc-ans-coordinator"><b>Coordina:</b> {html.escape(coordinador)}</div>', unsafe_allow_html=True)
    if base.empty:
        st.info("No hay registros para esta categoría y prioridad en el período seleccionado.")
        st.dataframe(metas, hide_index=True, use_container_width=True)
        return base
    if presentacion is not None and diapositiva:
        presentacion(base, metas, periodo)
        return base
    resolucion = resumen_etapa(base, "resolucion")
    _render_indicadores(base, resolucion, key_prefix)
    pendientes_clasificacion = int(base["validacion_clasificacion"].ne("").sum())
    if pendientes_clasificacion:
        st.warning(f"{pendientes_clasificacion} registros requieren validar categoría o prioridad de origen. Consulta el detalle.")
    referencia = ('<div class="inc-ans-section"><strong>Plazos de resolución</strong>'
                  '<span>Máximos por prioridad · SER-PR-001</span></div>' + _metas_html(metas, categoria))
    if categoria in CATEGORIAS:
        st.markdown(referencia, unsafe_allow_html=True)
    if categoria == CONSULTA:
        st.caption("Consulta / Ayuda: prioridad Moderado y resolución máxima de 4 días de soporte.")
    resumen_tab, tendencia_tab, detalle_tab = st.tabs(["Cumplimiento", "SLA mensual", "Detalle de registros"])
    with resumen_tab:
        _render_leyenda(base)
        _render_grafica(base, "resolucion", key_prefix)
        if categoria not in CATEGORIAS:
            st.markdown(referencia, unsafe_allow_html=True)
        with st.expander("Resumen del SLA por tipo de incidente"):
            st.dataframe(tabla_resumen_categorias(base), hide_index=True, use_container_width=True,
                         column_config={
                             "Categoría": st.column_config.TextColumn("Tipo de incidente"),
                             "Resolución %": st.column_config.NumberColumn("SLA resolución", format="%.2f%%"),
                             "Evaluados resolución": st.column_config.NumberColumn("Evaluados"),
                             "No evaluables resolución": st.column_config.NumberColumn("No evaluables"),
                             "Pendientes resolución": st.column_config.NumberColumn("Abiertos"),
                             "Vencidos resolución": st.column_config.NumberColumn("Vencidos"),
                         })
    with tendencia_tab:
        tendencia = tabla_tendencia(base)
        _render_tendencia(tendencia, key_prefix)
    with detalle_tab:
        detalle = tabla_detalle(base)
        st.dataframe(detalle, height=420, hide_index=True, use_container_width=True)
        # Neutralizar fórmulas en los valores textuales de la exportación.
        exportacion = detalle.map(lambda valor: "'" + valor if isinstance(valor, str) and valor.startswith(("=", "+", "-", "@")) else valor)
        st.download_button("Descargar detalle de tiempos (CSV)", exportacion.to_csv(index=False).encode("utf-8-sig"),
                           file_name=f"{key_prefix}_tiempos.csv", mime="text/csv", key=f"{key_prefix}_descarga")
    with st.expander("Horario y criterios de medición"):
        st.caption("Soporte: lunes a viernes, 8:00 a. m. a 5:00 p. m. (Colombia). "
                   "4 días = 36 horas de soporte; 8 días = 72 horas. "
                   + (f"Festivos excluidos: {len(horario['festivos'])}." if horario["festivos"]
                      else "Sin pausa de almuerzo ni festivos configurados."))
        st.dataframe(metas, hide_index=True, use_container_width=True)
        st.caption("Cumplimiento = cumplidos / evaluados. Los abiertos corresponden a registros creados en el período seleccionado.")
        st.caption("Se compara el tiempo exacto en segundos de soporte con el máximo de la tabla; el límite está incluido. "
                   "La resolución requiere una fecha explícita. Última actualización y cierre administrativo "
                   "no sustituyen esa evidencia. Cuando falta la notificación/apertura se identifica la creación como inicio.")
    return base
