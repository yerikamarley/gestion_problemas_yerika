"""Lámina ejecutiva con la distribución del estándar de comportamiento de casos."""

from html import escape


MESES = ("", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")


def porcentaje(valor):
    return "Sin dato" if valor is None else f"{valor:.2f}%"


def grafico_volumen(anterior, actual, nombre_anterior, nombre_actual):
    techo = max(anterior, actual, 1)
    barras = []
    for x, valor, nombre, color in [(55, anterior, nombre_anterior, "#252b48"), (250, actual, nombre_actual, "#ee6325")]:
        alto = valor * 130 / techo
        barras.append(f'<rect x="{x}" y="{160-alto}" width="132" height="{alto}" fill="{color}"/>'
                      f'<text x="{x+66}" y="{150-alto}" text-anchor="middle" class="valor">{valor:,}</text>'
                      f'<text x="{x+66}" y="184" text-anchor="middle" class="mes">{escape(nombre)}</text>')
    return ('<svg viewBox="0 0 445 198" role="img" aria-label="Casos de soporte por período">'
            f'<title>{escape(nombre_anterior)}: {anterior}; {escape(nombre_actual)}: {actual}</title>'
            '<line x1="16" y1="160" x2="433" y2="160" stroke="#aaa"/>' + ''.join(barras) + '</svg>')


def lamina_resumen_casos(reporte, foco="", decision=""):
    a, b = reporte["metricas"]["anterior"], reporte["metricas"]["actual"]
    inicio_a, fin_a = reporte["periodos"]["anterior"]
    inicio_b, fin_b = reporte["periodos"]["actual"]
    mes_a, mes_b = MESES[inicio_a.month], MESES[inicio_b.month]
    titulo_periodo = f"{mes_a} - {mes_b} {fin_b.year}" if fin_a.year == fin_b.year else f"{mes_a} {fin_a.year} - {mes_b} {fin_b.year}"
    causas = reporte["causas"]
    if not foco.strip():
        foco = f"Principal causa agrupada: {causas.iloc[0]['Causa']}" if not causas.empty else "Sin casos de soporte registrados en el período actual"
    decision = decision.strip() or "Pendiente de definir por el responsable del reporte."
    diferencia = reporte["diferencia"]
    dif_sla = reporte["diferencia_sla"]
    delta_sla = "Sin comparación SLA" if dif_sla is None else f"{dif_sla:+.2f} p.p."
    color_sla = "#157044" if dif_sla is not None and dif_sla >= 0 else "#bd341a"
    barras = []
    for i, row in reporte["causas_lamina"].iterrows():
        causa = escape(str(row["Causa"]))
        barras.append(f'''<div class="causa"><div class="causa-titulo"><span title="{causa}">{causa}</span>
            <small><b>{int(row['Casos']):,} casos</b> · {row['Porcentaje']:.2f}%</small></div>
            <div class="pista"><div style="width:{row['Porcentaje']:.6f}%;background:{'#ee6325' if i == 0 else '#c9c5bc'}"></div></div></div>''')
    causas_html = ''.join(barras) or '<p class="vacio">Sin casos para agrupar.</p>'
    grafico = grafico_volumen(a["total"], b["total"], mes_a, mes_b)
    return f'''<!doctype html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Comportamiento de casos | {escape(titulo_periodo)}</title><style>
    *{{box-sizing:border-box}}body{{margin:0;background:#fffafa;font-family:Arial,Helvetica,sans-serif;color:#142b4c}}
    .lamina{{width:100%;max-width:1120px;min-width:960px;margin:auto;padding:30px 38px 14px;background:#fffafa}}
    .etiqueta{{display:flex;justify-content:flex-end;font-size:8px;letter-spacing:.2px;height:12px}}
    .etiqueta b{{background:#e5f2fa;padding:2px 24px}}.etiqueta span{{background:#f1f1f1;padding:2px 26px}}
    h1{{font-size:26px;line-height:1.18;color:#050505;margin:3px 0 6px;font-weight:750;letter-spacing:-.25px}}
    .subtitulo{{color:#f46022;font-weight:700;font-size:17px;letter-spacing:.4px;margin-bottom:25px}}
    .foco{{border:1px solid #fa682d;border-radius:6px;background:#fbeae2;min-height:57px;display:flex;align-items:center;justify-content:space-between;padding:7px 12px;gap:22px}}
    .foco strong{{font-size:14px;line-height:1.3;overflow-wrap:anywhere}}.diferencia{{text-align:right;flex-shrink:0}}
    .diferencia b{{font:700 25px Georgia,serif;color:#cd4a11}}.diferencia small{{display:block;font-size:10px;margin-top:2px}}
    .tarjetas{{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:13px;margin-top:14px}}
    .tarjeta{{background:white;border:1px solid #e0dcd7;border-radius:6px;padding:12px 13px 9px;box-shadow:0 2px 4px #00000015;min-height:87px}}
    .tarjeta .et{{font-size:9px;line-height:1.3}}.tarjeta .cifra{{font:700 27px Georgia,serif;margin:5px 0 3px}}
    .tarjeta .nota{{font-size:10px;line-height:1.45}}.tarjeta.naranja{{background:#ed6326;color:white;border-color:#ed6326}}
    .tarjeta.actual{{border-color:#f35c21}}.tarjeta.actual .cifra{{color:#d54c17}}.rojo{{color:#c52d19}}
    .graficas{{display:grid;grid-template-columns:1fr 1.05fr;gap:20px;margin-top:15px}}
    .panel{{background:white;box-shadow:0 2px 4px #00000015;min-width:0;height:220px}}
    .panel h2{{font:700 15px Georgia,serif;margin:14px 0 0}}.volumen{{padding:0 8px}}
    .volumen svg{{width:100%;height:183px}}.valor{{font:bold 14px Arial;fill:#30313d}}.mes{{font:14px Arial;fill:#666}}
    .causas{{border:1px solid #e0dcd7;border-radius:6px;padding:0 13px 10px}}.causas h2{{margin-bottom:10px}}
    .causa{{margin-bottom:13px}}.causa-titulo{{display:flex;align-items:center;gap:9px;justify-content:space-between;font-size:10px;margin-bottom:5px;min-height:18px}}
    .causa-titulo span{{font-weight:600;max-width:66%;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;overflow-wrap:anywhere}}
    .causa-titulo small{{font-size:9px;white-space:nowrap}}.causa-titulo b{{color:#ce4a16;font-weight:500}}
    .pista{{height:17px;background:#f0efeb;border-radius:3px;overflow:hidden}}.pista>div{{height:100%;border-radius:3px}}
    .decision{{margin:15px 20px 0 auto;width:76%;min-height:65px;background:#50318b;color:white;border-radius:6px;display:grid;grid-template-columns:180px 1fr;gap:12px;align-items:center;padding:14px 23px}}
    .decision b{{font:700 11px Georgia,serif}}.decision p{{font-size:12px;font-weight:600;line-height:1.3;margin:0;overflow-wrap:anywhere}}
    .pie{{font-size:9px;color:#676572;line-height:1.5;margin-top:11px}}.vacio{{color:#676572;font-size:12px}}
    @media print{{body{{background:white}}.lamina{{padding:18px;width:1080px}}}}
    </style></head><body><main class="lamina">
    <div class="etiqueta"><b>Etiquetado:</b><span>Interna</span></div>
    <h1>Comportamiento de casos | {escape(titulo_periodo)} (corte {fin_b.day})</h1>
    <div class="subtitulo">Impacto al cliente</div>
    <div class="foco"><strong>Foco: {escape(foco)}</strong><div class="diferencia"><b>{diferencia:+,}</b><small>casos vs. {mes_a}</small></div></div>
    <div class="tarjetas">
      <div class="tarjeta naranja"><div class="et"><b>CASOS {mes_b} (corte {fin_b.day})</b></div><div class="cifra">{b['total']:,}</div><div class="nota">{mes_a} ({fin_a.day}): {a['total']:,} casos · {diferencia:+,}</div></div>
      <div class="tarjeta"><div class="et">BACKLOG DEL PERÍODO</div><div class="cifra">{b['abiertos']:,}</div><div class="nota">casos abiertos · última carga</div></div>
      <div class="tarjeta"><div class="et">SLA {mes_a}</div><div class="cifra">{porcentaje(a['sla'])}</div><div class="nota"><span class="rojo">{a['no_cumple']}</span> no cumplen ANS · {a['evaluados']} evaluados</div></div>
      <div class="tarjeta actual"><div class="et">SLA {mes_b}</div><div class="cifra">{porcentaje(b['sla'])}</div><div class="nota"><span class="rojo">{b['no_cumple']}</span> no cumplen ANS · <b style="color:{color_sla}">{delta_sla}</b></div></div>
    </div>
    <div class="graficas"><section class="panel volumen"><h2>Casos por mes</h2>{grafico}</section>
      <section class="panel causas"><h2>Causas agrupadas · {mes_b}</h2>{causas_html}</section></div>
    <div class="decision"><b>DECISIÓN REQUERIDA</b><p>{escape(decision)}</p></div>
    <div class="pie">Solo equipo de soporte · Creación: {inicio_a:%d/%m/%Y}–{fin_a:%d/%m/%Y} vs. {inicio_b:%d/%m/%Y}–{fin_b:%d/%m/%Y}.<br>
    Estados y asignación según última carga; no reconstruye el estado histórico al corte. SLA por vencimiento oficial: {a['evaluados']}/{b['evaluados']} cerrados evaluados; {a['sin_evaluar']}/{b['sin_evaluar']} sin evaluación ({mes_a}/{mes_b}).</div>
    </main></body></html>'''
