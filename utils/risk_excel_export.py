"""Generación en memoria del informe ejecutivo y anexos auditables."""
from io import BytesIO
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

NAVY, BLUE, WHITE, GREEN, RED = "17365D", "2F75B5", "FFFFFF", "70AD47", "C00000"
THIN = Side(style="thin", color="D9E2F3")
MONTHS = {1:"Ene",2:"Feb",3:"Mar",4:"Abr",5:"May",6:"Jun",7:"Jul",8:"Ago",9:"Sep",10:"Oct",11:"Nov",12:"Dic"}

def _safe(value):
    if value is None or (not isinstance(value, (list, tuple, dict)) and pd.isna(value)): return None
    return value.to_pydatetime().replace(tzinfo=None) if isinstance(value, pd.Timestamp) else value

def _header(ws, row, start, end, fill=NAVY):
    for col in range(start, end + 1):
        cell = ws.cell(row, col); cell.font = Font(bold=True, color=WHITE)
        cell.fill = PatternFill("solid", fgColor=fill)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=THIN)

def _body(ws, first, last, start, end):
    for row in ws.iter_rows(min_row=first, max_row=max(first, last), min_col=start, max_col=end):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True); cell.border = Border(bottom=THIN)

def _write_df(ws, df):
    columns = list(df.columns); ws.append(columns)
    for row in df.itertuples(index=False, name=None): ws.append([_safe(v) for v in row])
    if columns:
        _header(ws, 1, 1, len(columns)); _body(ws, 2, ws.max_row, 1, len(columns))
        ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{max(1, ws.max_row)}"
    ws.freeze_panes = "A2"; ws.sheet_view.showGridLines = False
    for index, column in enumerate(columns, 1):
        values = [str(column)] + [str(v or "") for v in df[column].head(200)]
        ws.column_dimensions[get_column_letter(index)].width = min(55, max(11, max(map(len, values)) + 2))
        if any(token in str(column).casefold() for token in ("fecha", "date", "created", "updated", "changed", "modified", "event_start", "event_end")):
            for cell in ws.iter_cols(min_col=index, max_col=index, min_row=2, max_row=ws.max_row):
                for item in cell: item.number_format = "yyyy-mm-dd hh:mm"

def _treatment(row):
    problem = row.get("Problemas asociados") or "Sin problema asociado"
    state = row.get("Estado del problema") or "Sin problema asociado"
    cause = row.get("Estado causa raíz") or "Pendiente de investigación"
    action = "Seguimiento mediante problema relacionado." if problem != "Sin problema asociado" else "Evaluar creación de problema y plan de trabajo."
    return f"Causa raíz: {cause}\nProblema asociado: {problem}\nEstado: {state}\nAcción: {action}"

def _executive(wb, analysis, year, month_from, month_to):
    ws = wb.create_sheet("Informe Ejecutivo"); ws.sheet_view.showGridLines = False
    period = f"Año {year}" if (month_from, month_to) == (1, 12) else f"{year} · {MONTHS[month_from]}-{MONTHS[month_to]}"
    ws.merge_cells("A1:H1"); ws["A1"] = "Riesgos Materializados, Volumetría, Responsabilidades y Tratamiento"
    ws["A1"].font = Font(size=16, bold=True, color=WHITE); ws["A1"].fill = PatternFill("solid", fgColor=NAVY)
    ws["A1"].alignment = Alignment(horizontal="center"); ws.row_dimensions[1].height = 30
    validation = analysis["validation"]
    ws.merge_cells("A2:H2"); ws["A2"] = f"Periodo: {period} | Fuente dinámica | {validation['total']} INC únicos"
    ws["A2"].alignment = Alignment(horizontal="center"); ws["A2"].font = Font(italic=True)
    ws["A4"] = "Tabla 1: Matriz de trazabilidad de riesgos materializados"; ws["A4"].font = Font(size=12, bold=True, color=NAVY)
    risk_count = int(analysis["risks"]["ID"].nunique()) if not analysis["risks"].empty else 0
    ws["A5"] = f"Total de Riesgos Materializados en el periodo: {risk_count}"; ws["A5"].font = Font(bold=True)
    headers = ["ID","Riesgo Materializado","Dueño del Riesgo","Impacto Escala","Cantidad tickets asociados","Tickets Asociados","Asignación Operativa RACI","Estado de Mejora, Causa Raíz y Problema"]
    for col, value in enumerate(headers, 1): ws.cell(7, col, value)
    _header(ws, 7, 1, 8); row_no = 8
    month_cols = [MONTHS[m] for m in range(month_from, month_to + 1)]
    for _, row in analysis["risks"].iterrows():
        volume = [f"Total periodo: {int(row['Cantidad tickets asociados'])}"]
        volume += [f"{m}: {int(row.get(m, 0) or 0)}" for m in month_cols]
        volume += [f"Eventos reales: {int(row.get('Eventos reales', 0) or 0)}", f"Estado: {row['Estado']}"]
        values = [row["ID"],row["Riesgo Materializado"],row["Dueño del Riesgo"],row["Impacto Escala"],"\n".join(volume),row["Tickets Asociados"],row["Asignación Operativa RACI"],_treatment(row)]
        for col, value in enumerate(values, 1): ws.cell(row_no, col, _safe(value))
        ws.row_dimensions[row_no].height = min(180, max(115, 25 + len(str(row["Tickets Asociados"])) // 9)); row_no += 1
    _body(ws, 8, row_no - 1, 1, 8); row_no += 2
    ws.cell(row_no, 1, "Tabla 2: Conciliación matemática de incidentes"); ws.cell(row_no, 1).font = Font(size=12, bold=True, color=NAVY); row_no += 1
    for col, value in enumerate(["Categoría de los Tickets","Cantidad tickets asociados","Impacto en Matriz"], 1): ws.cell(row_no, col, value)
    _header(ws, row_no, 1, 3); row_no += 1
    rows = [(f"A. Riesgos Materializados Efectivos ({risk_count} riesgos)",validation["risk"],"Sí"),("B. Falsos Positivos y Exclusiones",validation["exclusion"],"No"),("C. Pendientes de clasificación",validation["pending"],"Pendiente de análisis"),("Total Incidentes Registrados en la Base",validation["total"],"100% conciliado" if validation["reconciled"] else "ERROR DE CONCILIACIÓN")]
    start = row_no
    for values in rows:
        for col, value in enumerate(values, 1): ws.cell(row_no, col, value)
        row_no += 1
    _body(ws, start, row_no - 1, 1, 3); ws.cell(row_no - 1, 3).font = Font(bold=True, color=GREEN if validation["reconciled"] else RED); row_no += 2
    ws.cell(row_no, 1, "Tabla 3: Categorías de Exclusión (Falsos Positivos)"); ws.cell(row_no, 1).font = Font(size=12, bold=True, color=NAVY); row_no += 1
    for col, value in enumerate(["Categoría de Exclusión","Cantidad tickets asociados","Justificación Metodológica y Ejemplos reales"], 1): ws.cell(row_no, col, value)
    _header(ws, row_no, 1, 3); row_no += 1; start = row_no
    for values in analysis["exclusions"].itertuples(index=False, name=None):
        for col, value in enumerate(values, 1): ws.cell(row_no, col, _safe(value))
        ws.row_dimensions[row_no].height = 75; row_no += 1
    _body(ws, start, row_no - 1, 1, 3)
    for i, width in enumerate([36,46,29,18,25,58,32,45], 1): ws.column_dimensions[get_column_letter(i)].width = width
    ws.freeze_panes = "A8"; ws.auto_filter.ref = f"A7:H{max(7, 7 + len(analysis['risks']))}"
    ws.page_setup.orientation = "landscape"; ws.page_setup.fitToWidth = 1; ws.sheet_properties.pageSetUpPr.fitToPage = True

def build_risk_workbook(analysis, year, month_from, month_to):
    wb = Workbook(); wb.remove(wb.active); _executive(wb, analysis, year, month_from, month_to)
    v = analysis["validation"]; top = analysis["risks"].iloc[0]["ID"] if not analysis["risks"].empty else "N/A"
    summary = pd.DataFrame([("Periodo analizado",f"{year}-{month_from:02d} a {year}-{month_to:02d}"),("Total incidentes únicos",v["total"]),("Incidentes materializados",v["risk"]),("Exclusiones",v["exclusion"]),("Pendientes",v["pending"]),("Porcentaje conciliado",v["percentage"]),("Riesgos materializados",analysis["risks"]["ID"].nunique() if not analysis["risks"].empty else 0),("Eventos materiales",len(analysis.get("events",pd.DataFrame()))),("Riesgo más recurrente",top)], columns=["Indicador","Valor"])
    wanted = ["numero","creado_dt","breve_descripcion","descripcion","categoria","servicio_negocio","prioridad","estado","grupo_asignacion","classification_type","risk_id","exclusion_category","classification_reason","confidence","classification_source","automatic_classification_type","automatic_risk_id","modified_by","modified_at","change_reason"]
    detail = analysis["detail"][[c for c in wanted if c in analysis["detail"]]].copy()
    empty_links = pd.DataFrame(columns=["risk_id","problem_number","link_status","notes","created_by","created_at","updated_at","declaracion_problema","problem_status","problem_priority","problem_owner"])
    empty_history = pd.DataFrame(columns=["risk_id","problem_number","action","notes","changed_by","changed_at"])
    sheets = {"Resumen":summary,"Riesgos":analysis["risks"],"Eventos":analysis.get("events",pd.DataFrame()),"Problemas Asociados":analysis.get("problem_links",empty_links) if not analysis.get("problem_links",pd.DataFrame()).empty else empty_links,"Conciliación":analysis["reconciliation"],"Exclusiones":analysis["exclusions"],"Detalle Incidentes":detail,"Pendientes":analysis["pending"],"Historial":analysis.get("link_history",empty_history) if not analysis.get("link_history",pd.DataFrame()).empty else empty_history}
    for name, df in sheets.items(): _write_df(wb.create_sheet(name), df)
    wb["Resumen"]["B7"].number_format = "0.00%"
    output = BytesIO(); wb.save(output); output.seek(0); return output.getvalue()
