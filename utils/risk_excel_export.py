"""Exportación profesional en memoria del análisis de riesgos."""

from io import BytesIO

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


def _safe(value):
    if value is None or (not isinstance(value, (list, tuple, dict)) and pd.isna(value)):
        return None
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime().replace(tzinfo=None)
    return value


def _write_df(ws, df):
    columns = list(df.columns)
    ws.append(columns)
    for row in df.itertuples(index=False, name=None):
        ws.append([_safe(value) for value in row])
    ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions
    header = PatternFill("solid", fgColor="16324F")
    thin = Side(style="thin", color="D9E2F3")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF"); cell.fill = header; cell.alignment = Alignment(wrap_text=True)
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = Border(bottom=thin)
    for index, column in enumerate(columns, 1):
        values = [str(column)] + [str(value or "") for value in df[column].head(200)]
        ws.column_dimensions[get_column_letter(index)].width = min(55, max(11, max(map(len, values)) + 2))


def build_risk_workbook(analysis, year, month_from, month_to):
    wb = Workbook(); wb.remove(wb.active)
    validation = analysis["validation"]
    risk_top = analysis["risks"].iloc[0]["ID"] if not analysis["risks"].empty else "N/A"
    summary = pd.DataFrame([
        ("Periodo analizado", f"{year}-{month_from:02d} a {year}-{month_to:02d}"),
        ("Total incidentes", validation["total"]), ("Incidentes materializados", validation["risk"]),
        ("Exclusiones", validation["exclusion"]), ("Pendientes", validation["pending"]),
        ("Porcentaje conciliado", validation["percentage"]),
        ("Total riesgos materializados", int(analysis["detail"].loc[analysis["detail"]["classification_type"] == "RISK", "risk_id"].nunique())),
        ("Riesgo más recurrente", risk_top),
    ], columns=["Indicador", "Valor"])
    sheets = {
        "Resumen": summary, "Riesgos": analysis["risks"], "Conciliación": analysis["reconciliation"],
        "Exclusiones": analysis["exclusions"],
        "Detalle": analysis["detail"].rename(columns={"numero":"número", "creado_dt":"fecha", "classification_type":"clasificación", "exclusion_category":"exclusión", "classification_reason":"motivo clasificación", "confidence":"confianza"}),
        "Pendientes": analysis["pending"],
    }
    for name, df in sheets.items():
        ws = wb.create_sheet(name); _write_df(ws, df)
        if name == "Resumen":
            ws[7][1].number_format = "0.00%"
    output = BytesIO(); wb.save(output); output.seek(0); return output.getvalue()
