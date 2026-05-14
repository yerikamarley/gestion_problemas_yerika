import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const workspace = process.cwd();
const outDir = path.join(workspace, "outputs", "rpost_incidentes");
const dataPath = path.join(outDir, "rpost_report_data.json");
const outputPath = path.join(outDir, "analisis_incidentes_rpost.xlsx");

const data = JSON.parse(await fs.readFile(dataPath, "utf8"));

const workbook = Workbook.create();

const colors = {
  navy: "#1F4E78",
  blue: "#D9EAF7",
  teal: "#0F766E",
  green: "#E2F0D9",
  amber: "#FFF2CC",
  red: "#FCE4D6",
  gray: "#F3F4F6",
  darkGray: "#374151",
  white: "#FFFFFF",
};

function addSheet(name) {
  const sheet = workbook.worksheets.add(name);
  sheet.showGridLines = false;
  return sheet;
}

function setTitle(sheet, range, title, subtitle = "") {
  sheet.getRange(range).merge();
  sheet.getRange(range.split(":")[0]).values = [[title]];
  sheet.getRange(range.split(":")[0]).format = {
    fill: colors.navy,
    font: { bold: true, color: colors.white, size: 16 },
    horizontalAlignment: "left",
    verticalAlignment: "center",
    rowHeightPx: 36,
  };
  if (subtitle) {
    sheet.getRange("A2:H2").merge();
    sheet.getRange("A2").values = [[subtitle]];
    sheet.getRange("A2").format = {
      fill: colors.blue,
      font: { color: colors.darkGray, size: 10 },
      wrapText: true,
      verticalAlignment: "center",
      rowHeightPx: 30,
    };
  }
}

function writeTable(sheet, startRow, startCol, headers, rows, tableName) {
  const matrix = [headers, ...rows.map((row) => headers.map((header) => row[header] ?? ""))];
  const range = sheet.getRangeByIndexes(startRow, startCol, matrix.length, headers.length);
  range.values = matrix;
  const headerRange = sheet.getRangeByIndexes(startRow, startCol, 1, headers.length);
  headerRange.format = {
    fill: colors.teal,
    font: { bold: true, color: colors.white },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
    rowHeightPx: 34,
  };
  const bodyRange = sheet.getRangeByIndexes(startRow + 1, startCol, Math.max(rows.length, 1), headers.length);
  bodyRange.format = {
    wrapText: true,
    verticalAlignment: "top",
  };
  if (rows.length > 0) {
    const address = range.address ?? "";
    if (address) {
      const table = sheet.tables.add(address, true, tableName);
      table.style = "TableStyleMedium2";
      table.showFilterButton = true;
    }
  }
  return range;
}

function setWidths(sheet, widths) {
  widths.forEach((width, idx) => {
    sheet.getRangeByIndexes(0, idx, 1, 1).format.columnWidthPx = width;
  });
}

function conditionEventTypes(sheet, rowStart, rowCount, colIndex) {
  const range = sheet.getRangeByIndexes(rowStart, colIndex, rowCount, 1);
  range.conditionalFormats.addCustom(`=ISNUMBER(SEARCH("Indisponibilidad",${range.address.split(":")[0]}))`, {
    fill: colors.red,
  });
  range.conditionalFormats.addCustom(`=ISNUMBER(SEARCH("Alerta",${range.address.split(":")[0]}))`, {
    fill: colors.green,
  });
  range.conditionalFormats.addCustom(`=ISNUMBER(SEARCH("funcional",${range.address.split(":")[0]}))`, {
    fill: colors.amber,
  });
}

const summary = addSheet("Resumen");
setTitle(
  summary,
  "A1:H1",
  "Analisis de incidentes RPOST",
  `Fuente: ${data.summary.source} | Hoja: ${data.summary.sheet} | Generado: ${data.summary.generated}`,
);
summary.freezePanes.freezeRows(4);
summary.getRange("A4:H4").values = [["Indicadores clave", "", "", "", "", "", "", ""]];
summary.getRange("A4:H4").merge();
summary.getRange("A4").format = {
  fill: colors.gray,
  font: { bold: true, color: colors.darkGray, size: 12 },
};

const kpis = [
  ["Registros fuente", data.summary.total_source_rows, "Incidentes RPOST", data.summary.total_rpost],
  ["Con ticket proveedor", data.summary.provider_ticket_count, "Ventanas mantenimiento", data.summary.maintenance_count],
  ["Duracion total oficial", data.summary.duration_total, "Duracion promedio", data.summary.duration_avg],
  ["Duracion mediana", data.summary.duration_median, "Casos con discrepancia duracion", data.summary.discrepancy_count],
];
summary.getRange("A5:D8").values = kpis;
summary.getRange("A5:D8").format = {
  fill: colors.blue,
  font: { color: colors.darkGray },
  verticalAlignment: "center",
};
summary.getRange("A5:A8").format.font = { bold: true };
summary.getRange("C5:C8").format.font = { bold: true };

writeTable(summary, 10, 0, ["Categoria", "Cantidad"], data.by_group, "ResumenGrupoEvento");
writeTable(summary, 10, 3, ["Categoria", "Cantidad"], data.by_service, "ResumenServicio");
writeTable(summary, 10, 6, ["Mes", "Cantidad"], data.by_month, "ResumenMes");

summary.getRange("A20:H20").merge();
summary.getRange("A20").values = [["Lectura ejecutiva"]];
summary.getRange("A20").format = { fill: colors.navy, font: { bold: true, color: colors.white } };
const execRows = [
  ["Hallazgo", "Detalle"],
  ["Recurrencia", "La mayoria de eventos se concentran en PORTAL.RPOST.COM - WBLM sobre nodo Agrio."],
  ["Proveedor", "Hay 8 incidentes con caso/ticket de proveedor; se recomienda solicitar RCA formal por cada uno."],
  ["Mantenimiento", "No hay ventanas de mantenimiento evidenciadas en los tickets analizados."],
  ["Calidad del dato", "Algunos casos tienen diferencia entre Duracion oficial y Creado/Cerrado; conviene validar el criterio de medicion."],
  ["Caso no disponibilidad", "INC0017041 corresponde a correo/DKIM y debe analizarse separado de indisponibilidad del portal."],
];
summary.getRange("A21:B26").values = execRows;
summary.getRange("A21:B21").format = { fill: colors.teal, font: { bold: true, color: colors.white } };
summary.getRange("A22:B26").format = { wrapText: true, verticalAlignment: "top" };
summary.getRange("A:A").format.columnWidthPx = 180;
summary.getRange("B:B").format.columnWidthPx = 440;
summary.getRange("C:C").format.columnWidthPx = 180;
summary.getRange("D:D").format.columnWidthPx = 170;
summary.getRange("G:G").format.columnWidthPx = 120;
summary.getRange("H:H").format.columnWidthPx = 110;

const detail = addSheet("Detalle RPOST");
setTitle(detail, "A1:X1", "Detalle por incidente RPOST");
const detailHeaders = [
  "Numero",
  "Servicio",
  "Categoria",
  "Prioridad",
  "Impacto",
  "Estado",
  "Tipo falla ServiceNow",
  "Creado",
  "Cerrado",
  "Duracion oficial",
  "Tiempo calendario",
  "Diferencia duracion",
  "Caso proveedor",
  "Tipo evento",
  "Causa raiz / causa probable",
  "Evidencia relevante",
  "Ventana mantenimiento",
  "Detalle ventana mantenimiento",
  "Observacion nosotros",
  "Observacion proveedor",
  "Breve descripcion",
  "Actualizaciones",
];
writeTable(detail, 2, 0, detailHeaders, data.detail, "DetalleRPOST");
setWidths(detail, [
  105, 150, 115, 105, 105, 95, 170, 140, 140, 120, 125, 190, 120, 190, 340, 330, 145, 250, 330, 350, 300, 95,
]);
detail.freezePanes.freezeRows(3);
detail.freezePanes.freezeColumns(1);
conditionEventTypes(detail, 3, data.detail.length, 13);

const us = addSheet("Obs Nosotros");
setTitle(us, "A1:C1", "Observaciones para nosotros / Certicamara");
writeTable(us, 2, 0, ["Tema", "Observacion", "Accion sugerida"], data.observations_us, "ObservacionesNosotros");
setWidths(us, [180, 500, 500]);
us.freezePanes.freezeRows(3);

const provider = addSheet("Obs Proveedor");
setTitle(provider, "A1:C1", "Observaciones y solicitudes para proveedor RPOST");
writeTable(provider, 2, 0, ["Solicitud", "Detalle", "Prioridad"], data.observations_provider, "ObservacionesProveedor");
setWidths(provider, [230, 680, 110]);
provider.freezePanes.freezeRows(3);

const maintenance = addSheet("Ventanas Mantto");
setTitle(maintenance, "A1:F1", "Validacion de ventanas de mantenimiento");
writeTable(
  maintenance,
  2,
  0,
  ["Numero", "Creado", "Caso proveedor", "Ventana mantenimiento", "Conclusion", "Requiere confirmacion proveedor"],
  data.maintenance,
  "VentanasMantenimiento",
);
setWidths(maintenance, [115, 145, 140, 165, 460, 210]);
maintenance.freezePanes.freezeRows(3);

const catalog = addSheet("Metricas");
setTitle(catalog, "A1:F1", "Metricas de soporte");
writeTable(catalog, 2, 0, ["Categoria", "Cantidad"], data.by_type, "MetricasTipoEvento");
writeTable(catalog, 2, 3, ["Categoria", "Cantidad"], data.by_failure, "MetricasTipoFalla");
setWidths(catalog, [320, 100, 40, 320, 100, 40]);

await fs.mkdir(outDir, { recursive: true });
const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 50 },
  summary: "formula error scan",
});
console.log(errors.ndjson);
const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(outputPath);
console.log(outputPath);
