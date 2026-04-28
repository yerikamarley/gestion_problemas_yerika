import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const workspace = "C:\\Users\\yerik\\OneDrive\\Desktop\\gestion_problemas_yerika";
const inputJson = path.join(workspace, "outputs", "tablero_casos", "case_data.json");
const outputDir = path.join(workspace, "outputs", "tablero_casos");
const outputXlsx = path.join(outputDir, "tablero_tipificacion_tiempos_resolucion.xlsx");

const payload = JSON.parse(await fs.readFile(inputJson, "utf8"));
const rows = payload.rows;
const originalHeaders = payload.columns;

const addedHeaders = [
  "Tipificacion sugerida",
  "Descripcion tipificacion",
  "Criterio aplicado",
  "Tiempo resolucion (horas)",
  "Tiempo resolucion (dias)",
  "Rango tiempo resolucion",
  "Mes creacion",
  "Resultado",
  "Referencia 48h",
];

const headers = [...originalHeaders, ...addedHeaders];
const totalRows = rows.length;
const baseSheetName = "Base tipificada";

const palette = {
  navy: "#17324D",
  teal: "#0F766E",
  sky: "#DDF2F0",
  amber: "#F59E0B",
  amberSoft: "#FEF3C7",
  redSoft: "#FEE2E2",
  greenSoft: "#DCFCE7",
  gray: "#F3F4F6",
  border: "#CBD5E1",
  text: "#111827",
  muted: "#64748B",
  white: "#FFFFFF",
};

function excelCol(index) {
  let n = index + 1;
  let col = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    col = String.fromCharCode(65 + rem) + col;
    n = Math.floor((n - 1) / 26);
  }
  return col;
}

function toExcelSerial(value) {
  if (!value || typeof value !== "string") return value;
  const match = value.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})/);
  if (!match) return value;
  const [, y, m, d, hh, mm, ss] = match.map(Number);
  const utcDate = Date.UTC(y, m - 1, d, hh, mm, ss);
  const excelEpoch = Date.UTC(1899, 11, 30);
  return (utcDate - excelEpoch) / 86400000;
}

function toDateOrValue(header, value) {
  if (value === null || value === undefined) return null;
  if (["Actualizado", "Creado", "Cerrado"].includes(header)) {
    return toExcelSerial(value);
  }
  return value;
}

function average(values) {
  const nums = values.filter((v) => typeof v === "number" && Number.isFinite(v));
  if (!nums.length) return 0;
  return nums.reduce((a, b) => a + b, 0) / nums.length;
}

function median(values) {
  const nums = values.filter((v) => typeof v === "number" && Number.isFinite(v)).sort((a, b) => a - b);
  if (!nums.length) return 0;
  const mid = Math.floor(nums.length / 2);
  return nums.length % 2 ? nums[mid] : (nums[mid - 1] + nums[mid]) / 2;
}

function groupSummary(keyName, valueName = "_tiempo_resolucion_dias") {
  const map = new Map();
  for (const row of rows) {
    const key = row[keyName] || "Sin dato";
    if (!map.has(key)) map.set(key, []);
    map.get(key).push(row[valueName]);
  }
  return [...map.entries()]
    .map(([key, values]) => ({
      key,
      count: values.length,
      avgDays: Math.round(average(values) * 100) / 100,
      pct: totalRows ? values.length / totalRows : 0,
    }))
    .sort((a, b) => b.count - a.count || a.key.localeCompare(b.key));
}

function countBy(keyName) {
  const map = new Map();
  for (const row of rows) {
    const key = row[keyName] || "Sin dato";
    map.set(key, (map.get(key) || 0) + 1);
  }
  return [...map.entries()].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]));
}

function setTitle(range, value) {
  range.values = [[value]];
  range.format = {
    fill: palette.navy,
    font: { bold: true, color: palette.white, size: 16 },
    horizontalAlignment: "left",
    verticalAlignment: "center",
  };
  range.format.rowHeightPx = 34;
}

function styleTableHeader(range, fill = palette.teal) {
  range.format = {
    fill,
    font: { bold: true, color: palette.white },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
  };
  range.format.rowHeightPx = 28;
}

function styleSection(range, value) {
  range.values = [[value]];
  range.format = {
    fill: palette.sky,
    font: { bold: true, color: palette.navy },
    horizontalAlignment: "left",
    verticalAlignment: "center",
  };
}

function writeSimpleTable(sheet, startCell, headersRow, dataRows) {
  const startCol = startCell.match(/[A-Z]+/)[0];
  const startRow = Number(startCell.match(/[0-9]+/)[0]);
  const startIndex = colToIndex(startCol);
  const endCol = excelCol(startIndex + headersRow.length - 1);
  const endRow = startRow + dataRows.length;
  const range = sheet.getRange(`${startCell}:${endCol}${endRow}`);
  range.values = [headersRow, ...dataRows];
  styleTableHeader(sheet.getRange(`${startCell}:${endCol}${startRow}`));
  return { range, endRow, endCol };
}

function colToIndex(col) {
  let n = 0;
  for (const char of col) {
    n = n * 26 + (char.charCodeAt(0) - 64);
  }
  return n - 1;
}

const workbook = Workbook.create();
const dashboard = workbook.worksheets.add("Tablero");
const base = workbook.worksheets.add(baseSheetName);
const rules = workbook.worksheets.add("Reglas tipificacion");

for (const sheet of [dashboard, base, rules]) {
  sheet.showGridLines = false;
}

const dataValues = rows.map((row) => {
  const original = originalHeaders.map((h) => toDateOrValue(h, row[h]));
  const added = [
    row._tipificacion,
    row._tipificacion_descripcion,
    row._criterio_tipificacion,
    null,
    null,
    null,
    null,
    null,
    null,
  ];
  return [...original, ...added];
});

const baseEndCol = excelCol(headers.length - 1);
const baseEndRow = totalRows + 1;
base.getRange(`A1:${baseEndCol}${baseEndRow}`).values = [headers, ...dataValues];
styleTableHeader(base.getRange(`A1:${baseEndCol}1`), palette.navy);

if (totalRows > 0) {
  const formulaRows = [];
  for (let i = 2; i <= baseEndRow; i++) {
    formulaRows.push([
      `=IF(OR(L${i}="",N${i}=""),"",ROUND((N${i}-L${i})*24,2))`,
      `=IF(W${i}="","",ROUND(W${i}/24,2))`,
      `=IF(W${i}="","Sin cierre",IF(W${i}<=24,"<=24h",IF(W${i}<=48,"24-48h",IF(W${i}<=72,"48-72h",IF(W${i}<=120,"3-5 dias",">5 dias")))))`,
      `=IF(L${i}="","Sin fecha",TEXT(L${i},"yyyy-mm"))`,
      `=IF(ISNUMBER(SEARCH("ANULADO",E${i})),"Anulado","Resuelto")`,
      `=IF(W${i}="","Sin fecha",IF(W${i}<=48,"<=48h",">48h"))`,
    ]);
  }
  base.getRange(`W2:AB${baseEndRow}`).formulas = formulaRows;
}

base.tables.add(`A1:${baseEndCol}${baseEndRow}`, true, "TablaCasosTipificados");
base.freezePanes.freezeRows(1);
base.getRange(`A1:${baseEndCol}${baseEndRow}`).format.wrapText = true;
base.getRange(`J2:J${baseEndRow}`).format.numberFormat = "yyyy-mm-dd hh:mm";
base.getRange(`L2:L${baseEndRow}`).format.numberFormat = "yyyy-mm-dd hh:mm";
base.getRange(`N2:N${baseEndRow}`).format.numberFormat = "yyyy-mm-dd hh:mm";
base.getRange(`W2:X${baseEndRow}`).format.numberFormat = "0.00";
base.getRange(`A1:A${baseEndRow}`).format.columnWidthPx = 92;
base.getRange(`B1:B${baseEndRow}`).format.columnWidthPx = 260;
base.getRange(`E1:E${baseEndRow}`).format.columnWidthPx = 210;
base.getRange(`J1:J${baseEndRow}`).format.columnWidthPx = 132;
base.getRange(`L1:N${baseEndRow}`).format.columnWidthPx = 132;
base.getRange(`O1:R${baseEndRow}`).format.columnWidthPx = 260;
base.getRange(`T1:V${baseEndRow}`).format.columnWidthPx = 230;
base.getRange(`W1:AB${baseEndRow}`).format.columnWidthPx = 128;
base.getRange(`A2:${baseEndCol}${baseEndRow}`).format.font = { color: palette.text, size: 10 };

const hours = rows.map((row) => row._tiempo_resolucion_horas);
const days = rows.map((row) => row._tiempo_resolucion_dias);
const resolved = rows.filter((row) => row._resultado === "Resuelto").length;
const annulled = rows.filter((row) => row._resultado === "Anulado").length;
const in48 = rows.filter((row) => row._cumplimiento_48h === "<=48h").length;
const maxRow = rows.reduce((best, row) => {
  if (!best) return row;
  return (row._tiempo_resolucion_horas || -1) > (best._tiempo_resolucion_horas || -1) ? row : best;
}, null);

dashboard.getRange("A1:H1").merge();
setTitle(dashboard.getRange("A1"), "Tablero de casos - tipificacion y tiempos de resolucion");
dashboard.getRange("A2:H2").merge();
dashboard.getRange("A2").values = [[`Fuente: ${payload.source} | Generado: ${payload.generated_at}`]];
dashboard.getRange("A2").format = { font: { color: palette.muted, size: 9 }, fill: palette.gray };

const kpiHeaders = [["Total casos", "Resueltos", "Anulados", "Prom. dias", "Mediana dias", "% <=48h", "Mayor tiempo", "Caso mayor"]];
const kpiValues = [[
  totalRows,
  resolved,
  annulled,
  Math.round(average(days) * 100) / 100,
  Math.round(median(days) * 100) / 100,
  totalRows ? in48 / totalRows : 0,
  Math.round((maxRow?._tiempo_resolucion_dias || 0) * 100) / 100,
  maxRow?.Número || "",
]];
dashboard.getRange("A4:H5").values = [...kpiHeaders, ...kpiValues];
styleTableHeader(dashboard.getRange("A4:H4"), palette.navy);
dashboard.getRange("A5:H5").format = {
  fill: palette.sky,
  font: { bold: true, color: palette.text, size: 12 },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
dashboard.getRange("F5").format.numberFormat = "0%";
dashboard.getRange("D5:E5").format.numberFormat = "0.00";
dashboard.getRange("G5").format.numberFormat = "0.00";

styleSection(dashboard.getRange("A7"), "Casos por tipificacion");
const categorySummary = groupSummary("_tipificacion");
const categoryRows = categorySummary.map((item) => [item.key, item.count, item.avgDays, item.pct]);
writeSimpleTable(dashboard, "A8", ["Tipificacion", "Casos", "Prom. dias", "% total"], categoryRows);
dashboard.getRange(`C9:C${8 + categoryRows.length}`).format.numberFormat = "0.00";
dashboard.getRange(`D9:D${8 + categoryRows.length}`).format.numberFormat = "0%";

styleSection(dashboard.getRange("A18"), "Casos por producto");
const productRows = groupSummary("Producto").map((item) => [item.key, item.count, item.avgDays, item.pct]);
writeSimpleTable(dashboard, "A19", ["Producto", "Casos", "Prom. dias", "% total"], productRows);
dashboard.getRange(`C20:C${19 + productRows.length}`).format.numberFormat = "0.00";
dashboard.getRange(`D20:D${19 + productRows.length}`).format.numberFormat = "0%";

styleSection(dashboard.getRange("F7"), "Rangos de resolucion");
const rangeRows = countBy("_rango_tiempo").map(([key, count]) => [key, count, totalRows ? count / totalRows : 0]);
writeSimpleTable(dashboard, "F8", ["Rango", "Casos", "% total"], rangeRows);
dashboard.getRange(`H9:H${8 + rangeRows.length}`).format.numberFormat = "0%";

styleSection(dashboard.getRange("F18"), "Resultado de cierre");
const resultRows = countBy("_resultado").map(([key, count]) => [key, count, totalRows ? count / totalRows : 0]);
writeSimpleTable(dashboard, "F19", ["Resultado", "Casos", "% total"], resultRows);
dashboard.getRange(`H20:H${19 + resultRows.length}`).format.numberFormat = "0%";

styleSection(dashboard.getRange("J18"), "Casos por mes");
const monthRows = countBy("_mes_creacion").sort((a, b) => a[0].localeCompare(b[0])).map(([key, count]) => [key, count]);
writeSimpleTable(dashboard, "J19", ["Mes", "Casos"], monthRows);

const chart = dashboard.charts.add("bar", dashboard.getRange(`A8:B${8 + categoryRows.length}`));
chart.title = "Casos por tipificacion";
chart.hasLegend = false;
chart.setPosition("J7", "P17");
chart.xAxis = { axisType: "textAxis" };
chart.yAxis = { numberFormatCode: "0" };

const chart2 = dashboard.charts.add("bar", dashboard.getRange(`F8:G${8 + rangeRows.length}`));
chart2.title = "Rangos de resolucion";
chart2.hasLegend = false;
chart2.setPosition("J24", "P34");
chart2.xAxis = { axisType: "textAxis" };
chart2.yAxis = { numberFormatCode: "0" };

dashboard.freezePanes.freezeRows(5);
dashboard.getRange("A:A").format.columnWidthPx = 220;
dashboard.getRange("B:D").format.columnWidthPx = 100;
dashboard.getRange("F:F").format.columnWidthPx = 150;
dashboard.getRange("G:H").format.columnWidthPx = 90;
dashboard.getRange("J:P").format.columnWidthPx = 105;

rules.getRange("A1:B1").merge();
setTitle(rules.getRange("A1"), "Reglas de tipificacion utilizadas");
rules.getRange("A3:B3").values = [["Tipificacion", "Descripcion"]];
styleTableHeader(rules.getRange("A3:B3"), palette.navy);
rules.getRange(`A4:B${3 + payload.categories.length}`).values = payload.categories.map((item) => [
  item.tipificacion,
  item.descripcion,
]);
rules.tables.add(`A3:B${3 + payload.categories.length}`, true, "TablaReglasTipificacion");
rules.getRange(`A1:B${3 + payload.categories.length}`).format.wrapText = true;
rules.getRange("A:A").format.columnWidthPx = 180;
rules.getRange("B:B").format.columnWidthPx = 680;

const notesStart = 14;
rules.getRange(`A${notesStart}:B${notesStart}`).merge();
styleSection(rules.getRange(`A${notesStart}`), "Notas de uso");
rules.getRange(`A${notesStart + 1}:B${notesStart + 3}`).values = [
  ["Tipificacion sugerida", "Se calculo con reglas de palabras clave y ajustes manuales para los 11 casos del archivo."],
  ["Tiempo de resolucion", "Se recalcula como diferencia calendario entre Cerrado y Creado, porque la columna Duracion del export viene en 0."],
  ["Referencia 48h", "Es una referencia operativa para segmentar casos; no reemplaza el SLA contractual si existe uno diferente."],
];
rules.getRange(`A${notesStart + 1}:B${notesStart + 3}`).format.wrapText = true;
rules.getRange(`A${notesStart + 1}:B${notesStart + 3}`).format = {
  wrapText: true,
  verticalAlignment: "top",
};
rules.getRange(`A${notesStart + 1}:B${notesStart + 3}`).format.rowHeightPx = 44;

const formulaErrors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "final formula error scan",
});
console.log(formulaErrors.ndjson);

await fs.mkdir(outputDir, { recursive: true });
for (const sheetName of ["Tablero", baseSheetName, "Reglas tipificacion"]) {
  const imageBlob = await workbook.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  const imageBytes = new Uint8Array(await imageBlob.arrayBuffer());
  await fs.writeFile(path.join(outputDir, `${sheetName.replaceAll(" ", "_")}.png`), imageBytes);
}

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputXlsx);
console.log(`Saved ${outputXlsx}`);
