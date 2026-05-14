import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const workspace = process.cwd();
const sourcePath =
  process.argv[2] ??
  process.env.ANS_RPOST_SOURCE ??
  path.join(os.homedir(), "Downloads", "ANS_RPOST_con_Dashboard.xlsx");
const outDir = path.join(workspace, "outputs", "ans_rpost");
const dataPath = path.join(outDir, "ans_rpost_analysis.json");
const outputPath = path.join(outDir, "ANS_RPOST_analisis_afectacion.xlsx");

const data = JSON.parse(await fs.readFile(dataPath, "utf8"));
const input = await FileBlob.load(sourcePath);
const workbook = await SpreadsheetFile.importXlsx(input);

const colors = {
  navy: "#1F4E78",
  teal: "#0F766E",
  blue: "#D9EAF7",
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
  const start = range.split(":")[0];
  sheet.getRange(start).values = [[title]];
  sheet.getRange(start).format = {
    fill: colors.navy,
    font: { bold: true, color: colors.white, size: 16 },
    horizontalAlignment: "left",
    verticalAlignment: "center",
    rowHeightPx: 38,
  };
  if (subtitle) {
    sheet.getRange("A2:H2").merge();
    sheet.getRange("A2").values = [[subtitle]];
    sheet.getRange("A2").format = {
      fill: colors.blue,
      font: { color: colors.darkGray, size: 10 },
      wrapText: true,
      verticalAlignment: "center",
      rowHeightPx: 32,
    };
  }
}

function writeTable(sheet, startRow, startCol, headers, rows, tableName) {
  const matrix = [headers, ...rows.map((row) => headers.map((header) => row[header] ?? ""))];
  const range = sheet.getRangeByIndexes(startRow, startCol, matrix.length, headers.length);
  range.values = matrix;
  sheet.getRangeByIndexes(startRow, startCol, 1, headers.length).format = {
    fill: colors.teal,
    font: { bold: true, color: colors.white },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
    rowHeightPx: 34,
  };
  if (rows.length > 0) {
    sheet.getRangeByIndexes(startRow + 1, startCol, rows.length, headers.length).format = {
      wrapText: true,
      verticalAlignment: "top",
    };
    const table = sheet.tables.add(range.address, true, tableName);
    table.style = "TableStyleMedium2";
    table.showFilterButton = true;
  }
  return range;
}

function setWidths(sheet, widths) {
  widths.forEach((width, idx) => {
    sheet.getRangeByIndexes(0, idx, 1, 1).format.columnWidthPx = width;
  });
}

function applyStatusColors(sheet, rowStart, rowCount, colIndex) {
  const range = sheet.getRangeByIndexes(rowStart, colIndex, rowCount, 1);
  const topLeft = range.address.split(":")[0];
  range.conditionalFormats.addCustom(`=${topLeft}="Afectado"`, { fill: colors.red });
  range.conditionalFormats.addCustom(`=${topLeft}="Cumple"`, { fill: colors.green });
}

function updateDashboard() {
  const dashboard = workbook.worksheets.getItem("Dashboard");
  dashboard.showGridLines = false;
  dashboard.getRange("B1:E1").merge();
  dashboard.getRange("B1").values = [["Resultado calculado con incidentes RPOST"]];
  dashboard.getRange("B1").format = {
    fill: colors.navy,
    font: { bold: true, color: colors.white },
    horizontalAlignment: "center",
  };
  const affectedMonths = data.summary.months_affected_availability;
  const rows = [
    ["Disponibilidad ANS", affectedMonths === "Ninguno" ? "Sin afectacion" : `Afectada: ${affectedMonths}`],
    ["Incidentes RPOST analizados", data.summary.total_incidents],
    ["Ventanas mantenimiento evidenciadas", data.summary.maintenance_mentions],
    ["Casos fuera ANS solucion", data.summary.outside_solution_cases],
    ["Casos condicionales solucion", data.summary.conditional_solution_cases],
    ["Acuses afectados", data.summary.acuses_affected],
    ["Compensacion acuses", data.summary.acuse_compensation],
    ["Umbral disponibilidad", data.summary.availability_sla],
    ["Targets solucion", data.summary.solution_targets],
  ];
  dashboard.getRange("B3:C11").values = rows;
  dashboard.getRange("B3:B11").format = {
    fill: colors.blue,
    font: { bold: true, color: colors.darkGray },
    wrapText: true,
  };
  dashboard.getRange("C3:C11").format = {
    fill: colors.gray,
    font: { color: colors.darkGray },
    wrapText: true,
  };
  dashboard.getRange("B13:E13").merge();
  dashboard.getRange("B13").values = [[
    "Nota: los calculos usan la duracion oficial del ticket. Para compensacion final se recomienda exigir al proveedor hora real de inicio/fin, RCA, confirmacion de ventanas de mantenimiento y trazabilidad individual de acuses.",
  ]];
  dashboard.getRange("B13").format = {
    fill: colors.amber,
    font: { color: colors.darkGray },
    wrapText: true,
    rowHeightPx: 48,
  };
  setWidths(dashboard, [280, 250, 280, 160, 160]);
}

updateDashboard();

const evalSheet = addSheet("Evaluacion ANS");
setTitle(
  evalSheet,
  "A1:D1",
  "Evaluacion de afectacion ANS RPOST",
  `ANS: ${data.summary.ans_source} | Incidentes: ${data.summary.incident_source} | Generado: ${data.summary.generated}`,
);
const kpiRows = [
  ["Indicador", "Resultado", "Lectura", "Accion"],
  ["Disponibilidad", data.summary.months_affected_availability, "Meses con disponibilidad por debajo de 99.9%.", "Solicitar RCA y soporte de mantenimiento."],
  ["Ventanas de mantenimiento", data.summary.maintenance_mentions, "No hay evidencia si el valor es 0.", "No excluir caidas sin soporte formal del proveedor."],
  ["Casos fuera solucion ANS", data.summary.outside_solution_cases, "Calculado con duracion oficial del ticket.", "Validar hora real de solucion con proveedor."],
  ["Casos condicionales", data.summary.conditional_solution_cases, "Caso sujeto a responsabilidad o configuracion del cliente.", "Validar si aplica o no penalizacion."],
  ["Acuses", `${data.summary.acuses_affected} afectados / ${data.summary.acuse_compensation}`, "El rango 101 a 250 acuses no generados oportunamente aplica 5%.", "Incluir CS0196170 y PQRF-2026-1902 en reclamo al proveedor."],
  ["Atencion inicial", "No evaluable", "No hay hora de primera atencion.", "Agregar campo de primera respuesta."],
];
evalSheet.getRange("A4:D10").values = kpiRows;
evalSheet.getRange("A4:D4").format = { fill: colors.teal, font: { bold: true, color: colors.white } };
evalSheet.getRange("A5:D10").format = { wrapText: true, verticalAlignment: "top" };
writeTable(evalSheet, 12, 0, ["Tema", "Conclusion", "Evidencia", "Accion"], data.conclusions, "ConclusionesANS");
setWidths(evalSheet, [190, 420, 420, 420]);
evalSheet.freezePanes.freezeRows(4);

const availSheet = addSheet("Disponibilidad Mes");
setTitle(availSheet, "A1:K1", "Calculo mensual de disponibilidad");
const availabilityHeaders = [
  "Mes",
  "Segundos mes",
  "Minutos caida permitidos 99.9",
  "Indisponibilidad oficial",
  "Indisponibilidad oficial seg",
  "Disponibilidad calculada",
  "ANS disponibilidad >=99.9",
  "Indisponibilidad calendario",
  "Disponibilidad calendario",
  "ANS calendario >=99.9",
  "Exceso sobre permitido",
  "Casos que cuentan disponibilidad",
  "Ventanas mantenimiento evidenciadas",
];
writeTable(availSheet, 2, 0, availabilityHeaders, data.monthly_availability, "DisponibilidadMensual");
setWidths(availSheet, [90, 120, 190, 170, 160, 160, 180, 180, 170, 170, 170, 190, 210]);
applyStatusColors(availSheet, 3, data.monthly_availability.length, 6);
applyStatusColors(availSheet, 3, data.monthly_availability.length, 9);
availSheet.freezePanes.freezeRows(3);

const detailSheet = addSheet("Detalle ANS RPOST");
setTitle(detailSheet, "A1:Y1", "Detalle de incidentes y evaluacion ANS");
const detailHeaders = [
  "Numero",
  "Mes",
  "Creado",
  "Cerrado",
  "Servicio",
  "Categoria",
  "Prioridad ServiceNow",
  "Prioridad ANS",
  "Target solucion ANS",
  "Duracion oficial",
  "Tiempo calendario",
  "Diferencia duracion",
  "Caso proveedor",
  "Tipo evento",
  "Grupo evento",
  "Cuenta disponibilidad ANS",
  "Fuera ANS solucion",
  "Notas solucion ANS",
  "Causa raiz / probable",
  "Evidencia",
  "Ventana mantenimiento",
  "Detalle mantenimiento",
  "Responsabilidad preliminar",
  "Breve descripcion",
];
writeTable(detailSheet, 2, 0, detailHeaders, data.detail, "DetalleANSRPOST");
setWidths(detailSheet, [
  105, 85, 145, 145, 150, 105, 135, 110, 130, 130, 135, 190, 135, 190, 180, 170, 150, 250, 350, 280, 160, 260, 260, 330,
]);
detailSheet.freezePanes.freezeRows(3);
detailSheet.freezePanes.freezeColumns(1);

const solutionSheet = addSheet("Solucion ANS");
setTitle(solutionSheet, "A1:F1", "Cumplimiento de tiempos de solucion");
writeTable(
  solutionSheet,
  2,
  0,
  ["Mes", "Incidentes fuera ANS solucion", "Casos fuera ANS solucion", "Casos condicionales", "Riesgo compensacion incidentes", "Nota"],
  data.solution_by_month,
  "SolucionANS",
);
setWidths(solutionSheet, [90, 200, 360, 220, 210, 540]);
solutionSheet.freezePanes.freezeRows(3);

const maintenanceSheet = addSheet("Mantenimiento ANS");
setTitle(maintenanceSheet, "A1:G1", "Validacion de ventanas de mantenimiento");
writeTable(
  maintenanceSheet,
  2,
  0,
  ["Numero", "Mes", "Caso proveedor", "Ventana mantenimiento", "Conclusion", "Impacto en ANS", "Solicitud al proveedor"],
  data.maintenance,
  "MantenimientoANS",
);
setWidths(maintenanceSheet, [105, 85, 135, 170, 390, 170, 500]);
maintenanceSheet.freezePanes.freezeRows(3);

const acusesSheet = addSheet("Acuses ANS");
setTitle(
  acusesSheet,
  "A1:N1",
  "Evaluacion ANS por acuses no generados oportunamente",
  "Caso CS0196170 / PQRF-2026-1902 - Empresa de Energia de Pereira S.A. E.S.P.",
);
writeTable(
  acusesSheet,
  3,
  0,
  [
    "Caso",
    "PQRS",
    "Cliente",
    "Servicio",
    "Tipo",
    "Fecha reporte",
    "Fecha asignacion",
    "Fecha escalamiento proveedor",
    "Fechas afectadas",
    "Mes ANS",
    "Cantidad acuses afectados",
    "Rango ANS",
    "Compensacion acuses",
    "Aplica compensacion",
    "Base ANS",
    "Observacion",
  ],
  data.acuse_cases,
  "AcusesANS",
);
acusesSheet.getRange("A7:N7").merge();
acusesSheet.getRange("A7").values = [[
  "Conclusion: con 217 acuses no generados oportunamente en mayo de 2026, el evento cae en el rango contractual de 101 a 250 acuses y aplica compensacion del 5% sobre la facturacion total del respectivo mes.",
]];
acusesSheet.getRange("A7").format = {
  fill: colors.amber,
  font: { bold: true, color: colors.darkGray },
  wrapText: true,
  rowHeightPx: 50,
};
setWidths(acusesSheet, [105, 130, 300, 120, 110, 150, 150, 190, 190, 90, 180, 150, 150, 150, 320, 520]);
acusesSheet.freezePanes.freezeRows(4);

const providerSheet = addSheet("Solicitudes Proveedor");
setTitle(providerSheet, "A1:D1", "Solicitudes al proveedor para cierre de ANS");
const providerRows = [
  {
    Solicitud: "RCA por evento",
    Detalle: "Entregar causa raiz, linea de tiempo, impacto, hora real de inicio/fin y accion preventiva por cada ticket proveedor.",
    Casos: "59483, 59743, 59820, 60054, 60183, 60305, 60306, 60836",
    Prioridad: "Alta",
  },
  {
    Solicitud: "Ventanas de mantenimiento",
    Detalle: "Confirmar por escrito si existieron mantenimientos, cambios o ventanas acordadas con al menos 15 dias de anticipacion.",
    Casos: "Todos los incidentes que cuentan para disponibilidad",
    Prioridad: "Alta",
  },
  {
    Solicitud: "Disponibilidad RPOST/WBLM",
    Detalle: "Explicar recurrencia de estados Down/unknown/time out sobre PORTAL.RPOST.COM - WBLM en Agrio.",
    Casos: "Marzo y abril",
    Prioridad: "Alta",
  },
  {
    Solicitud: "Evidencia de solucion",
    Detalle: "Entregar hora real de restauracion del servicio, no solo hora de cierre administrativo del ticket.",
    Casos: "Casos fuera de ANS solucion",
    Prioridad: "Alta",
  },
  {
    Solicitud: "Acuses",
    Detalle: "Responder por los 217 acuses no generados oportunamente en mayo, indicando causa raiz, trazabilidad por mensaje, fecha/hora de generacion o certificacion y plan preventivo.",
    Casos: "CS0196170 / PQRF-2026-1902",
    Prioridad: "Alta",
  },
];
writeTable(providerSheet, 2, 0, ["Solicitud", "Detalle", "Casos", "Prioridad"], providerRows, "SolicitudesProveedor");
setWidths(providerSheet, [230, 600, 320, 110]);
providerSheet.freezePanes.freezeRows(3);

await fs.mkdir(outDir, { recursive: true });
const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "formula error scan",
});
console.log(errors.ndjson);
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(outputPath);
