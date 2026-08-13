import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = "C:/Users/yerik/Downloads/incident (9).xlsx";
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));

for (const request of [
  { kind: "sheet", include: "id,name", maxChars: 8000 },
  { kind: "table", maxChars: 20000, tableMaxRows: 8, tableMaxCols: 20, tableMaxCellChars: 120 },
  { kind: "definedName", maxChars: 8000 },
  { kind: "formula", maxChars: 30000, options: { maxResults: 400 } },
]) {
  const result = await workbook.inspect(request);
  process.stdout.write(`\n=== ${request.kind} ===\n${result.ndjson}\n`);
}

const sheet = workbook.worksheets.getItem("Page 1");
const values = sheet.getRange("A1:AE218").values;
console.log("\n=== headers ===");
console.log(values[0].map((v, i) => `${i + 1}:${v}`).join(" | "));
console.log("\n=== filas con datos de proveedor (AB:AE) ===");
for (let r = 1; r < values.length; r++) {
  if (values[r].slice(27, 31).some(v => v !== null && v !== "")) {
    console.log(JSON.stringify({row:r+1, numero:values[r][0], estadoF:values[r][5], estadoP:values[r][15], resumen:values[r][2], AB:values[r][27], AC:values[r][28], AD:values[r][29], AE:values[r][30]}));
  }
}

for (const request of [
  { kind: "region", sheetId: "Page 1", range: "A1:AE12", maxChars: 30000 },
  { kind: "match", searchTerm: "proveedor|solucion|resuelt", options: { useRegex: true, maxResults: 300 }, maxChars: 30000 },
]) {
  const result = await workbook.inspect(request);
  process.stdout.write(`\n=== ${request.kind} detalle ===\n${result.ndjson}\n`);
}
