import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath =
  "C:/Users/yerik/Downloads/CNC_MST-PR-002 Gestión de incidentes_v6.xlsx";
const outputDir = path.resolve("outputs/analisis_incidentes");

await fs.mkdir(outputDir, { recursive: true });

const input = await FileBlob.load(inputPath);
const workbook = await SpreadsheetFile.importXlsx(input);

async function saveInspect(name, options) {
  const result = await workbook.inspect(options);
  await fs.writeFile(path.join(outputDir, `${name}.ndjson`), result.ndjson, "utf8");
  return result.ndjson;
}

const overview = await saveInspect("01_overview", {
  kind: "workbook,sheet,table,definedName,drawing",
  maxChars: 20000,
  tableMaxRows: 8,
  tableMaxCols: 10,
  tableMaxCellChars: 120,
});

const sheets = await workbook.inspect({
  kind: "sheet",
  include: "id,name,visibility",
  maxChars: 10000,
});
await fs.writeFile(path.join(outputDir, "02_sheets.ndjson"), sheets.ndjson, "utf8");

const sheetRecords = sheets.ndjson
  .split(/\r?\n/)
  .filter(Boolean)
  .map((line) => JSON.parse(line));

for (const sheet of sheetRecords) {
  const name = sheet.name;
  const safeName = name.replace(/[^\p{L}\p{N}_-]+/gu, "_").slice(0, 60);
  await saveInspect(`sheet_${safeName}_region`, {
    kind: "region,table,formula,drawing",
    sheetId: name,
    maxChars: 18000,
    tableMaxRows: 20,
    tableMaxCols: 12,
    tableMaxCellChars: 180,
    options: { maxResults: 120 },
  });
}

console.log(overview);
