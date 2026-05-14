import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const workspace = process.cwd();
const outDir = path.join(workspace, "outputs", "rpost_incidentes");
const outputPath = path.join(outDir, "analisis_incidentes_rpost.xlsx");
const input = await FileBlob.load(outputPath);
const workbook = await SpreadsheetFile.importXlsx(input);

const overview = await workbook.inspect({
  kind: "sheet,table",
  maxChars: 5000,
  tableMaxRows: 3,
  tableMaxCols: 6,
});
console.log(overview.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 50 },
  summary: "final formula error scan",
});
console.log(errors.ndjson);

for (const sheetName of ["Resumen", "Detalle RPOST", "Obs Nosotros", "Obs Proveedor", "Ventanas Mantto", "Metricas"]) {
  const blob = await workbook.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  const bytes = new Uint8Array(await blob.arrayBuffer());
  await fs.writeFile(path.join(outDir, `${sheetName.replaceAll(" ", "_")}.png`), bytes);
  console.log(`rendered ${sheetName} ${bytes.length}`);
}
