import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const workspace = "C:\\Users\\yerik\\OneDrive\\Desktop\\gestion_problemas_yerika";
const outDir = path.join(workspace, "outputs", "ans_rpost");
const outputPath = path.join(outDir, "ANS_RPOST_analisis_afectacion.xlsx");
const input = await FileBlob.load(outputPath);
const workbook = await SpreadsheetFile.importXlsx(input);

const overview = await workbook.inspect({
  kind: "sheet,table",
  maxChars: 6000,
  tableMaxRows: 3,
  tableMaxCols: 7,
});
console.log(overview.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "final formula error scan",
});
console.log(errors.ndjson);

for (const sheetName of [
  "ANS RPOST",
  "Dashboard",
  "Evaluacion ANS",
  "Disponibilidad Mes",
  "Detalle ANS RPOST",
  "Solucion ANS",
  "Mantenimiento ANS",
  "Acuses ANS",
  "Solicitudes Proveedor",
]) {
  const blob = await workbook.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  const bytes = new Uint8Array(await blob.arrayBuffer());
  await fs.writeFile(path.join(outDir, `${sheetName.replaceAll(" ", "_")}.png`), bytes);
  console.log(`rendered ${sheetName} ${bytes.length}`);
}
