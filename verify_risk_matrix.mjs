import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = "outputs/matriz_riesgos_ajustada/Matriz_Riesgos_Materializados_2026.xlsx";
const previewDir = "outputs/matriz_riesgos_ajustada/previews";
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));
await fs.mkdir(previewDir, { recursive: true });

const sheets = ["Informe Ejecutivo", "Tendencia Servicios", "Resumen", "Patrones Operativos", "Riesgos", "Conciliación", "Exclusiones", "Metodología"];
for (const sheetName of sheets) {
  const blob = await workbook.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  const safeName = sheetName.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replaceAll(" ", "_");
  await fs.writeFile(`${previewDir}/${safeName}.png`, new Uint8Array(await blob.arrayBuffer()));
}

const overview = await workbook.inspect({ kind: "sheet", include: "id,name", maxChars: 4000 });
const executive = await workbook.inspect({ kind: "table", range: "Informe Ejecutivo!A1:L24", tableMaxRows: 24, tableMaxCols: 12, maxChars: 16000 });
const traceability = await workbook.inspect({ kind: "match", searchTerm: "INC-DEMO", options: { useRegex: false, maxResults: 100 }, maxChars: 4000 });
const sensitive = await workbook.inspect({ kind: "match", searchTerm: "Error 504|Falla OTP|Caída OCSP|breve_descripcion|descripcion|notes", options: { useRegex: true, maxResults: 100 }, maxChars: 4000 });
const errors = await workbook.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options: { useRegex: true, maxResults: 100 }, maxChars: 4000 });
console.log(overview.ndjson);
console.log(executive.ndjson);
console.log("TRACEABILITY_SCAN");
console.log(traceability.ndjson);
console.log("SENSITIVE_SCAN");
console.log(sensitive.ndjson);
console.log("ERROR_SCAN");
console.log(errors.ndjson);
