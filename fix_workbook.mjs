import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = "C:/Users/yerik/Downloads/incident (9).xlsx";
const outputDir = "C:/Users/yerik/OneDrive/Desktop/certicamrara/gestion_problemas_yerika/outputs/incidentes_proveedores";
const outputPath = `${outputDir}/incident_9_corregido.xlsx`;

const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));
const sheet = workbook.worksheets.getItem("Page 1");
const providerSource = sheet.getRange("AE2:AE218").values;
const providerTarget = sheet.getRange("AB2:AB218").values;
const escalated = sheet.getRange("AD2:AD218").values;

let corrected = 0;
for (let i = 0; i < providerTarget.length; i++) {
  const targetBlank = providerTarget[i][0] === null || providerTarget[i][0] === "";
  const source = providerSource[i][0];
  if (targetBlank && source !== null && source !== "" && escalated[i][0] === true) {
    providerTarget[i][0] = source;
    corrected++;
  }
}
sheet.getRange("AB2:AB218").values = providerTarget;

await fs.mkdir(outputDir, { recursive: true });
const preview = await workbook.render({ sheetName: "Page 1", range: "AB34:AE36", scale: 1, format: "png" });
await fs.writeFile(`${outputDir}/verificacion_proveedores.png`, new Uint8Array(await preview.arrayBuffer()));

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  maxChars: 4000,
});
const check = await workbook.inspect({ kind: "table", range: "Page 1!AB1:AE218", tableMaxRows: 218, tableMaxCols: 4, maxChars: 16000 });

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(JSON.stringify({ outputPath, corrected }));
console.log(errors.ndjson);
console.log(check.ndjson);
