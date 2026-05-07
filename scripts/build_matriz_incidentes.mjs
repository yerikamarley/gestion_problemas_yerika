import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = path.resolve("outputs", "matriz_incidentes");
const outputPath = path.join(outputDir, "Matriz_operativa_gestion_incidentes.xlsx");

await fs.mkdir(outputDir, { recursive: true });

const wb = Workbook.create();

const palette = {
  navy: "#1F4E78",
  blue: "#5B9BD5",
  teal: "#0F766E",
  green: "#70AD47",
  amber: "#FFC000",
  orange: "#ED7D31",
  red: "#C00000",
  gray: "#D9E2F3",
  lightBlue: "#DDEBF7",
  lightGreen: "#E2F0D9",
  lightAmber: "#FFF2CC",
  lightRed: "#FCE4D6",
  white: "#FFFFFF",
  text: "#1F2937",
};

function setupSheet(sheet, title, subtitle, lastCol = "J") {
  sheet.showGridLines = false;
  sheet.getRange(`A1:${lastCol}1`).merge();
  sheet.getRange("A1").values = [[title]];
  sheet.getRange("A1").format = {
    fill: palette.navy,
    font: { color: palette.white, bold: true, size: 16 },
    horizontalAlignment: "left",
    verticalAlignment: "center",
    rowHeightPx: 34,
  };
  sheet.getRange(`A2:${lastCol}2`).merge();
  sheet.getRange("A2").values = [[subtitle]];
  sheet.getRange("A2").format = {
    fill: palette.lightBlue,
    font: { color: palette.text, italic: true },
    wrapText: true,
    verticalAlignment: "center",
    rowHeightPx: 42,
  };
}

function writeTable(sheet, startCell, headers, rows, tableName, widths = []) {
  const start = sheet.getRange(startCell);
  const rowCount = rows.length + 1;
  const colCount = headers.length;
  const range = start.resize(rowCount, colCount);
  range.values = [headers, ...rows];
  const headerRange = start.resize(1, colCount);
  headerRange.format = {
    fill: palette.teal,
    font: { color: palette.white, bold: true },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
    rowHeightPx: 34,
  };
  start.offset(1, 0).resize(rows.length, colCount).format = {
    wrapText: true,
    verticalAlignment: "top",
  };
  const table = sheet.tables.add(range.address, true, tableName);
  table.style = "TableStyleMedium2";
  widths.forEach((width, idx) => {
    if (width) start.offset(0, idx).resize(rowCount, 1).format.columnWidthPx = width;
  });
  return table;
}

function addValidation(sheet, range, values) {
  sheet.getRange(range).dataValidation = {
    rule: { type: "list", values },
  };
}

const portada = wb.worksheets.add("Resumen");
setupSheet(
  portada,
  "Matriz operativa para gestión de incidentes",
  "Versión base construida a partir del procedimiento MST-PR-002 y los hallazgos de la encuesta interna. Completar responsables nominales, contactos y horarios antes de publicar.",
  "H",
);
writeTable(
  portada,
  "A4",
  ["Elemento", "Definición operativa"],
  [
    ["Objetivo", "Reducir mala clasificación, tickets sin información, demoras de asignación, escalamiento confuso y tiempos muertos."],
    ["Entradas permitidas", "NOC: Alerta, Consulta o Incidente. Cliente interno/externo: Consulta o Incidente."],
    ["Regla de alerta", "Solo el NOC puede registrar alertas. Una alerta se convierte en incidente si se confirma afectación, degradación, interrupción, riesgo o indicador de compromiso."],
    ["Responsable inicial", "Todo incidente confirmado se asigna a Soporte N2 para gestión, coordinación, seguimiento y documentación."],
    ["Escalamiento clave", "Infraestructura se activa inmediatamente para incidentes críticos/altos o caída, degradación, indisponibilidad, plataforma, servidores, red o comunicaciones."],
    ["No incidente", "Si no corresponde a incidente, Mesa de ayuda crea el caso con la tipología correcta y se conserva trazabilidad."],
    ["Seguimiento", "NOC recibe notificación automática para visibilidad, monitoreo de tiempos y apoyo en escalamiento."],
  ],
  "ResumenMatriz",
  [230, 760],
);

const origen = wb.worksheets.add("01_Origen_Clasificacion");
setupSheet(origen, "Origen y clasificación preliminar", "Define qué puede registrar cada fuente y cómo se decide si continúa como incidente.", "H");
writeTable(
  origen,
  "A4",
  ["Origen", "Tipos permitidos", "Regla", "Responsable validación", "Salida si confirma incidente", "Salida si NO confirma incidente"],
  [
    ["NOC", "Alerta, Consulta, Incidente", "Solo NOC puede registrar alertas. Validar monitoreo, SOC/NOC, disponibilidad o indicador de compromiso.", "NOC", "Incidente confirmado -> Soporte N2 + notificación NOC", "Mesa de ayuda crea caso/consulta"],
    ["Cliente interno", "Consulta, Incidente", "No puede registrar alerta. Si reporta afectación real o potencial, validar como incidente.", "Soporte N2 / responsable inicial", "Incidente confirmado -> Soporte N2", "Mesa de ayuda crea caso"],
    ["Cliente externo", "Consulta, Incidente", "No puede registrar alerta. Validar impacto sobre servicio, cliente, usuarios y evidencia.", "Soporte N2 / responsable inicial", "Incidente confirmado -> Soporte N2", "Mesa de ayuda crea caso"],
  ],
  "OrigenClasificacion",
  [150, 190, 340, 180, 240, 220],
);

writeTable(
  origen,
  "A10",
  ["Clasificación preliminar", "Criterio", "Acción"],
  [
    ["Alerta", "Señal de monitoreo generada por NOC/SOC/herramientas. No necesariamente es incidente.", "NOC valida. Si hay afectación, degradación, interrupción, riesgo o indicador de compromiso, convertir en incidente."],
    ["Consulta", "Duda, validación o revisión sin afectación confirmada.", "Validar. Si se confirma afectación real o potencial, convertir en incidente; si no, Mesa de ayuda crea caso."],
    ["Incidente operativo", "Interrupción, caída, degradación o reducción de calidad de un servicio.", "Asignar a Soporte N2 y priorizar."],
    ["Incidente de seguridad", "Afecta o puede afectar confidencialidad, integridad o disponibilidad.", "Asignar a Soporte N2, notificar Seguridad/Riesgos y priorizar."],
  ],
  "TiposEntrada",
  [190, 500, 520],
);

const campos = wb.worksheets.add("02_Campos_Minimos");
setupSheet(campos, "Campos mínimos obligatorios", "Ninguna entrada debe avanzar a gestión resolutiva sin datos suficientes para clasificar, asignar y diagnosticar.", "I");
writeTable(
  campos,
  "A4",
  ["Campo", "Obligatorio", "Aplica a", "Propósito", "Ejemplo / guía", "Si falta"],
  [
    ["Servicio afectado", "Sí", "Consulta / Incidente / Alerta NOC", "Permite asignar a área correcta.", "Portal, PKI, red, servidor, base de datos, aplicación.", "Pendiente información"],
    ["Cliente o área afectada", "Sí", "Todos", "Determina impacto y comunicación.", "Cliente externo, Operaciones, Mesa de ayuda.", "Pendiente información"],
    ["Descripción clara", "Sí", "Todos", "Evita reprocesos y mala priorización.", "Qué ocurre, desde cuándo, en qué pantalla/proceso.", "Pendiente información"],
    ["Fecha y hora de detección", "Sí", "Todos", "Base para MTTA, SLA y cronología.", "2026-05-06 14:30.", "Pendiente información"],
    ["Impacto observado", "Sí", "Incidentes", "Determina prioridad.", "Caída total, degradación, usuarios afectados.", "Pendiente información"],
    ["Usuarios afectados", "Sí", "Incidentes", "Mide alcance.", "1 usuario, varios usuarios, todos los clientes.", "Pendiente información"],
    ["Evidencia", "Sí", "Todos si aplica", "Soporta diagnóstico.", "Pantallazo, log, alerta, correo.", "Solicitar evidencia"],
    ["Contacto del reportante", "Sí", "Todos", "Permite validación y cierre.", "Nombre, correo, teléfono.", "Pendiente información"],
    ["Ambiente afectado", "Sí", "Tecnología", "Evita confundir producción/pruebas.", "Producción, pruebas, desarrollo.", "Pendiente información"],
    ["IP / servidor / aplicación", "Cuando aplique", "Infraestructura / seguridad", "Reduce tiempo de diagnóstico.", "srv-app01, 10.x.x.x, URL.", "Solicitar dato técnico"],
  ],
  "CamposMinimos",
  [210, 110, 190, 300, 360, 180],
);

const prioridad = wb.worksheets.add("03_Priorizacion");
setupSheet(prioridad, "Priorización de incidentes", "La prioridad se asigna por impacto y urgencia. Crítico/Alto activa escalamiento inmediato a Infraestructura si hay caída o degradación tecnológica.", "H");
writeTable(
  prioridad,
  "A4",
  ["Prioridad", "Impacto", "Urgencia", "Criterios", "Tiempo atención", "Tiempo resolución", "Escalamiento obligatorio"],
  [
    ["Crítico", "Alto / servicio crítico", "Inmediata", "Interrupción total, múltiples usuarios/clientes, datos sensibles, indicador de compromiso confirmado.", "15 min", "4 horas", "Soporte N2 + Infraestructura inmediata + NOC + Seguridad/Riesgos si aplica"],
    ["Alto", "Alto o medio-alto", "Media-alta", "Degradación severa, varios usuarios afectados, impacto operativo importante, accesos no autorizados sin compromiso total.", "20 min", "8 horas", "Soporte N2 + Infraestructura inmediata si hay caída/degradación + NOC"],
    ["Moderado", "Medio", "Media", "Afectación parcial, grupo reducido, incidente de seguridad sin evidencia clara de compromiso.", "30 min", "24 horas seguridad / 4 días operativo", "Soporte N2; escalar según diagnóstico"],
    ["Bajo", "Bajo", "Baja", "Impacto mínimo, falla menor, evento informativo sin afectación relevante.", "30 min", "48 horas seguridad / 8 días operativo", "Soporte N2; escalar si se identifica afectación mayor"],
  ],
  "Priorizacion",
  [110, 170, 150, 430, 130, 160, 360],
);

const escalamiento = wb.worksheets.add("04_Matriz_Escalamiento");
setupSheet(escalamiento, "Matriz de asignación y escalamiento", "Completar responsables, suplentes y contactos reales. Esta hoja resuelve el problema de no saber a qué área/persona escalar.", "N");
writeTable(
  escalamiento,
  "A4",
  [
    "Servicio / componente",
    "Tipo de incidente",
    "Prioridad",
    "Responsable inicial",
    "Grupo resolutor prioritario",
    "Cuándo escalar",
    "Responsable principal",
    "Suplente",
    "Horario",
    "Contacto hábil",
    "Contacto no hábil",
    "Proveedor asociado",
    "Tiempo antes de reescalar",
    "Canal",
  ],
  [
    ["Plataforma / servicio crítico", "Caída total", "Crítico", "Soporte N2", "Infraestructura", "Inmediato", "Por definir", "Por definir", "7x24 / según matriz", "Por definir", "Por definir", "Si aplica", "15 min sin respuesta", "Correo + llamada + herramienta"],
    ["Plataforma / servicio crítico", "Degradación severa", "Alto", "Soporte N2", "Infraestructura", "Inmediato", "Por definir", "Por definir", "7x24 / según matriz", "Por definir", "Por definir", "Si aplica", "20 min sin respuesta", "Correo + llamada + herramienta"],
    ["Servidores / infraestructura", "Indisponibilidad / recurso saturado", "Crítico/Alto", "Soporte N2", "Infraestructura", "Inmediato", "Por definir", "Por definir", "7x24 / según matriz", "Por definir", "Por definir", "Si aplica", "15-20 min", "Correo + llamada + herramienta"],
    ["Red / comunicaciones", "Caída o latencia severa", "Crítico/Alto", "Soporte N2", "Infraestructura / Redes", "Inmediato", "Por definir", "Por definir", "7x24 / según matriz", "Por definir", "Por definir", "Si aplica", "15-20 min", "Correo + llamada + herramienta"],
    ["Aplicación / servicio", "Error funcional con afectación parcial", "Moderado", "Soporte N2", "Grupo según servicio", "Según diagnóstico", "Por definir", "Por definir", "Horario hábil", "Por definir", "Por definir", "Si aplica", "30 min sin avance", "Herramienta + correo"],
    ["Seguridad de información", "Malware, acceso no autorizado, fuga o indicador de compromiso", "Según impacto", "Soporte N2", "Seguridad de la información / Riesgos", "Inmediato si hay CIA, datos personales o indicador de compromiso", "Por definir", "Por definir", "7x24 si crítico/alto", "Por definir", "Por definir", "Si aplica", "15 min si crítico", "Correo + llamada + herramienta"],
    ["No corresponde a incidente", "Consulta / solicitud / cambio / problema / bug", "No aplica", "Responsable inicial", "Mesa de ayuda", "Al confirmar que no es incidente", "Por definir", "Por definir", "Horario hábil", "Por definir", "No aplica", "No aplica", "No aplica", "Herramienta"],
  ],
  "MatrizEscalamiento",
  [210, 180, 120, 160, 220, 260, 180, 160, 150, 160, 160, 150, 160, 190],
);
addValidation(escalamiento, "C5:C200", ["Crítico", "Alto", "Moderado", "Bajo", "Crítico/Alto", "Según impacto", "No aplica"]);
addValidation(escalamiento, "D5:D200", ["Soporte N2", "NOC", "Responsable inicial", "Mesa de ayuda"]);
addValidation(escalamiento, "E5:E200", ["Infraestructura", "Infraestructura / Redes", "Seguridad de la información / Riesgos", "Grupo según servicio", "Mesa de ayuda"]);

const horario = wb.worksheets.add("05_Horario_No_Habil");
setupSheet(horario, "Disponibilidad y horario no hábil", "Define cómo activar atención cuando NOC identifica un incidente fuera de horario laboral.", "J");
writeTable(
  horario,
  "A4",
  ["Condición", "Quién activa", "A quién contactar", "Canal mínimo", "Tiempo máximo", "Si no responde", "Evidencia requerida"],
  [
    ["Incidente crítico fuera de horario", "NOC", "Soporte N2 + Infraestructura", "Llamada + correo + herramienta", "15 min", "Reescalar a suplente y nivel gerencial", "Hora de contacto, respuesta, acciones"],
    ["Incidente alto fuera de horario", "NOC", "Soporte N2 + Infraestructura si hay caída/degradación", "Llamada + correo + herramienta", "20 min", "Reescalar a suplente", "Hora de contacto, respuesta, acciones"],
    ["Incidente moderado fuera de horario", "NOC / Soporte N2", "Responsable según matriz", "Herramienta + correo", "30 min", "Revisar si amerita disponibilidad", "Registro en ticket"],
    ["Incidente de seguridad crítico/alto", "NOC / Soporte N2", "Seguridad de la información / Riesgos", "Llamada + correo + herramienta", "15-20 min", "Escalar a Gerencia / Comité si aplica", "Cronología y evidencia preservada"],
  ],
  "HorarioNoHabil",
  [240, 150, 320, 220, 140, 300, 300],
);

const terceros = wb.worksheets.add("06_Terceros");
setupSheet(terceros, "Gestión de dependencia de terceros", "Usar cuando la recuperación dependa de proveedor, fabricante, datacenter u otra parte externa.", "K");
writeTable(
  terceros,
  "A4",
  ["Incidente", "Proveedor / tercero", "Servicio afectado", "Hora solicitud", "SLA externo", "Responsable interno", "Estado", "Última actualización", "Próximo seguimiento", "Impacto SLA interno", "Evidencia"],
  [
    ["INC-0000", "Por definir", "Por definir", "", "Por definir", "Soporte N2 / responsable servicio", "Abierto", "", "", "Por evaluar", "Correo, ticket proveedor, llamada"],
  ],
  "Terceros",
  [120, 200, 200, 150, 150, 230, 130, 180, 180, 180, 260],
);
addValidation(terceros, "G5:G200", ["Abierto", "En seguimiento", "Resuelto", "Bloqueado"]);

const cierre = wb.worksheets.add("07_Checklist_Cierre");
setupSheet(cierre, "Checklist de cierre controlado", "El incidente no debe cerrarse si falta trazabilidad mínima.", "H");
writeTable(
  cierre,
  "A4",
  ["Validación", "Obligatorio", "Responsable", "Criterio de aceptación", "Estado"],
  [
    ["Clasificación correcta", "Sí", "Soporte N2", "Incidente operativo o seguridad definido.", "Pendiente"],
    ["Prioridad registrada", "Sí", "Soporte N2", "Crítico, Alto, Moderado o Bajo con justificación.", "Pendiente"],
    ["Responsable asignado", "Sí", "Soporte N2", "Grupo resolutor y persona/equipo asignado.", "Pendiente"],
    ["Diagnóstico documentado", "Sí", "Grupo resolutor", "Causa probable, alcance, evidencias y plan.", "Pendiente"],
    ["Acciones realizadas", "Sí", "Grupo resolutor", "Contención, solución, recuperación o mitigación registradas.", "Pendiente"],
    ["Validación de solución", "Sí", "Soporte N2 / NOC", "Cliente interno, externo o monitoreo confirma restablecimiento.", "Pendiente"],
    ["Dependencias documentadas", "Si aplica", "Soporte N2", "Proveedor, hora, SLA, evidencias y seguimiento.", "Pendiente"],
    ["Lecciones aprendidas", "Si aplica", "Soporte N2 / Seguridad", "Crítico, alto, seguridad o repetitivo con plan de acción.", "Pendiente"],
    ["Informe post-incidente", "Si aplica", "Soporte N2 / Seguridad", "Causa raíz, cronología, impacto, costo, acciones y plan.", "Pendiente"],
  ],
  "ChecklistCierre",
  [260, 120, 180, 520, 130],
);
addValidation(cierre, "E5:E200", ["Pendiente", "Completo", "No aplica"]);

const kpis = wb.worksheets.add("08_KPIs");
setupSheet(kpis, "KPIs y control de mejora continua", "Indicadores recomendados para evidenciar si el flujo reduce conflictos operativos.", "H");
writeTable(
  kpis,
  "A4",
  ["KPI", "Qué mide", "Fórmula / cálculo sugerido", "Frecuencia", "Meta sugerida", "Dueño"],
  [
    ["MTTA", "Tiempo medio hasta atención/asignación efectiva.", "Promedio(fecha primera atención - fecha creación).", "Mensual", "Según SLA por prioridad", "NOC / Soporte N2"],
    ["MTTR", "Tiempo medio de resolución.", "Promedio(fecha solución - fecha creación).", "Mensual", "Según SLA por prioridad", "Soporte N2"],
    ["Cumplimiento SLA", "Porcentaje de incidentes dentro del tiempo objetivo.", "Incidentes dentro de SLA / total incidentes.", "Mensual", ">= 85%", "NOC / Gestión"],
    ["Reasignaciones", "Mala asignación o falta de claridad.", "Cantidad de reasignaciones por incidente y por área.", "Mensual", "Tendencia decreciente", "Soporte N2"],
    ["Pendiente información", "Calidad del registro inicial.", "Entradas devueltas por falta de información / total entradas.", "Mensual", "Tendencia decreciente", "Mesa de ayuda / Soporte N2"],
    ["Incidentes fuera de horario", "Carga operativa no hábil.", "Cantidad por prioridad y área.", "Mensual", "Visibilidad completa", "NOC"],
    ["Dependencias de terceros", "Bloqueos por proveedor.", "Cantidad y tiempo acumulado en dependencia externa.", "Mensual", "Controlado por proveedor", "Soporte N2 / responsable servicio"],
    ["Incidentes repetitivos", "Necesidad de gestión de problemas.", "Incidentes por servicio/causa repetidos.", "Mensual", "Activar gestión de problemas", "Gestión / Soporte N2"],
  ],
  "KPIs",
  [180, 300, 360, 130, 170, 180],
);

for (const sheet of wb.worksheets.items) {
  sheet.freezePanes.freezeRows(4);
}

const errorScan = await wb.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
});
if (errorScan.ndjson.trim()) {
  console.warn(errorScan.ndjson);
}

const xlsx = await SpreadsheetFile.exportXlsx(wb);
await xlsx.save(outputPath);

console.log(outputPath);
