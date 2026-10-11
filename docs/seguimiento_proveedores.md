# Seguimiento de proveedores

Una vista reúne casos e incidentes y permite filtrar mes, proveedor, cliente,
relación y estado. Sustituye los dos menús de RPost y Autentic. Conserva los
roles que ya tenían acceso; las selecciones antiguas se redirigen al menú común.

El catálogo `config/proveedores_ans.json` procede del Excel
`ANS_Consolidado_Proveedores.xlsx`, hoja `Consolidado ANS Proveedores`:
información general en filas 6–9, matriz de prioridades en filas 14–28 y
deducciones de referencia en filas 33–49. El original no se modifica ni se
necesita subir al servidor. Se utiliza la matriz detallada: AutenTIC P1 tiene
una hora de atención y diagnóstico N/A, aunque el resumen general dice
diagnóstico de una hora. Software Colombia P4 conserva solución de diez horas
tal como aparece en la matriz.

## Detección y clientes

Se buscan nombres y alias del proveedor en notas de trabajo, historial,
descripción, observaciones, resolución y campos de proveedor de ambos tipos.
«Certicámara», «token» y «firma digital» por sí solos no identifican proveedor.
Los pies de correo y mensajes anteriores citados se excluyen.

- **Escalamiento confirmado:** indicador marcado con proveedor identificable,
  nota interna que documenta el escalamiento, o fecha explícita de escalamiento.
- **Mención del proveedor:** aparece el nombre pero no hay constancia de
  escalamiento. No participa en el porcentaje de ANS.
- **Proveedor sin identificar:** está marcado el escalamiento y falta un nombre
  que permita asignarlo. No se descarta el registro.

Un proveedor adicional escrito en el campo se conserva sin inventarle un ANS.
Cada asociación conserva la fuente y evidencia. El cliente se toma de Cuenta
en casos y Empresa en incidentes; si falta, se muestra «Cliente sin identificar».
Una mención no demuestra responsabilidad del proveedor ni afectación confirmada.
Las tarjetas cuentan registros únicos por tipo y número, aunque un registro
esté relacionado con varios proveedores. Cada proveedor cuenta su asociación.

## Notas que permiten medir los tiempos

Registrar fechas reales en hora de Colombia, en líneas separadas:

```text
Proveedor: RPost
Escalamiento al proveedor: 09/10/2026 10:00
Primera respuesta del proveedor: 09/10/2026 10:20
Diagnóstico del proveedor: 09/10/2026 10:30
Mitigación del proveedor: 09/10/2026 10:45
Solución del proveedor: 09/10/2026 11:00
```

También se aceptan fechas ISO con hora, segundos y zona explícita. Si hay varios
proveedores, escribir un bloque para cada uno con su nombre. Si se vuelve a
escalar al mismo proveedor, registrar un nuevo escalamiento: se conserva otro
ciclo con sus propias fechas. No copiar como eventos actuales las fechas de un
ciclo previo. La creación del caso, su cierre y el encabezado de un correo no
sustituyen fechas del proveedor. Sin prioridad P1–P4 no se inventa un objetivo.

## Cálculo y corte mensual

Según la confirmación del responsable, atención, diagnóstico, mitigación y
solución corren desde el escalamiento, dentro del horario contractual:

- Software Colombia: 24/7.
- Grafólogos Forenses: P1/P2 24/7; P3/P4 08:00–22:00 todos los días.
- RPost/eSoluciona: lunes–viernes 07:00–18:00; sábado 08:00–13:00;
  domingo sin atención. El Excel no especifica una excepción por festivos para
  este proveedor; se aplica la jornada del día de semana.
- AutenTIC: P1 continuo; otras prioridades lunes–sábado 06:00–22:00 y
  domingos/festivos 08:00–19:00.

Para solución Grafólogos P3/P4, tres días hábiles se interpretan como tres
fechas siguientes de lunes–viernes, sin festivos, conservando la hora del
escalamiento. No se convierten en 72 horas. Los festivos colombianos fijos,
Semana Santa y traslados al lunes se calculan según la
[Ley 51 de 1983](https://www.funcionpublica.gov.co/eva/gestornormativo/norma.php?i=4954).

El mes incluye escalados del mes y pendientes de meses anteriores que seguían
sin solución al inicio. Un escalado sin fecha se conserva pendiente de fecha
hasta su cierre, desde su creación. Las menciones se ubican por creación del
caso/incidente. El corte es el final del mes o la hora actual si el mes no ha
terminado: una solución posterior al corte no cambia el estado en ese corte.

**Cumplimiento = ciclos que cumplen / ciclos evaluados con fechas completas.**
Se informa el denominador, los que incumplen, los abiertos en plazo/vencidos y
los registros sin datos. No se muestra 100 % cuando no hay evaluados. Una fecha
inválida o una respuesta anterior al escalamiento se presenta como inconsistente.

El seguimiento periódico del Excel es un intervalo entre actualizaciones,
distinto de atención y solución. Se presenta como objetivo contractual, sin
evaluarlo a partir de solo esas dos fechas. Disponibilidad, transacciones,
FAR/FRR y deducciones requieren sus datos específicos; se presentan sus objetivos
y referencias, sin inferir cumplimiento o cobros a partir de conteos de casos.

## Datos e informes

Dos consultas a la base compartida recuperan candidatos de casos e incidentes
por nombres, campo de proveedor o indicador de escalamiento, sin paginación
limitante y sin filtrar solo por creación del mes. La detección queda en caché
cinco minutos por actor y vista; la invalidación de casos la limpia. Cambiar
filtros no consulta ServiceNow ni modifica datos o credenciales.

La pantalla permite descargar tres CSV que se abren en Excel: resumen mensual
de ANS con período, corte y cobertura; clientes por proveedor; y detalle con
fechas y evidencia. Los archivos reflejan los filtros seleccionados. Los campos
de texto que pueden interpretarse como fórmula se escapan al exportar.

El corte sobre datos almacenados no reconstruye ediciones históricas eliminadas
de las notas ni incorpora información que aún no se haya sincronizado.
