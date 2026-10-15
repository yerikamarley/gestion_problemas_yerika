# Tiempos y categorías de incidentes

Dashboard Incidentes y KPI Incidentes comparten el cálculo de
`services/incidentes_sla.py` y el panel `components/incidentes_ans.py`.
La fuente de objetivos es la tabla 1 de SER-PR-001 v003.

## Horario confirmado

Lunes a viernes de 08:00 a 17:00, zona America/Bogota. Se cuentan nueve horas
de soporte por día, sin descontar almuerzo. Cuatro días equivalen a 36 horas
de soporte y ocho días a 72. Noches y fines de semana no consumen el plazo.
El mismo horario aplica a atención y resolución de las tres categorías.

La configuración está en `config/incidentes_ans.json`. La lista `festivos`
está vacía porque aún no se indicó un calendario de festivos. El panel lo
expone expresamente. Para excluir fechas se agregan como `AAAA-MM-DD` en
esa lista; no se deducen ni descargan automáticamente.

## Tabla de objetivos

| Categoría | Prioridad | Atención | Resolución |
| --- | --- | --- | --- |
| Incidente | Crítico | 15 minutos | 4 horas |
| Incidente | Alto | 20 minutos | 8 horas |
| Incidente | Moderado | 30 minutos | 4 días |
| Incidente | Bajo | 30 minutos | 8 días |
| Incidente Seguridad | Crítico | 15 minutos | 4 horas |
| Incidente Seguridad | Alto | 20 minutos | 8 horas |
| Incidente Seguridad | Moderado | 30 minutos | 24 horas |
| Incidente Seguridad | Bajo | 30 minutos | 48 horas |
| Consulta / Ayuda | Moderado | 30 minutos | 4 días |

Cada categoría muestra el cargo coordinador definido en la tabla. Una consulta
externa conserva Consulta / Ayuda; no se transforma en un caso de mesa de ayuda
por pertenecer a un cliente externo. Para consultas se aplica siempre el plazo
único de la guía (fila Moderado), sin depender de la prioridad registrada: incluso
si falta o difiere, no genera una validación pendiente. Se conserva el dato de
origen. Las categorías desconocidas y las prioridades no válidas de incidentes
quedan pendientes de clasificación; no se asigna Moderado por defecto a incidentes.

## Fechas y exactitud

- Inicio: notificación/apertura explícita; en su ausencia, creación del registro,
  identificada como tal en el detalle.
- Atención: fecha explícita de primera atención.
- Resolución: fecha explícita de resolución. El cierre administrativo se
  conserva por separado y no se usa para acreditar la solución.
- La última actualización, las notas y los correos no se interpretan por sí solos
  como primera atención. Una duración histórica sin semántica de resolución
  validada tampoco sustituye una fecha de solución.
- Se comparan segundos exactos de soporte, con el máximo incluido según la
  tabla facilitada. No se redondean duraciones antes de decidir el resultado.
  Un segundo adicional de soporte representa incumplimiento. Fuera del horario
  no avanza el reloj, aunque haya transcurrido tiempo calendario.
- Las fechas sin zona se interpretan en Colombia y las fechas UTC se convierten.
  Fechas anteriores al inicio, inválidas o futuras se muestran como inconsistentes.

## Integración con ServiceNow y datos históricos

La importación agrega columnas locales `fecha_notificacion`, `fecha_atencion`
y `fecha_resolucion` mediante la migración habitual de `init_db`.

Los candidatos de origen son:

| Columna local | Campos de API, en orden |
| --- | --- |
| fecha_notificacion | u_fecha_notificacion, opened_at |
| fecha_atencion | u_fecha_primera_atencion, u_first_response_at, first_response_at, u_fecha_atencion |
| fecha_resolucion | resolved_at |

Estos nombres son candidatos: solo se importan si vienen en la respuesta.
Para una instancia con otro campo, configurar el mapa existente
`SERVICENOW_INCIDENT_FIELD_MAP`, por ejemplo:

```json
{"fecha_atencion": "nombre_real_del_campo_de_primera_atencion"}
```

No debe asignarse un campo de actualización o asignación sin verificar su
significado. Si no existe evidencia de primera atención, el resultado será
«Sin fecha de atención», incluso si el incidente está abierto. La ausencia de
esta fecha no demuestra que el equipo no lo haya atendido y no genera por sí
sola un incumplimiento o vencimiento de atención.

Las consultas del Dashboard/KPI recuperan también estas fechas desde los JSON
ya sincronizados en `servicenow_records`, sin nuevas llamadas API. Si esos JSON
no contienen los campos, hará falta una sincronización con acceso a ellos. La
lectura mantiene los controles de autorización de las vistas; no modifica los
incidentes remotos.

## Métricas y gráficas

Los filtros de categoría y prioridad afectan la tabla de objetivos, las tarjetas,
las gráficas y la descarga. La tabla de tiempos aplicables permanece visible.
Por decisión de presentación, el Dashboard y el KPI muestran únicamente el SLA
de resolución: se retiraron atención de las tarjetas, gráficas, tendencia,
detalle descargable, objetivos visibles y diapositiva. Esto no modifica los
requisitos del procedimiento ni elimina los datos o reglas de atención del sistema.

- Cumplimiento de resolución = cumplen / (cumplen + no cumplen).
- Se muestran también evaluados, registros sin datos, pendientes y pendientes
  vencidos. No se convierten los faltantes en cumplidos ni en incumplidos.
- Si no hay evaluables se presenta «No calculable», con el campo faltante o
  el motivo de revisión por caso, nunca un 0 % ficticio. Si toda la etapa carece
  de evidencia de tiempo, se muestra la explicación en vez de una gráfica de
  barras completamente grises.
- Los abiertos vencidos se controlan por separado del cumplimiento de etapas
  ya terminadas. Los resueltos sin fecha de resolución siguen sin ser evaluables.
- El período se filtra por creación: el total de abiertos corresponde a ese
  conjunto, no al pendiente global de todos los períodos.
- La gráfica de cumplimiento agrupa los registros por tipo (Incidente,
  Incidente Seguridad y Consulta / Ayuda), en lugar de prioridad. Solo aparecen
  tipos con registros en el filtro actual; los segmentos muestran resolución y estados
  con colores consistentes. En «SLA mensual», un solo mes se representa con
  barras horizontales separadas por tipo, porcentajes visibles y una lectura
  de cumplidos/evaluados, fuera del plazo y abiertos. Varios meses se comparan
  con barras agrupadas. Los meses se nombran en español y la tabla de cifras
  permanece visible. Los porcentajes sin evidencia no se representan como 0 %.
- El detalle exportable conserva las fechas y prioridades originales, la fuente
  de inicio, objetivos, vencimientos y resultados de resolución.
- La vista Incidentes y su descarga incluyen también las fechas de notificación,
  primera atención y resolución, junto con el motivo de no evaluación de cada etapa.

El KPI conserva el modo diapositiva, con los tiempos de la categoría elegida y
la cobertura de resolución. Sus análisis de temas y repeticiones son
complementarios: se identifican como inferencias y se aclara el alcance temporal
de las repeticiones. La taxonomía detallada de seguridad, las aprobaciones de
cierre y las comunicaciones de escalamiento deben sustentarse con sus propias
evidencias de proceso.

## Verificación

`tests/test_incidentes_sla.py` cubre las nueve combinaciones de la tabla,
el límite exacto y un segundo adicional, noches y fines de semana, festivos
configurados, zona horaria, fechas faltantes/inconsistentes, clasificación,
importación, lectura histórica y filtrado de ambas vistas.
