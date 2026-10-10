# Clasificación de casos

Las reglas vigentes están en `services/casos_historial.py`. Los dashboards,
filtros, seguimiento de token y reportes reutilizan el mismo análisis local.
La sincronización y los datos originales de ServiceNow no se modifican.

Cada número de caso cuenta una sola vez, conservando la última fila cargada.
La **clasificación principal** describe el motivo del reporte, salvo las
derivaciones sin atención técnica registrada, que se agrupan exclusivamente
en **Enviado a otros canales**. Agenda, PQRS, ventas y comercial son destinos,
no tipologías adicionales. Una instalación derivada a agenda cuenta una vez
en Enviado a otros canales; su solicitud de instalación se conserva en el detalle.

La **tipología** resume esa clasificación en Soporte Uso, Incidentes,
Solicitudes, Enviado a otros canales o Por revisar. Las tipificaciones antiguas
se conservan como datos de origen y no determinan la clasificación vigente.

**Producto**, **problema** y **causa raíz** son datos distintos:

- Token físico requiere una mención explícita positiva de «token físico»
  (incluidas variantes ortográficas del término). USB, ePass, driver, SafeNet
  y una instalación genérica no identifican por sí solos un token físico.
- Los reportes de caídas o indisponibilidad se agrupan por ese problema;
  las solicitudes o dificultades de consulta de consumos tienen su propio grupo.
- La causa raíz requiere una explicación registrada en el campo de causa o
  en la atención. Una instalación, un producto, una redirección o un síntoma
  no confirman una causa técnica. Las causas documentadas se consultan aparte.

Se excluyen acuses automáticos, pies de canales, catálogos genéricos de
productos en respuestas de derivación y mensajes citados. Se conserva la traza
original y el fragmento que respalda cada detección. Si hay trabajo técnico
antes de una redirección, el caso conserva su problema como clasificación.
Sin evidencia de atención no significa que se haya probado falta de atención.

El seguimiento de producto es un subconjunto de los casos, no otra distribución
para sumar al total. Los porcentajes de clasificación usan los casos únicos;
los porcentajes de causas documentadas usan únicamente los casos con explicación.
El resumen usa soporte y casos sin asignación. El detalle permite revisar todos
los responsables con el alcance indicado en pantalla.
