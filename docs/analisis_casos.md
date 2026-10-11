# Clasificación de casos

Las reglas vigentes están en `services/casos_historial.py`. Los dashboards,
filtros, seguimiento de token y reportes reutilizan el mismo análisis local.
La sincronización y los datos originales de ServiceNow no se modifican.

El tablero, el KPI y el resumen ejecutivo comparten **siete grupos como máximo**:
Firma digital; Instalación o configuración; Descarga o activación (incluye
reactivación); Uso o funcionamiento del token; Fallas del servicio;
Solicitudes y consultas; Enviado a otros canales. No se acumulan varios grupos
en una sola etiqueta. Los detalles se conservan por caso.

La lámina del KPI muestra solo los títulos de los grupos, sus cantidades y
barras. Las explicaciones de motivos, causas y destinos se conservan en el
detalle del reporte y en soporte. Las
familias técnicas agrupan explicaciones equivalentes; otras explicaciones
conservan el texto registrado. Se indica cuántos casos del grupo tienen causa
documentada y, si no hay diagnóstico, se señala sin inventarlo. En las
derivaciones también se muestran los destinos. Una combinación ventas/comercial
cuenta como una ruta de un caso. Estos detalles usan la misma cohorte y no se
suman a los totales del KPI. El bloque «¿Qué explica los grupos del KPI?» permite
revisar los conteos, y la descarga del detalle incluye la evidencia por caso.

Cada número de caso cuenta una sola vez, conservando la última fila cargada.
La **clasificación principal** describe el motivo del reporte, salvo las
derivaciones sin atención técnica registrada, que se agrupan exclusivamente
en **Enviado a otros canales**. Agenda, PQRS, ventas y comercial son destinos,
no tipologías adicionales. Una instalación derivada a agenda cuenta una vez
en Enviado a otros canales; su solicitud de instalación se conserva en el detalle.

La **tipología** resume esa clasificación en Soporte Uso, Incidentes,
Solicitudes o Enviado a otros canales. Las tipificaciones antiguas
se conservan como datos de origen y no determinan la clasificación vigente.

El **motivo inicial** usa la descripción y el primer correo entrante, antes de
las instrucciones de soporte o de respuestas posteriores del cliente. Una
solicitud clara sin una regla específica entra en Solicitudes y consultas y
conserva el texto como evidencia; no se muestra como Por revisar. La falta de
solicitud registrada se señala en el detalle, sin inventar una causa técnica.
Las respuestas «se atiende a través de nuestro canal de ventas» y «será
atendido directamente por nuestros comerciales» identifican una derivación,
sin interpretar promesas de acompañamiento como trabajo técnico realizado.

En **Operación → Agenda** se seleccionan los casos por la derivación registrada
en la conversación, con el mismo criterio para el histórico y la reincidencia.
Los motivos se agrupan por su clasificación inicial; instalación, descarga,
PIN u otros detalles se ven por caso con su evidencia. Token físico es un
atributo separado y no reemplaza el motivo de la solicitud.

**Producto**, **problema** y **causa raíz** son datos distintos:

- Token físico requiere una mención explícita positiva de «token físico»
  (incluidas variantes ortográficas del término). USB, ePass, driver, SafeNet
  y una instalación genérica no identifican por sí solos un token físico.
- Caídas, indisponibilidad, fallas y problemas de visualización de consumos
  entran en Fallas del servicio. Una consulta de consumos sin error entra en
  Solicitudes y consultas. El motivo concreto se conserva en el detalle.
- La causa raíz requiere una explicación registrada en el campo de causa o
  en la atención. Una instalación, un producto, una redirección o un síntoma
  no confirman una causa técnica. Las causas documentadas se consultan aparte.
  Se agrupan en cinco familias técnicas (controladores, certificados vencidos,
  bloqueo por intentos, incompatibilidad de software y restricciones del equipo)
  más Otras causas documentadas. La explicación exacta se conserva por caso;
  el mismo grupo no se repite por tener distintos estados de documentación.

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
