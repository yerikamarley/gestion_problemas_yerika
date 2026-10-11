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

Los ejemplos del documento «Situaciones que se consideran envío a otros
canales» también cubren «lo manejamos exclusivamente a través del canal de
agendamiento», «programa tu cita», «es necesario agendar» y «canal de peticiones».
Peticiones/PQR/PQRF se consolidan en PQRS. Servicio al cliente y entidad
bancaria son destinos adicionales dentro de Enviado a otros canales.
La instrucción para descargar certificados en el portal de Certitítulo se
considera Autogestión, incluso si el caso queda cerrado; entregar el enlace
no prueba una intervención técnica. Se conserva el estado y la causa original.
Una verificación de estado para orientar al canal no equivale a instalar,
configurar o resolver una falla. El trabajo técnico realizado sí conserva la
clasificación del problema y se muestra como redirección con atención.
En observaciones sin autor estructurado se acepta el encabezado fechado
del comentario cuando identifica a un integrante del equipo de soporte.
Las menciones sueltas, mensajes del cliente, acuses y textos citados no
constituyen una derivación. «Es una instalación» sin destino tampoco basta.

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
  En Detalle y evidencia se muestran primero las explicaciones registradas,
  con cantidad de casos, ejemplos y evidencia. Se unen textos equivalentes
  que solo difieren en mayúsculas, acentos o espacios. La vista de grupos es
  opcional y conserva ejemplos de las explicaciones. Las familias técnicas
  incluyen controladores, certificados, descarga, configuración, dispositivo,
  autenticación, disponibilidad, lentitud, consumos, correo y documentos.
  Las explicaciones que no encajan se muestran como Sin agrupación técnica,
  conservando su texto y evidencia; no desaparecen en un único renglón de Otras.
  Estos grupos de detalle no amplían las siete clasificaciones del KPI.

En **Operación → Carga por agente**, la tabla sustituye la gráfica y muestra
total asignado, abiertos, cerrados y vencidos abiertos. El score de 0 a 100
es el cumplimiento SLA: cerrados que cumplen / cerrados evaluables × 100.
Los cierres sin fechas válidas se muestran sin evaluación; un agente sin
casos evaluables no tiene score. El cumplimiento total usa todos los casos
evaluables del período, no un promedio de porcentajes por agente.

En **Operación → Productos y principales causas**, se muestra una tabla
de todos los productos, sin agrupar los de menor volumen como Otros.
Cada producto incluye casos únicos, casos con y sin causa documentada,
las tres explicaciones documentadas más frecuentes y los principales motivos
iniciales en una columna separada. La tabla se puede descargar como CSV.

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

En **Resumen → ¿Por qué contactan por token físico?**, la cohorte coincide
con la gráfica de token físico: mismo período, soporte y sin asignación,
casos únicos y mención explícita positiva. Incluye casos que mencionan ambos
tipos de token. Las tres vistas muestran motivos de contacto, causas
documentadas y casos con evidencia. El motivo corresponde a la solicitud
inicial, incluso si después se redirige a agenda, ventas u otro canal.
Los motivos cuentan cada caso una vez y sus porcentajes usan toda la cohorte
de token físico. Las causas conservan la explicación registrada y sus
porcentajes usan solo los casos con causa documentada; se indica también
cuántos todavía no tienen diagnóstico. Las causas de casos que mencionan
ambos tokens se presentan como causas del caso, sin atribuirlas exclusivamente
al dispositivo físico. Cada vista permite descargar sus datos completos.
Este análisis reutiliza las notas y la conversación ya cargadas, sin nuevas
consultas a ServiceNow o a la base de datos.
