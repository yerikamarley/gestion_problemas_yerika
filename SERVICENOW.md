# ServiceNow: OAuth, casos e historial de la conversación

La pantalla **Sincronizar Casos** utiliza la API y sustituye la carga manual de
Excel para casos. Los registros nuevos se insertan y los existentes se
actualizan en PostgreSQL/Supabase, sin borrar los meses anteriores. Los
usuarios del portal consultan esa base compartida y no necesitan una cuenta
en ServiceNow. Los filtros existentes se conservan; la vista de casos abre
por defecto en **Hoy**, con fechas de Colombia.

## Configuración de ServiceNow

La integración entrante debe usar **OAuth — Concesión de credenciales de
cliente**, estar activa y tener un **Usuario de aplicación OAuth** con lectura
de `sn_customerservice_case`, `task_sla`, `sys_journal_field` y `sys_email`,
incluidos los campos de notas y los cuerpos de correo. Para incidentes se
requiere también lectura de `incident`. La propiedad
`glide.oauth.inbound.client.credential.grant_type.enabled` debe estar habilitada.
Véanse las guías oficiales de [credenciales de cliente](https://www.servicenow.com/docs/r/platform-security/authentication/configure-an-oauth-client-credential-grant.html)
y [usuario de aplicación OAuth](https://www.servicenow.com/docs/r/platform-security/authentication/add-oauth-application-user.html).

Este flujo obtiene el token con el ID y el secreto de la integración. No usa
una URL de redireccionamiento, el código de autorización de otra integración
ni el inicio de sesión personal con biometría/MFA. El usuario de aplicación
define los permisos efectivos de lectura; una respuesta HTTP 200 puede seguir
ocultando registros o campos por las ACLs de ServiceNow. Para comprobar la
integridad, contrastar un caso conocido, sus correos y sus notas con lo visible
en ServiceNow bajo esos permisos.

## Secretos del servidor y activación

En Streamlit Community Cloud, abrir **App settings → Secrets**, conservar las
claves existentes de la aplicación y añadir estas claves de nivel raíz:

```toml
SERVICENOW_INSTANCE = "certicamara.service-now.com"
SERVICENOW_AUTH_TYPE = "client_credentials"
SERVICENOW_CLIENT_ID = "ID_DE_LA_INTEGRACION"
SERVICENOW_CLIENT_SECRET = "SECRETO_DE_LA_INTEGRACION"
SERVICENOW_SYNC_INTERVAL_SECONDS = "120"
```

El secreto corresponde a **Secreto de cliente** de la integración de
credenciales de cliente. No se publica en Git ni se introduce en un formulario
para otros roles. La conexión a la base continúa usando
`SUPABASE_DATABASE_URL`. Para ejecución local, `.env.example` documenta las
claves; completar únicamente el `.env` privado. Las variables de entorno
tienen prioridad sobre los secretos de Streamlit.

Después de desplegar estos cambios, iniciar sesión como Admin del portal,
abrir **Sincronizar Casos** y pulsar **Validar conexión y activar actualización
automática**. La aplicación solicita un token y valida las cuatro fuentes.
Revalida el rol Admin antes de guardar la activación en
`servicenow_sync_state`. Allí se guarda el administrador local responsable,
sin credenciales ni tokens. La activación sobrevive al reinicio del servidor.

Los tokens se reutilizan únicamente en RAM hasta poco antes de expirar. El
servidor solicita otro token automáticamente y, ante HTTP 401, renueva y
reintenta una sola vez. Si vuelve a fallar la autenticación, el trabajador
espera a que un Admin revise y valide la integración. HTTP 403 indica que se
deben revisar los permisos de la fuente. Los errores públicos omiten los
secretos y los cuerpos de respuesta remotos.

Si ServiceNow rechaza el token, el mensaje muestra el estado HTTP y únicamente
los códigos OAuth conocidos. `invalid_client` orienta a comprobar el ID y el
secreto de la misma integración; `unsupported_grant_type`, a comprobar la
propiedad que habilita credenciales de cliente; `invalid_scope`, al alcance
solicitado. `server_error` con `access_denied` orienta a revisar el usuario de
aplicación y la habilitación de la concesión. El texto libre de
`error_description` nunca se muestra. Para precisar la causa en la instancia,
reintentar desde el portal y revisar **System Logs → All**, filtrando la fuente
por OAuth y ordenando por fecha descendente. Véase la sección de
[diagnóstico del SDK oficial](https://servicenow.github.io/sdk/config/ci-integration#troubleshooting).

## Historial de la conversación y análisis

Los correos y las entradas de `sys_journal_field` forman el **Historial de la
conversación**, ordenado por fecha e ID remoto. Incluye notas de trabajo,
comentarios, autor, asunto, destinatarios y mensaje. Cuando el correo solo
tiene cuerpo HTML, se extrae texto legible y se omiten estilos y scripts.

El historial estructurado se guarda una sola vez en el JSON de
`servicenow_records`. La columna adicional `cases.historial_conversacion`
guarda su texto analítico, calculado durante la sincronización. La interfaz
consulta únicamente el historial del registro solicitado mediante el índice
existente `(tabla, numero)`, sin descargar el snapshot completo ni llamar a
ServiceNow al abrir la conversación. El endpoint de snapshots mantiene el
alias anterior `traza_unificada` para compatibilidad.

El historial participa en la tipificación, la tipología de soporte, las causas
comunes, el análisis de agendamiento y la búsqueda de casos. Se añade a la
evidencia disponible de descripciones y observaciones; las columnas históricas
de Excel se conservan. Una actualización sin conversación no vacía el
historial previamente guardado ni elimina esa evidencia de la clasificación.

Al arrancar se añade la columna de texto y el estado de activación mediante
migraciones aditivas. No se crea una tabla adicional por cada fuente de la API.
Las tablas auxiliares tienen RLS habilitada sin políticas públicas: la conexión
PostgreSQL del backend debe ser su propietario o tener `BYPASSRLS`.

## Sincronización y rendimiento

El trabajador actualiza cada 120 segundos, configurable con mínimo de 30.
La primera ejecución parte del comienzo del día en Colombia; las siguientes
parten de la última ejecución correcta guardada, con dos minutos de
solapamiento. También busca cambios en SLAs, notas y correos para actualizar
casos antiguos que reciben actividad reciente.

Cada página guarda la fila analítica y el snapshot en una misma transacción.
El cursor solo avanza cuando terminan todas las fuentes y páginas. Si una
consulta falla, se conservan las páginas confirmadas y se reintenta desde el
cursor anterior. Un bloqueo PostgreSQL evita ciclos simultáneos.

Las fuentes relacionadas se consultan por lotes de hasta 100 casos. En la
prueba con una página de casos y como máximo una página de cada fuente,
consolidar 100 casos requiere 4 consultas GET en lugar de 301. Si las fuentes
relacionadas tienen más páginas, se consultan todas. La autenticación tiene
sus propias llamadas POST, reutilizadas mientras el token sea válido.

Las vistas leen la base local con filtros y caché. El estado se comprueba cada
30 segundos; un ciclo sin registros no vacía la caché. Cuando hay datos nuevos,
se invalidan las cargas de casos para mostrar los cambios. Se elimina además
una consulta redundante de metadatos por lectura filtrada. Estas mejoras se
verifican con pruebas; los tiempos reales dependen del volumen, los índices,
la latencia y las ACLs de la instancia.

La actualización continúa al cerrar sesión el Admin mientras el servidor
permanezca activo. Si el alojamiento suspende o reinicia la aplicación, el
trabajador también se detiene y retoma el cursor al volver a arrancar. Para
sincronización continua con el portal suspendido, ejecutar la tarea en un
servidor o programador externo disponible permanentemente.

## Ejecución y carga de históricos

```powershell
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
streamlit run app.py
```

Desde la pantalla Admin se puede completar un **Mes** mediante la API. Para
una tarea de servidor, indicar un administrador local activo autorizado:

```powershell
python -m sync_servicenow --actor admin@empresa.com --mes 2026-09
python -m sync_servicenow --actor admin@empresa.com
```

Sin `--mes` ni `--numero`, casos usa el cursor incremental. El proceso de la
tarea necesita las mismas credenciales OAuth y conexión PostgreSQL privadas.
Los endpoints de `backend.py` siguen autenticando usuarios locales del portal.
La compatibilidad interna con HTTP Basic y las funciones de importación Excel
anteriores se mantienen; el formulario de casos usa únicamente la integración
OAuth del servidor.
