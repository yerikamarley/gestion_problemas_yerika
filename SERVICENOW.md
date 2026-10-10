# Conexión y actualización de casos

La pantalla **Sincronizar Casos** sustituye la carga manual de Excel. No se
eliminan casos ni meses anteriores: el backend inserta registros nuevos y
actualiza los existentes en PostgreSQL/Supabase. Los filtros siguen consultando
la base compartida y abren por defecto en **Hoy**.

## Conectar como administrador

1. Iniciar sesión en la aplicación con un administrador local activo.
2. Abrir **Sincronizar Casos → Conectar ServiceNow**.
3. Introducir usuario y contraseña de ServiceNow. Si la API exige MFA, introducir
   también el código de verificación vigente.
4. Seleccionar **Conectar y activar actualización automática**.

El dominio de la instancia se configura exclusivamente en el servidor mediante
`SERVICENOW_INSTANCE`, sin protocolo ni ruta. `.env.example` contiene las claves.
La base usa la conexión `SUPABASE_DATABASE_URL` ya existente. El usuario remoto
necesita lectura de casos, `task_sla`, `sys_journal_field` y `sys_email`, incluidos
los campos necesarios. Una respuesta HTTP 200 no garantiza que las ACLs permitan
ver todas las filas; esto debe confirmarse con el administrador de ServiceNow.

La aplicación valida las cuatro fuentes antes de guardar la conexión. Revalida
el rol Admin local antes de consultar y otra vez antes de guardar. Las
credenciales se guardan mediante reemplazo atómico del `.env` privado del
servidor, conservando las otras variables. No se guardan en el JSON analítico,
en la base compartida ni en resultados públicos. La contraseña y el código se
vacían en el formulario después del envío. Los otros roles no necesitan
iniciar sesión ni tener una cuenta en ServiceNow.

## Verificación multifactor

Las páginas web `validate_multifactor_auth_code.do` y `navpage.do` no son
endpoints REST. La política de acceso web puede diferir de la política API.
El cliente usa HTTP Basic y admite `X-MFA_TOKEN` si la API lo exige. El código
no se almacena; las cookies de sesión API se conservan solo en RAM del servidor.
Cuando ServiceNow solicita renovar MFA, el trabajador deja de intentar hasta
que un Admin reconecte. Reiniciar el proceso puede requerir un código nuevo.

Si la instancia no permite reutilizar sesiones API o exige MFA en cada llamada,
para operar sin intervención se necesita una integración OAuth autorizada en
ServiceNow. El formulario no configura OAuth ni inicia sesión mediante las
páginas web. Véase [autenticación REST oficial](https://www.servicenow.com/docs/r/api-reference/rest-api-explorer/c_RESTAPI.html).

## Actualización automática

La conexión activa `SERVICENOW_SYNC_ENABLED=true` y asocia la ingesta al Admin
que la configuró mediante `SERVICENOW_SYNC_ACTOR`. El servidor Streamlit ejecuta
un ciclo cada 120 segundos, configurable mediante
`SERVICENOW_SYNC_INTERVAL_SECONDS` (mínimo 30). Sigue funcionando al cerrar la
sesión del administrador, mientras el proceso servidor siga activo.

La primera ejecución busca cambios desde el inicio del día en Colombia. Las
siguientes parten de la última ejecución correcta guardada en
`servicenow_sync_state`, con dos minutos de solapamiento. Se consultan cambios
de casos, SLAs, notas y correos; también se actualizan casos antiguos que
reciben actividad reciente, aunque no cambie la fecha del caso.

El cursor avanza solo al terminar todas las fuentes y páginas. Si falla una
consulta, se conservan las páginas ya guardadas y se reintenta desde el cursor
anterior. Cada página guarda conjuntamente las filas analíticas y su JSON. Un
bloqueo PostgreSQL impide ejecutar dos ciclos automáticos simultáneos.

Los snapshots conservan valores brutos y legibles, SLAs y actividades ordenadas.
Cada actividad conserva su ID remoto. Si la API no devuelve actividades, no se
vacían las columnas de notas históricas de Excel. El contenido remoto de meses
anteriores se puede completar desde la opción **Mes** mediante la API.

Las tablas nuevas habilitan RLS sin políticas públicas. La conexión PostgreSQL
del backend debe ser su propietario o tener `BYPASSRLS`. La vista de casos
comprueba la última actualización cada 30 segundos y refresca los datos cuando
cambia, manteniendo los filtros elegidos.

## Ejecución

```powershell
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
streamlit run app.py
```

El arranque crea las tablas auxiliares sin eliminar datos. Para completar un
histórico o ejecutar un ciclo incremental desde una tarea de servidor:

```powershell
python -m sync_servicenow --mes 2026-09
python -m sync_servicenow
```

Sin `--mes` ni `--numero`, casos usa el cursor incremental; incidentes conserva
el comportamiento mensual anterior. Una tarea ejecutada en otro proceso no
comparte las cookies MFA en RAM y necesita acceso API aceptado por la instancia.
Los endpoints de `backend.py` siguen disponibles para clientes autorizados,
autenticando usuarios locales de la aplicación y manteniendo los secretos
remotos exclusivamente en el servidor.
