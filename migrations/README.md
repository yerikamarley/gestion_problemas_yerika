# Migración manual de roles

Estos scripts se ejecutan manualmente en Supabase. No contienen asignaciones personales.

Orden seguro:

1. Respaldar o exportar `app_users` por un canal seguro.
2. Ejecutar `002_app_users_precheck.sql` y resolver duplicados, valores nulos o roles desconocidos.
3. Ejecutar `001_app_users_structure.sql`.
4. Desplegar el código compatible con `viewer`.
5. Probar inicio de sesión, menú, carga y administración con un Admin activo.
6. Copiar `003_app_users_role_assignment_template.sql`, reemplazar sus marcadores y ejecutar la copia revisada.
7. Probar individualmente los cinco roles con cuentas de prueba autorizadas.
8. Ejecutar `004_app_users_postcheck.sql` y conservar el resultado como evidencia.
9. En una fase futura, retirar `viewer` del código; solo después ejecutar `005_remove_viewer_future.sql`.

No ejecutar el script 005 mientras exista un registro `viewer` o el código desplegado todavía acepte ese rol.

Para habilitar los módulos de problemas, ejecutar `006_problems_table.sql` antes de desplegar la versión que los contiene.
