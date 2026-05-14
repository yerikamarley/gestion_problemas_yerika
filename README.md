# Gestion de problemas

Aplicacion Streamlit para cargar, clasificar y revisar casos e incidentes.

## Estructura

- `app.py`: punto de entrada de Streamlit.
- `app_ui.py`: vistas, dashboards y componentes visuales.
- `app_logic.py`: base de datos, carga de Excel, normalizacion y reglas de clasificacion.
- `scripts/`: scripts auxiliares para analisis y generacion de reportes.
- `.streamlit/config.toml`: tema visual de la app.

## Archivos locales que no se suben

Estos archivos se conservan en la maquina, pero no deben ir al repositorio:

- `data.db`: base SQLite local con casos, incidentes y usuarios.
- `outputs/`: reportes generados por scripts.
- `backups/`: copias locales de respaldo.
- `node_modules/`, `__pycache__/`, logs y archivos `.pyc`.

## Datos privados en despliegue

El repositorio puede ser publico, pero la base real debe mantenerse fuera de Git. La app usa `data.db` local por defecto y tambien acepta una ruta privada mediante `APP_DB_PATH`.

Variables recomendadas para despliegue:

```powershell
$env:APP_DB_PATH="C:\ruta\privada\data.db"
$env:APP_ADMIN_EMAIL="admin@tu-dominio.com"
$env:APP_ADMIN_PASSWORD="una-contrasena-segura"
```

Si la base ya tiene usuarios, la app no crea un admin nuevo. Si la base esta vacia, exige `APP_ADMIN_EMAIL` y `APP_ADMIN_PASSWORD` para evitar credenciales quemadas en un repositorio publico.

## Ejecutar localmente

```powershell
python -m streamlit run app.py
```

## Scripts con archivos fuente

Los scripts que leen Excel buscan los archivos en `Downloads` por defecto. Tambien aceptan variables de entorno o un argumento de ruta cuando aplica.

Ejemplo:

```powershell
$env:INCIDENTES_XLSX="C:\ruta\archivo.xlsx"
python scripts\analyze_incidentes_openpyxl.py
```
