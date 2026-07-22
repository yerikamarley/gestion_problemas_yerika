# Arquitectura del proyecto

La aplicación se está separando progresivamente del antiguo diseño concentrado en
`app_ui.py` y `app_logic.py`. Durante la transición, esos módulos conservan las
importaciones y firmas públicas existentes para evitar cambios funcionales.

## Capas

- `core/`: configuración, seguridad y utilidades transversales sin reglas de negocio.
- `config/`: catálogos y valores estáticos del dominio.
- `repositories/`: conexión y consultas a PostgreSQL/Supabase.
- `services/`: reglas de negocio, clasificación, ANS y preparación analítica.
- `components/`: componentes visuales reutilizables de Streamlit.
- `dashboards/`: composición de filtros, componentes y servicios para cada vista.
- `tests/`: regresiones unitarias que protegen cada extracción.

## Regla de dependencias

```text
app.py
  └── dashboards / app_ui (transición)
        ├── components
        └── services
              ├── repositories
              ├── config
              └── core
```

Una capa inferior no debe importar una capa superior. En particular:

- `repositories` no importa Streamlit ni dashboards.
- `services` no renderiza elementos visuales.
- `components` no ejecuta consultas SQL.
- `dashboards` coordina; no contiene SQL ni reglas extensas de negocio.

## Estrategia de migración

1. Extraer código sin cambiar firmas ni resultados.
2. Mantener reexportaciones temporales desde los módulos antiguos.
3. Añadir pruebas para el comportamiento extraído.
4. Cambiar consumidores hacia el nuevo módulo.
5. Eliminar la compatibilidad temporal solo cuando no queden referencias.

Antes de integrar una fase deben pasar:

```powershell
python -m unittest discover -s tests -v
python -m compileall -q .
```

