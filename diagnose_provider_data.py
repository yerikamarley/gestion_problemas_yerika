import app_logic


app_logic.init_db()
conn = app_logic.get_conn()
try:
    columns = conn.execute(
        """
        SELECT column_name
        FROM information_schema.columns
        WHERE table_name = 'incidents'
          AND column_name IN ('escalado_proveedor', 'nombre_proveedor')
        ORDER BY column_name
        """
    ).fetchall()
    counts = conn.execute(
        "SELECT COUNT(*), COUNT(NULLIF(nombre_proveedor, '')) FROM incidents"
    ).fetchone()
    providers = conn.execute(
        """
        SELECT numero, escalado_proveedor, nombre_proveedor
        FROM incidents
        WHERE COALESCE(nombre_proveedor, '') <> ''
        ORDER BY numero
        """
    ).fetchall()
    print("columns", columns)
    print("counts", counts)
    print("providers", providers)
finally:
    conn.close()
