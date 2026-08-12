-- Migración manual: tabla de problemas. No contiene datos personales.
BEGIN;
CREATE TABLE IF NOT EXISTS problems (
    numero TEXT PRIMARY KEY,
    declaracion_problema TEXT,
    creado TEXT,
    descripcion TEXT,
    solucion_temporal TEXT,
    notas_cierre TEXT,
    diagnosticado_solucion_temporal TEXT,
    asignado_a TEXT,
    comentarios TEXT,
    estado TEXT,
    incidentes_relacionados INTEGER NOT NULL DEFAULT 0,
    notas_trabajo TEXT,
    observaciones_trabajo TEXT,
    prioridad TEXT
);
CREATE INDEX IF NOT EXISTS idx_problems_creado ON problems (creado);
CREATE INDEX IF NOT EXISTS idx_problems_estado ON problems (estado);
CREATE INDEX IF NOT EXISTS idx_problems_prioridad ON problems (prioridad);
CREATE INDEX IF NOT EXISTS idx_problems_asignado ON problems (asignado_a);
COMMIT;
