-- Script C: plantilla. Reemplazar marcadores; no añadir filas no aprobadas.
-- Revisar el SELECT final antes de ejecutar COMMIT.
BEGIN;
LOCK TABLE app_users IN SHARE ROW EXCLUSIVE MODE;

CREATE TEMP TABLE requested_role_changes (
  email TEXT PRIMARY KEY,
  role TEXT NOT NULL CHECK (role IN ('admin', 'soporte', 'experiencia', 'gerencias', 'area_ia'))
) ON COMMIT DROP;

INSERT INTO requested_role_changes (email, role) VALUES
  ('<CORREO_1>', '<ROL_INTERNO_1>'),
  ('<CORREO_2>', '<ROL_INTERNO_2>');

DO $$
BEGIN
  IF EXISTS (
    SELECT 1 FROM requested_role_changes r
    LEFT JOIN app_users u ON lower(btrim(u.email)) = lower(btrim(r.email))
    WHERE u.email IS NULL
  ) THEN
    RAISE EXCEPTION 'Hay correos solicitados que no existen; transacción cancelada';
  END IF;
END $$;

UPDATE app_users u
SET role = r.role, updated_at = CURRENT_TIMESTAMP
FROM requested_role_changes r
WHERE lower(btrim(u.email)) = lower(btrim(r.email));

DO $$
BEGIN
  IF NOT EXISTS (SELECT 1 FROM app_users WHERE role = 'admin' AND active = TRUE) THEN
    RAISE EXCEPTION 'La asignación dejaría el sistema sin Admin activo';
  END IF;
END $$;

SELECT u.email, u.role, u.active, u.updated_at
FROM app_users u
JOIN requested_role_changes r ON lower(btrim(u.email)) = lower(btrim(r.email))
ORDER BY u.email;

-- Detenerse aquí para revisar el resultado si se ejecuta de forma interactiva.
COMMIT;
