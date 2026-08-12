-- Script D: verificación posterior, exclusivamente de lectura.
SELECT role, active, COUNT(*) AS users
FROM app_users GROUP BY role, active ORDER BY role, active DESC;

SELECT COUNT(*) AS active_admins,
       (COUNT(*) > 0) AS has_active_admin
FROM app_users WHERE role = 'admin' AND active = TRUE;

SELECT role, COUNT(*) AS users
FROM app_users
WHERE role IS NULL OR role NOT IN ('admin', 'soporte', 'experiencia', 'gerencias', 'area_ia', 'viewer')
GROUP BY role;

SELECT COUNT(*) AS remaining_viewers FROM app_users WHERE role = 'viewer';

SELECT column_name, data_type, is_nullable, column_default
FROM information_schema.columns
WHERE table_schema = 'public' AND table_name = 'app_users'
  AND column_name IN ('name', 'updated_at')
ORDER BY column_name;
