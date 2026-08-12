-- Script B: verificación previa, exclusivamente de lectura.
SELECT role, active, COUNT(*) AS users
FROM app_users GROUP BY role, active ORDER BY role, active DESC;

SELECT COUNT(*) AS active_admins
FROM app_users WHERE role = 'admin' AND active = TRUE;

SELECT role, COUNT(*) AS users
FROM app_users
WHERE role IS NULL OR role NOT IN ('admin', 'soporte', 'experiencia', 'gerencias', 'area_ia', 'viewer')
GROUP BY role;

SELECT lower(btrim(email)) AS normalized_email, COUNT(*) AS duplicates
FROM app_users GROUP BY lower(btrim(email)) HAVING COUNT(*) > 1;

SELECT
  COUNT(*) FILTER (WHERE email IS NULL OR btrim(email) = '') AS missing_email,
  COUNT(*) FILTER (WHERE role IS NULL OR btrim(role) = '') AS missing_role,
  COUNT(*) FILTER (WHERE active IS NULL) AS missing_active,
  COUNT(*) FILTER (WHERE password_hash IS NULL OR btrim(password_hash) = '') AS missing_password_hash
FROM app_users;

SELECT email, active, created_at, last_login
FROM app_users WHERE role = 'viewer' ORDER BY email;
