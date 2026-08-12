-- NO EJECUTAR HASTA CONFIRMAR QUE NO EXISTEN USUARIOS VIEWER Y QUE EL CÓDIGO
-- DESPLEGADO YA NO LO NECESITA.
BEGIN;
DO $$
DECLARE constraint_record RECORD;
BEGIN
  IF EXISTS (SELECT 1 FROM app_users WHERE role = 'viewer') THEN
    RAISE EXCEPTION 'Aún existen usuarios viewer; retiro cancelado';
  END IF;

  FOR constraint_record IN
    SELECT conname FROM pg_constraint
    WHERE conrelid = 'app_users'::regclass
      AND contype = 'c'
      AND pg_get_constraintdef(oid) ~* '\mrole\M'
  LOOP
    EXECUTE format('ALTER TABLE app_users DROP CONSTRAINT %I', constraint_record.conname);
  END LOOP;

  ALTER TABLE app_users ADD CONSTRAINT app_users_role_valid_check
    CHECK (role IN ('admin', 'soporte', 'experiencia', 'gerencias', 'area_ia'));
END $$;
COMMIT;
