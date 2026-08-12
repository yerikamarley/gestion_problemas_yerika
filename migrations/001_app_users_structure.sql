-- Script A: estructura compatible con viewer durante la transición.
BEGIN;

ALTER TABLE app_users ADD COLUMN IF NOT EXISTS name TEXT;
ALTER TABLE app_users ADD COLUMN IF NOT EXISTS updated_at TIMESTAMPTZ;

UPDATE app_users
SET updated_at = CURRENT_TIMESTAMP
WHERE updated_at IS NULL;

DO $$
DECLARE constraint_record RECORD;
BEGIN
  FOR constraint_record IN
    SELECT conname
    FROM pg_constraint
    WHERE conrelid = 'app_users'::regclass
      AND contype = 'c'
      AND pg_get_constraintdef(oid) ~* '\mrole\M'
  LOOP
    EXECUTE format('ALTER TABLE app_users DROP CONSTRAINT %I', constraint_record.conname);
  END LOOP;

  IF NOT EXISTS (
    SELECT 1 FROM pg_constraint
    WHERE conrelid = 'app_users'::regclass
      AND conname = 'app_users_role_valid_check'
  ) THEN
    ALTER TABLE app_users ADD CONSTRAINT app_users_role_valid_check
      CHECK (role IN ('admin', 'soporte', 'experiencia', 'gerencias', 'area_ia', 'viewer'))
      NOT VALID;
  END IF;
END $$;

ALTER TABLE app_users VALIDATE CONSTRAINT app_users_role_valid_check;
CREATE INDEX IF NOT EXISTS idx_app_users_role_active ON app_users (role, active);
CREATE UNIQUE INDEX IF NOT EXISTS idx_app_users_email_normalized
  ON app_users (lower(btrim(email)));

COMMIT;
