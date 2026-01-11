-- Add audit metadata columns to the per-org SIEM export config table.
--
-- Note: `org_siem_configs` is created by `0004_siem_configs.sql`; this migration
-- runs after it and only adds columns. Avoid `CREATE TABLE IF NOT EXISTS` here
-- so pg-mem (unit tests) doesn't choke on a no-op CREATE TABLE.
ALTER TABLE org_siem_configs
  ADD COLUMN IF NOT EXISTS created_by uuid REFERENCES users(id) ON DELETE SET NULL,
  ADD COLUMN IF NOT EXISTS updated_by uuid REFERENCES users(id) ON DELETE SET NULL;
