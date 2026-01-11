-- Per-org SIEM endpoint configuration.
--
-- Used by:
-- - Fastify SIEM config routes (services/api/src/routes/siem.ts)
-- - Background SIEM export worker (services/api/src/siem/*)

CREATE TABLE IF NOT EXISTS org_siem_configs (
  org_id uuid PRIMARY KEY REFERENCES organizations(id) ON DELETE CASCADE,
  enabled boolean NOT NULL DEFAULT false,
  config jsonb NOT NULL DEFAULT '{}'::jsonb,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS org_siem_configs_enabled_idx
  ON org_siem_configs(enabled);

