-- Per-org SIEM endpoint configuration.
--
-- Used by:
-- - Fastify SIEM config routes (services/api/src/routes/siem.ts)
-- - Background SIEM export worker (services/api/src/siem/*)

CREATE TABLE IF NOT EXISTS org_siem_configs (
  org_id uuid PRIMARY KEY REFERENCES organizations(id) ON DELETE CASCADE,
  enabled boolean NOT NULL DEFAULT false,
  -- Stores non-secret config plus secret references (never plaintext).
  config jsonb NOT NULL,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now()
);

-- Quickly find orgs with SIEM export enabled.
CREATE INDEX IF NOT EXISTS org_siem_configs_enabled_true_idx
  ON org_siem_configs(enabled)
  WHERE enabled = true;
