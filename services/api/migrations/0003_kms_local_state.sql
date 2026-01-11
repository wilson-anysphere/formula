-- Persisted state for the local (dev/test) KMS provider.

CREATE TABLE IF NOT EXISTS org_kms_local_state (
  org_id uuid PRIMARY KEY REFERENCES organizations(id) ON DELETE CASCADE,
  provider jsonb NOT NULL,
  updated_at timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS org_kms_local_state_updated_at_idx
  ON org_kms_local_state(updated_at);

