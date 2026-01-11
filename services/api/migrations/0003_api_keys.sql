-- API keys for org-scoped programmatic access.

CREATE TABLE IF NOT EXISTS api_keys (
  id uuid PRIMARY KEY,
  org_id uuid NOT NULL REFERENCES organizations(id) ON DELETE CASCADE,
  name text NOT NULL,
  -- Stored as `sha256:<salt-hex>:<digest-hex>` (salt is random per key; digest = sha256(salt || secret)).
  -- The raw API key is returned once at creation time and is never persisted.
  key_hash text NOT NULL,
  created_by uuid NOT NULL REFERENCES users(id) ON DELETE RESTRICT,
  created_at timestamptz NOT NULL DEFAULT now(),
  last_used_at timestamptz,
  revoked_at timestamptz,
  UNIQUE (org_id, name)
);

CREATE INDEX IF NOT EXISTS api_keys_org_id_idx ON api_keys(org_id);
CREATE INDEX IF NOT EXISTS api_keys_org_created_idx ON api_keys(org_id, created_at DESC);
