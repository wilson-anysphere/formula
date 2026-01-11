-- SCIM provisioning tokens (one per organization).

CREATE TABLE IF NOT EXISTS org_scim_tokens (
  org_id uuid PRIMARY KEY REFERENCES organizations(id) ON DELETE CASCADE,
  -- Stored as `sha256:<salt-hex>:<digest-hex>` (salt is random per token; digest = sha256(salt || secret)).
  -- The raw SCIM token is returned once at creation time and is never persisted.
  token_hash text NOT NULL,
  created_at timestamptz NOT NULL DEFAULT now(),
  revoked_at timestamptz
);

