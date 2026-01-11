-- SCIM provisioning tokens (org-scoped).
--
-- Tokens are stored as `sha256:<salt-hex>:<digest-hex>` where:
--   digest = sha256(salt || secret)
-- The raw token is returned once at creation time and is never persisted.
--
-- Upstream initially shipped `org_scim_tokens` as a single-token-per-org table keyed by `org_id`.
-- This migration upgrades it to support multiple named tokens per org.

ALTER TABLE org_scim_tokens ADD COLUMN IF NOT EXISTS id uuid;
ALTER TABLE org_scim_tokens ADD COLUMN IF NOT EXISTS name text;
ALTER TABLE org_scim_tokens ADD COLUMN IF NOT EXISTS created_by uuid;
ALTER TABLE org_scim_tokens ADD COLUMN IF NOT EXISTS last_used_at timestamptz;

-- Backfill existing single-token rows.
UPDATE org_scim_tokens SET id = org_id WHERE id IS NULL;
UPDATE org_scim_tokens SET name = 'default' WHERE name IS NULL;
UPDATE org_scim_tokens
SET created_by = picked.user_id
FROM (
  SELECT om.org_id AS member_org_id, om.user_id
  FROM org_members om
  JOIN (
    SELECT org_id, MIN(created_at) AS first_created_at
    FROM org_members
    GROUP BY org_id
  ) first_member
    ON first_member.org_id = om.org_id AND first_member.first_created_at = om.created_at
) picked
WHERE org_id = picked.member_org_id AND created_by IS NULL;

ALTER TABLE org_scim_tokens ALTER COLUMN id SET NOT NULL;
ALTER TABLE org_scim_tokens ALTER COLUMN name SET NOT NULL;
ALTER TABLE org_scim_tokens ALTER COLUMN created_by SET NOT NULL;

ALTER TABLE org_scim_tokens
  ADD CONSTRAINT org_scim_tokens_created_by_fkey
  FOREIGN KEY (created_by) REFERENCES users(id) ON DELETE RESTRICT;

-- Switch primary key from `org_id` -> `id` to enable multiple tokens per org.
ALTER TABLE org_scim_tokens DROP CONSTRAINT IF EXISTS org_scim_tokens_pkey;
ALTER TABLE org_scim_tokens ADD PRIMARY KEY (id);

ALTER TABLE org_scim_tokens
  ADD CONSTRAINT org_scim_tokens_org_id_name_key
  UNIQUE (org_id, name);

CREATE INDEX IF NOT EXISTS org_scim_tokens_org_id_idx ON org_scim_tokens(org_id);
CREATE INDEX IF NOT EXISTS org_scim_tokens_org_created_idx ON org_scim_tokens(org_id, created_at DESC);
