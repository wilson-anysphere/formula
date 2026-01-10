-- Enterprise security policies: encryption options, data residency extensions,
-- audit log archiving, and legal holds.

-- Org-level policy fields
ALTER TABLE org_settings
  ADD COLUMN IF NOT EXISTS cloud_encryption_at_rest boolean NOT NULL DEFAULT true;

ALTER TABLE org_settings
  ADD COLUMN IF NOT EXISTS kms_provider text NOT NULL DEFAULT 'local';

ALTER TABLE org_settings
  ADD COLUMN IF NOT EXISTS kms_key_id text;

ALTER TABLE org_settings
  ADD COLUMN IF NOT EXISTS key_rotation_days integer NOT NULL DEFAULT 90;

ALTER TABLE org_settings
  ADD COLUMN IF NOT EXISTS certificate_pinning_enabled boolean NOT NULL DEFAULT false;

ALTER TABLE org_settings
  ADD COLUMN IF NOT EXISTS certificate_pins jsonb NOT NULL DEFAULT '[]'::jsonb;

ALTER TABLE org_settings
  ADD COLUMN IF NOT EXISTS data_residency_allowed_regions jsonb;

ALTER TABLE org_settings
  ADD COLUMN IF NOT EXISTS legal_hold_overrides_retention boolean NOT NULL DEFAULT true;

-- Audit log archive table (move cold events out of hot storage)
CREATE TABLE IF NOT EXISTS audit_log_archive (
  id uuid PRIMARY KEY,
  org_id uuid,
  user_id uuid,
  user_email text,
  event_type text NOT NULL,
  resource_type text NOT NULL,
  resource_id text,
  ip_address text,
  user_agent text,
  session_id uuid,
  success boolean NOT NULL,
  error_code text,
  error_message text,
  details jsonb NOT NULL DEFAULT '{}'::jsonb,
  created_at timestamptz NOT NULL,
  archived_at timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS audit_log_archive_org_created_idx
  ON audit_log_archive(org_id, created_at DESC);

CREATE INDEX IF NOT EXISTS audit_log_archive_event_type_idx
  ON audit_log_archive(event_type);

-- Document-level legal holds (override retention deletion/archiving when enabled).
CREATE TABLE IF NOT EXISTS document_legal_holds (
  document_id uuid PRIMARY KEY REFERENCES documents(id) ON DELETE CASCADE,
  org_id uuid NOT NULL REFERENCES organizations(id) ON DELETE CASCADE,
  enabled boolean NOT NULL DEFAULT true,
  reason text,
  created_by uuid REFERENCES users(id) ON DELETE SET NULL,
  created_at timestamptz NOT NULL DEFAULT now(),
  released_by uuid REFERENCES users(id) ON DELETE SET NULL,
  released_at timestamptz
);

CREATE INDEX IF NOT EXISTS document_legal_holds_org_enabled_idx
  ON document_legal_holds(org_id, enabled);

