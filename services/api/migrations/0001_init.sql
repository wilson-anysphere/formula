-- Initial schema for enterprise/cloud backend foundation.

CREATE TABLE IF NOT EXISTS users (
  id uuid PRIMARY KEY,
  email text NOT NULL UNIQUE,
  name text NOT NULL,
  password_hash text,
  mfa_totp_secret text,
  mfa_totp_enabled boolean NOT NULL DEFAULT false,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS organizations (
  id uuid PRIMARY KEY,
  name text NOT NULL,
  created_at timestamptz NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS org_settings (
  org_id uuid PRIMARY KEY REFERENCES organizations(id) ON DELETE CASCADE,
  require_mfa boolean NOT NULL DEFAULT false,
  allowed_auth_methods jsonb NOT NULL DEFAULT '["password"]'::jsonb,
  ip_allowlist jsonb,
  allow_external_sharing boolean NOT NULL DEFAULT true,
  allow_public_links boolean NOT NULL DEFAULT true,
  default_permission text NOT NULL DEFAULT 'viewer' CHECK (default_permission IN ('viewer', 'commenter', 'editor')),
  ai_enabled boolean NOT NULL DEFAULT false,
  ai_data_processing_consent boolean NOT NULL DEFAULT false,
  data_residency_region text NOT NULL DEFAULT 'us',
  allow_cross_region_processing boolean NOT NULL DEFAULT true,
  ai_processing_region text,
  audit_log_retention_days integer NOT NULL DEFAULT 365,
  document_version_retention_days integer NOT NULL DEFAULT 365,
  deleted_document_retention_days integer NOT NULL DEFAULT 30,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS org_members (
  org_id uuid NOT NULL REFERENCES organizations(id) ON DELETE CASCADE,
  user_id uuid NOT NULL REFERENCES users(id) ON DELETE CASCADE,
  role text NOT NULL CHECK (role IN ('owner', 'admin', 'member')),
  created_at timestamptz NOT NULL DEFAULT now(),
  PRIMARY KEY (org_id, user_id)
);

CREATE TABLE IF NOT EXISTS documents (
  id uuid PRIMARY KEY,
  org_id uuid NOT NULL REFERENCES organizations(id) ON DELETE CASCADE,
  title text NOT NULL,
  created_by uuid NOT NULL REFERENCES users(id),
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now(),
  deleted_at timestamptz
);

CREATE INDEX IF NOT EXISTS documents_org_id_idx ON documents(org_id);

CREATE TABLE IF NOT EXISTS document_members (
  document_id uuid NOT NULL REFERENCES documents(id) ON DELETE CASCADE,
  user_id uuid NOT NULL REFERENCES users(id) ON DELETE CASCADE,
  role text NOT NULL CHECK (role IN ('owner', 'admin', 'editor', 'commenter', 'viewer')),
  created_at timestamptz NOT NULL DEFAULT now(),
  created_by uuid REFERENCES users(id),
  PRIMARY KEY (document_id, user_id)
);

CREATE INDEX IF NOT EXISTS document_members_user_id_idx ON document_members(user_id);

CREATE TABLE IF NOT EXISTS document_range_permissions (
  id uuid PRIMARY KEY,
  document_id uuid NOT NULL REFERENCES documents(id) ON DELETE CASCADE,
  sheet_name text NOT NULL,
  start_row integer NOT NULL,
  start_col integer NOT NULL,
  end_row integer NOT NULL,
  end_col integer NOT NULL,
  permission_type text NOT NULL CHECK (permission_type IN ('read', 'edit')),
  allowed_user_id uuid NOT NULL REFERENCES users(id) ON DELETE CASCADE,
  created_by uuid REFERENCES users(id),
  created_at timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS drp_document_id_idx ON document_range_permissions(document_id);
CREATE INDEX IF NOT EXISTS drp_allowed_user_idx ON document_range_permissions(allowed_user_id);

CREATE TABLE IF NOT EXISTS sessions (
  id uuid PRIMARY KEY,
  user_id uuid NOT NULL REFERENCES users(id) ON DELETE CASCADE,
  token_hash text NOT NULL UNIQUE,
  created_at timestamptz NOT NULL DEFAULT now(),
  expires_at timestamptz NOT NULL,
  revoked_at timestamptz,
  last_used_at timestamptz,
  ip_address text,
  user_agent text
);

CREATE INDEX IF NOT EXISTS sessions_user_id_idx ON sessions(user_id);

CREATE TABLE IF NOT EXISTS document_versions (
  id uuid PRIMARY KEY,
  document_id uuid NOT NULL REFERENCES documents(id) ON DELETE CASCADE,
  created_at timestamptz NOT NULL DEFAULT now(),
  created_by uuid REFERENCES users(id),
  description text,
  data bytea
);

CREATE INDEX IF NOT EXISTS document_versions_doc_created_idx ON document_versions(document_id, created_at DESC);

CREATE TABLE IF NOT EXISTS audit_log (
  id uuid PRIMARY KEY,
  org_id uuid REFERENCES organizations(id) ON DELETE CASCADE,
  user_id uuid REFERENCES users(id) ON DELETE SET NULL,
  user_email text,
  event_type text NOT NULL,
  resource_type text NOT NULL,
  resource_id text,
  ip_address text,
  user_agent text,
  session_id uuid REFERENCES sessions(id) ON DELETE SET NULL,
  success boolean NOT NULL,
  error_code text,
  error_message text,
  details jsonb NOT NULL DEFAULT '{}'::jsonb,
  created_at timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS audit_log_org_created_idx ON audit_log(org_id, created_at DESC);
CREATE INDEX IF NOT EXISTS audit_log_event_type_idx ON audit_log(event_type);

