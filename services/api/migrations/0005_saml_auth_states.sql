-- Server-side state storage for SAML login flows (RelayState anti-CSRF token).

CREATE TABLE IF NOT EXISTS saml_auth_states (
  state text PRIMARY KEY,
  org_id uuid NOT NULL REFERENCES organizations(id) ON DELETE CASCADE,
  provider_id text NOT NULL,
  created_at timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS saml_auth_states_org_provider_idx
  ON saml_auth_states(org_id, provider_id, created_at DESC);

