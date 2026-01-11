-- OIDC SSO support: per-org providers, encrypted secret store, and identity links.

CREATE TABLE IF NOT EXISTS secrets (
  name text PRIMARY KEY,
  encrypted_value text NOT NULL,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS org_oidc_providers (
  org_id uuid NOT NULL REFERENCES organizations(id) ON DELETE CASCADE,
  provider_id text NOT NULL,
  issuer_url text NOT NULL,
  client_id text NOT NULL,
  scopes jsonb NOT NULL DEFAULT '["openid","email","profile"]'::jsonb,
  enabled boolean NOT NULL DEFAULT true,
  created_at timestamptz NOT NULL DEFAULT now(),
  PRIMARY KEY (org_id, provider_id)
);

CREATE TABLE IF NOT EXISTS user_identities (
  user_id uuid NOT NULL REFERENCES users(id) ON DELETE CASCADE,
  provider text NOT NULL,
  subject text NOT NULL,
  email text,
  org_id uuid NOT NULL REFERENCES organizations(id) ON DELETE CASCADE,
  PRIMARY KEY (org_id, provider, subject)
);

CREATE INDEX IF NOT EXISTS user_identities_user_id_idx ON user_identities(user_id);

-- Server-side state storage for OIDC authorization code + PKCE login flows.
CREATE TABLE IF NOT EXISTS oidc_auth_states (
  state text PRIMARY KEY,
  org_id uuid NOT NULL REFERENCES organizations(id) ON DELETE CASCADE,
  provider_id text NOT NULL,
  nonce text NOT NULL,
  pkce_verifier text NOT NULL,
  redirect_uri text NOT NULL,
  created_at timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS oidc_auth_states_org_provider_idx
  ON oidc_auth_states(org_id, provider_id, created_at DESC);
