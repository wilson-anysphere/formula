-- SAML 2.0 SSO support: per-org IdP providers.

CREATE TABLE IF NOT EXISTS org_saml_providers (
  org_id uuid NOT NULL REFERENCES organizations(id) ON DELETE CASCADE,
  provider_id text NOT NULL,
  idp_entry_point text NOT NULL,
  sp_entity_id text NOT NULL,
  idp_cert_pem text NOT NULL,
  want_assertions_signed boolean NOT NULL DEFAULT true,
  want_response_signed boolean NOT NULL DEFAULT false,
  enabled boolean NOT NULL DEFAULT true,
  attribute_mapping jsonb,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now(),
  PRIMARY KEY (org_id, provider_id)
);

-- Replay protection cache for SAML assertion IDs (post-signature validation).
CREATE TABLE IF NOT EXISTS saml_assertion_replays (
  assertion_id text PRIMARY KEY,
  org_id uuid NOT NULL REFERENCES organizations(id) ON DELETE CASCADE,
  provider_id text NOT NULL,
  expires_at timestamptz NOT NULL,
  created_at timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS saml_assertion_replays_expires_idx
  ON saml_assertion_replays(expires_at);
