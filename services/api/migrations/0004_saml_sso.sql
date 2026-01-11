-- SAML 2.0 SSO support: per-org IdP providers.

CREATE TABLE IF NOT EXISTS org_saml_providers (
  org_id uuid NOT NULL REFERENCES organizations(id) ON DELETE CASCADE,
  provider_id text NOT NULL,
  entry_point text NOT NULL,
  issuer text NOT NULL,
  idp_cert_pem text NOT NULL,
  want_assertions_signed boolean NOT NULL DEFAULT true,
  want_response_signed boolean NOT NULL DEFAULT true,
  attribute_mapping jsonb NOT NULL,
  enabled boolean NOT NULL DEFAULT true,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now(),
  PRIMARY KEY (org_id, provider_id)
);

