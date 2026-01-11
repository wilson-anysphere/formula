-- Persistent cache for SAML AuthnRequest IDs (InResponseTo validation).

CREATE TABLE IF NOT EXISTS saml_request_cache (
  id text PRIMARY KEY,
  value text NOT NULL,
  created_at timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS saml_request_cache_created_at_idx
  ON saml_request_cache(created_at DESC);

