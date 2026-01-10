-- Document share links (public/private) with expirations.

CREATE TABLE IF NOT EXISTS document_share_links (
  id uuid PRIMARY KEY,
  document_id uuid NOT NULL REFERENCES documents(id) ON DELETE CASCADE,
  token_hash text NOT NULL UNIQUE,
  visibility text NOT NULL CHECK (visibility IN ('public', 'private')),
  role text NOT NULL CHECK (role IN ('editor', 'commenter', 'viewer')),
  created_by uuid REFERENCES users(id),
  created_at timestamptz NOT NULL DEFAULT now(),
  expires_at timestamptz,
  revoked_at timestamptz
);

CREATE INDEX IF NOT EXISTS document_share_links_document_id_idx ON document_share_links(document_id);
CREATE INDEX IF NOT EXISTS document_share_links_token_hash_idx ON document_share_links(token_hash);

