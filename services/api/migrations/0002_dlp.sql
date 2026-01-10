-- DLP policies + document classification metadata.

CREATE TABLE IF NOT EXISTS org_dlp_policies (
  org_id uuid PRIMARY KEY REFERENCES organizations(id) ON DELETE CASCADE,
  policy jsonb NOT NULL,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS document_dlp_policies (
  document_id uuid PRIMARY KEY REFERENCES documents(id) ON DELETE CASCADE,
  policy jsonb NOT NULL,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS document_classifications (
  id uuid PRIMARY KEY,
  document_id uuid NOT NULL REFERENCES documents(id) ON DELETE CASCADE,
  selector_key text NOT NULL,
  selector jsonb NOT NULL,
  classification jsonb NOT NULL,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now(),
  UNIQUE (document_id, selector_key)
);

CREATE INDEX IF NOT EXISTS document_classifications_document_id_idx
  ON document_classifications(document_id);

