CREATE TABLE IF NOT EXISTS publisher_keys (
  id TEXT PRIMARY KEY,
  publisher TEXT NOT NULL,
  public_key_pem TEXT NOT NULL,
  created_at TEXT NOT NULL,
  revoked INTEGER NOT NULL DEFAULT 0,
  revoked_at TEXT,
  is_primary INTEGER NOT NULL DEFAULT 0,
  FOREIGN KEY (publisher) REFERENCES publishers(publisher) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_publisher_keys_publisher
  ON publisher_keys(publisher);

CREATE UNIQUE INDEX IF NOT EXISTS idx_publisher_keys_primary
  ON publisher_keys(publisher)
  WHERE is_primary = 1;

ALTER TABLE extension_versions ADD COLUMN signing_key_id TEXT REFERENCES publisher_keys(id);
ALTER TABLE extension_versions ADD COLUMN signing_public_key_pem TEXT;

