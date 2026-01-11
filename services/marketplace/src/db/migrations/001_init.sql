PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS publishers (
  publisher TEXT PRIMARY KEY,
  token_sha256 TEXT NOT NULL UNIQUE,
  public_key_pem TEXT NOT NULL,
  verified INTEGER NOT NULL DEFAULT 0,
  created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS extensions (
  id TEXT PRIMARY KEY,
  name TEXT NOT NULL,
  display_name TEXT NOT NULL,
  publisher TEXT NOT NULL,
  description TEXT NOT NULL DEFAULT '',
  categories_json TEXT NOT NULL DEFAULT '[]',
  tags_json TEXT NOT NULL DEFAULT '[]',
  screenshots_json TEXT NOT NULL DEFAULT '[]',
  verified INTEGER NOT NULL DEFAULT 0,
  featured INTEGER NOT NULL DEFAULT 0,
  deprecated INTEGER NOT NULL DEFAULT 0,
  blocked INTEGER NOT NULL DEFAULT 0,
  malicious INTEGER NOT NULL DEFAULT 0,
  download_count INTEGER NOT NULL DEFAULT 0,
  created_at TEXT NOT NULL,
  updated_at TEXT NOT NULL,
  FOREIGN KEY (publisher) REFERENCES publishers(publisher) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS extension_versions (
  extension_id TEXT NOT NULL,
  version TEXT NOT NULL,
  sha256 TEXT NOT NULL,
  signature_base64 TEXT NOT NULL,
  manifest_json TEXT NOT NULL,
  readme TEXT NOT NULL DEFAULT '',
  package_bytes BLOB NOT NULL,
  uploaded_at TEXT NOT NULL,
  yanked INTEGER NOT NULL DEFAULT 0,
  yanked_at TEXT,
  PRIMARY KEY (extension_id, version),
  FOREIGN KEY (extension_id) REFERENCES extensions(id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_extension_versions_extension
  ON extension_versions(extension_id);

CREATE TABLE IF NOT EXISTS audit_log (
  id TEXT PRIMARY KEY,
  actor TEXT NOT NULL,
  action TEXT NOT NULL,
  extension_id TEXT,
  version TEXT,
  ip TEXT,
  details_json TEXT NOT NULL DEFAULT '{}',
  created_at TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_audit_log_created_at
  ON audit_log(created_at DESC);

