ALTER TABLE publishers ADD COLUMN revoked INTEGER NOT NULL DEFAULT 0;
ALTER TABLE publishers ADD COLUMN revoked_at TEXT;

CREATE TABLE IF NOT EXISTS package_scans (
  extension_id TEXT NOT NULL,
  version TEXT NOT NULL,
  status TEXT NOT NULL,
  findings_json TEXT NOT NULL DEFAULT '[]',
  scanned_at TEXT,
  PRIMARY KEY (extension_id, version),
  FOREIGN KEY (extension_id, version) REFERENCES extension_versions(extension_id, version) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_package_scans_status
  ON package_scans(status);
