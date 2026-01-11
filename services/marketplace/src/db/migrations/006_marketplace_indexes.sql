CREATE INDEX IF NOT EXISTS idx_extensions_publisher
  ON extensions(publisher);

CREATE INDEX IF NOT EXISTS idx_package_scans_scanned_at
  ON package_scans(scanned_at DESC);

