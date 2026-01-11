ALTER TABLE extension_versions ADD COLUMN format_version INTEGER NOT NULL DEFAULT 1;
ALTER TABLE extension_versions ADD COLUMN file_count INTEGER NOT NULL DEFAULT 0;
ALTER TABLE extension_versions ADD COLUMN unpacked_size INTEGER NOT NULL DEFAULT 0;
ALTER TABLE extension_versions ADD COLUMN files_json TEXT NOT NULL DEFAULT '[]';

