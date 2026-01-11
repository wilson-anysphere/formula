-- Envelope encryption for sensitive stored blobs (starting with document_versions.data).
--
-- This keeps backwards compatibility with existing plaintext rows by:
-- - Leaving document_versions.data in place for plaintext storage
-- - Adding parallel columns for an encrypted envelope representation
--
-- The application is responsible for:
-- - Encrypting on write when org_settings.cloud_encryption_at_rest = true
-- - Decrypting on read when encrypted columns are present
-- - Supporting mixed plaintext/encrypted rows during rollout

-- Track last key rotation time separately from org_settings.updated_at (which changes for many settings).
ALTER TABLE org_settings
  ADD COLUMN IF NOT EXISTS kms_key_rotated_at timestamptz NOT NULL DEFAULT now();

-- Encrypted envelope columns for document_versions.data
ALTER TABLE document_versions ADD COLUMN IF NOT EXISTS data_envelope_version integer;
ALTER TABLE document_versions ADD COLUMN IF NOT EXISTS data_algorithm text;
-- NOTE: These are stored as text instead of bytea because our unit tests use
-- pg-mem, which does not preserve arbitrary bytea values. In a real Postgres
-- deployment, consider switching ciphertext/iv/tag to bytea for storage
-- efficiency.
--
-- `data_encrypted_dek` is text as well:
-- - legacy envelope schema v1 stores base64-encoded bytes
-- - canonical envelope schema v2 stores a JSON-serialized wrapped-key object
ALTER TABLE document_versions ADD COLUMN IF NOT EXISTS data_ciphertext text;
ALTER TABLE document_versions ADD COLUMN IF NOT EXISTS data_iv text;
ALTER TABLE document_versions ADD COLUMN IF NOT EXISTS data_tag text;
ALTER TABLE document_versions ADD COLUMN IF NOT EXISTS data_encrypted_dek text;
ALTER TABLE document_versions ADD COLUMN IF NOT EXISTS data_kms_provider text;
ALTER TABLE document_versions ADD COLUMN IF NOT EXISTS data_kms_key_id text;
ALTER TABLE document_versions ADD COLUMN IF NOT EXISTS data_aad jsonb;
