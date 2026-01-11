-- MFA hardening: move TOTP secrets to encrypted secret store + recovery codes.

-- TOTP secrets were historically stored plaintext in `users.mfa_totp_secret`.
-- Rename the column so application code no longer reads/writes it directly.
ALTER TABLE users RENAME COLUMN mfa_totp_secret TO mfa_totp_secret_legacy;

CREATE TABLE IF NOT EXISTS user_mfa_recovery_codes (
  id uuid PRIMARY KEY,
  user_id uuid NOT NULL REFERENCES users(id) ON DELETE CASCADE,
  code_hash text NOT NULL,
  created_at timestamptz NOT NULL DEFAULT now(),
  used_at timestamptz
);

CREATE INDEX IF NOT EXISTS user_mfa_recovery_codes_user_id_idx ON user_mfa_recovery_codes(user_id);
CREATE INDEX IF NOT EXISTS user_mfa_recovery_codes_user_unused_idx ON user_mfa_recovery_codes(user_id, used_at);
