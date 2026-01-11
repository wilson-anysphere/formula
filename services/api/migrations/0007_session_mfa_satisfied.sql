-- Track whether a session has satisfied MFA.
--
-- This is used for org-level `require_mfa` enforcement across sensitive endpoints
-- and supports both:
-- - local TOTP MFA (password login / enrollment), and
-- - upstream SSO MFA (OIDC/SAML amr/acr/assertion signals).
ALTER TABLE sessions
ADD COLUMN mfa_satisfied boolean NOT NULL DEFAULT false;

-- Best-effort backfill: sessions for users that currently have local TOTP enabled
-- are assumed to have satisfied MFA during login (since password login requires it).
UPDATE sessions
SET mfa_satisfied = true
WHERE user_id IN (SELECT id FROM users WHERE mfa_totp_enabled = true);
