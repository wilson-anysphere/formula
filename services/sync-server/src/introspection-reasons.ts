/**
 * Shared mapping from API introspection "reason"/"error" values to HTTP status codes.
 *
 * sync-server uses the same API internal endpoint (`/internal/sync/introspect`) in two places:
 * - `SYNC_SERVER_AUTH_MODE=introspect` (auth mode)
 * - `SYNC_SERVER_INTROSPECTION_*` (optional JWT session revalidation)
 *
 * Keeping the mapping in one place ensures both paths make consistent decisions.
 */
export function statusCodeForIntrospectionReason(reason: string | undefined): 401 | 403 {
  switch (reason) {
    case "invalid_token":
    case "token_expired":
    case "invalid_claims":
    case "session_not_found":
    case "session_revoked":
    case "session_expired":
    case "session_user_mismatch":
    case "api_key_not_found":
    case "api_key_revoked":
    case "api_key_user_mismatch":
    case "api_key_org_mismatch":
      return 401;
    default:
      return 403;
  }
}
