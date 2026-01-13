# Sync Server Authentication

Formula’s sync server (Yjs / `services/sync-server`) accepts WebSocket connections for a given document ID (`ws://…/<docId>?token=…`).

## Background: sync JWTs

The API issues short-lived sync JWTs via:

- `POST /docs/:docId/sync-token`

These tokens are signed with `SYNC_TOKEN_SECRET` and include claims like:

- `sub` (userId)
- `docId`
- `orgId`
- `role` (document role at issuance time)
- optional `sessionId` (only present for cookie/session auth)
- optional `apiKeyId` (only present for API key auth)

## Auth modes

Sync server auth is configured via environment variables and maps to `services/sync-server/src/config.ts`.

### `jwt-hs256` (legacy / local verification)

In this mode, the sync server verifies the JWT signature locally and authorizes solely from token claims.

**Pros**

- No dependency on the API during WebSocket upgrade.
- Simple operationally.

**Cons**

- Token claims can become stale. A token minted before a membership change or session revocation can continue to grant access until it expires.

#### Optional JWT session revalidation (recommended)

Even in `jwt-hs256` mode, the sync server can optionally call the API internal introspection endpoint during WebSocket
upgrade to ensure revoked sessions / permission changes take effect quickly.

Configure with:

- `SYNC_SERVER_INTROSPECTION_URL=<api base url>` (example: `https://api.internal.example.com`) **or** full URL to
  `/internal/sync/introspect`
- `SYNC_SERVER_INTROSPECTION_TOKEN=<same as INTERNAL_ADMIN_TOKEN>`
- Optional: `SYNC_SERVER_INTROSPECTION_CACHE_TTL_MS` (default: `15000`; set to `0` to disable caching)

The sync server forwards `clientIp` and `userAgent` to the introspection endpoint when available.

### `introspect` (recommended for production)

In this mode, the sync server calls an API internal endpoint during WebSocket upgrade:

- `POST /internal/sync/introspect`

The API verifies the token signature + audience, revalidates the token against current DB state, and returns an
introspection-style response:

- `active: true` with `{ userId, orgId, role, sessionId?, rangeRestrictions? }` for valid tokens
- `active: false` with a string `reason` for invalid/revoked tokens

Revalidation checks include:

- session exists / not revoked / not expired (when `sessionId` claim is present)
- API key exists / not revoked (when `apiKeyId` claim is present)
- user is still an org member (`org_members`)
- user is still a document member (`document_members`)
- org IP allowlist (`org_settings.ip_allowlist`)
- org MFA enforcement (`org_settings.require_mfa`) for session-issued tokens
- role clamping: the token `role` is treated as an upper bound and clamped to the current DB role (demotions take effect without forcing token refresh)

The sync server then uses the returned `{ userId, orgId, role, sessionId?, rangeRestrictions? }` as the
authoritative `AuthContext`.

If `rangeRestrictions` are present and `SYNC_SERVER_ENFORCE_RANGE_RESTRICTIONS=1`, the sync server will
enforce per-cell edit permissions server-side (same schema as the JWT `rangeRestrictions` claim).

**Pros**

- Session revocation and document membership changes take effect quickly (within the cache TTL).
- Sync JWT TTL can be increased safely because permissions are revalidated against server-side state.

**Cons**

- Sync server depends on the API during WebSocket upgrade (mitigated by caching and retries).

## Configuration

### API

Enable internal endpoints by setting:

- `INTERNAL_ADMIN_TOKEN` (shared secret used by internal callers)

### Sync server: introspection mode

Required:

- `SYNC_SERVER_AUTH_MODE=introspect`
- `SYNC_SERVER_INTROSPECT_URL=<api base url>` (example: `https://api.internal.example.com`)
- `SYNC_SERVER_INTROSPECT_TOKEN=<same as INTERNAL_ADMIN_TOKEN>`

Optional:

- `SYNC_SERVER_INTROSPECT_CACHE_MS=30000` (default: `30000`)
- `SYNC_SERVER_INTROSPECT_FAIL_OPEN=true` (non-production only; **not** honored in production)

The sync server forwards `clientIp` and `userAgent` to the introspection endpoint when available.

## Operational notes

- The sync server caches successful introspection results in-memory keyed by a SHA-256 hash of `(token, docId, clientIp)`
  to reduce load on the API and prevent cached results from being replayed across documents or IPs.
- Introspection is fail-closed by default: if the API cannot be reached (timeouts, network errors, 5xx), the WebSocket upgrade is rejected.
- Revocation latency is bounded by `SYNC_SERVER_INTROSPECT_CACHE_MS`.
