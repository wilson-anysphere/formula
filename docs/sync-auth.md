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

## Auth modes

Sync server auth is configured via environment variables and maps to `services/sync-server/src/config.ts`.

### `jwt-hs256` (legacy / local verification)

In this mode, the sync server verifies the JWT signature locally and authorizes solely from token claims.

**Pros**

- No dependency on the API during WebSocket upgrade.
- Simple operationally.

**Cons**

- Token claims can become stale. A token minted before a membership change or session revocation can continue to grant access until it expires.

### `introspect` (recommended for production)

In this mode, the sync server calls an API internal endpoint during WebSocket upgrade:

- `POST /internal/sync/introspect`

The API verifies the token signature + audience, checks (if present) that the session is still valid, and recomputes the user’s current `document_members` role.

The sync server then uses the returned `{ userId, orgId, role }` as the authoritative `AuthContext`.

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

## Operational notes

- The sync server caches successful introspection results in-memory keyed by a SHA-256 hash of the token to reduce load on the API.
- Introspection is fail-closed by default: if the API cannot be reached (timeouts, network errors, 5xx), the WebSocket upgrade is rejected.
- Revocation latency is bounded by `SYNC_SERVER_INTROSPECT_CACHE_MS`.

