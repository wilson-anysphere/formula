# Sync Server

`services/sync-server` is a production-oriented [y-websocket](https://github.com/yjs/y-websocket) server with:

- Opaque token or JWT (HS256) authentication
- Role-based enforcement for read-only users
- Awareness anti-spoofing / identity sanitization
- Connection attempt + message rate limiting
- Persistence to disk (file or LevelDB) with optional at-rest encryption
- Health endpoints (`/healthz`, `/readyz`) and internal admin endpoints under `/internal/*`
- Prometheus metrics (`/metrics`)

## Running

```bash
pnpm --filter @formula/sync-server dev
```

Required environment variables:

- `SYNC_SERVER_HOST` (default: `127.0.0.1`)
- `SYNC_SERVER_PORT` (default: `1234`)
- Authentication (pick one):
  - Opaque token:
    - `SYNC_SERVER_AUTH_TOKEN` (opaque token)
    - Optional: `SYNC_SERVER_AUTH_MODE=opaque`
  - JWT (HS256):
    - `SYNC_SERVER_JWT_SECRET`
    - Optional: `SYNC_SERVER_AUTH_MODE=jwt-hs256` (or `jwt`)
  - Token introspection (for centralized auth):
    - `SYNC_SERVER_AUTH_MODE=introspect`
    - `SYNC_SERVER_INTROSPECT_URL`
    - `SYNC_SERVER_INTROSPECT_TOKEN`

## Auth modes

### Opaque token

Set `SYNC_SERVER_AUTH_TOKEN` to a shared secret string.

Clients authenticate with:

- Query param: `?token=<token>`, or
- Header: `Authorization: Bearer <token>`

### JWT (HS256)

Set:

- `SYNC_SERVER_JWT_SECRET` (required)
- `SYNC_SERVER_JWT_AUDIENCE` (default: `formula-sync`)
- `SYNC_SERVER_JWT_ISSUER` (optional)

JWT payload must include:

- `sub` (user id)
- `docId` (document id / room name)
- `role` (`owner|admin|editor|commenter|viewer`)

Optional claims:

- `rangeRestrictions` (array) – when `SYNC_SERVER_ENFORCE_RANGE_RESTRICTIONS=1` (default: `true` in
  production), the sync-server validates that **incoming Yjs updates do not modify cells outside
  the allowed edit permissions**.

  Each entry must match the `@formula/collab-permissions` schema (see
  `packages/collab/permissions/normalizeRestriction`), e.g.:

  ```json
  {
    "rangeRestrictions": [
      {
        "sheetId": "Sheet1",
        "startRow": 0,
        "startCol": 0,
        "endRow": 0,
        "endCol": 0,
        "editAllowlist": ["user-123"]
      }
    ]
  }
  ```

  Notes:
  - `sheetName` is also accepted as an alias for `sheetId`.
  - Older clients may send `{ "range": { ... }, "editAllowlist": [...] }`, which is also accepted.

### Token introspection

Set:

- `SYNC_SERVER_AUTH_MODE=introspect`
- `SYNC_SERVER_INTROSPECT_URL` (base URL for the introspection service)
- `SYNC_SERVER_INTROSPECT_TOKEN` (sent as `x-internal-admin-token` to the introspection service)
- Optional:
  - `SYNC_SERVER_INTROSPECT_CACHE_MS` (default: `30000`)
  - `SYNC_SERVER_INTROSPECT_FAIL_OPEN` (default: `false`; ignored in production)

The sync-server will call:

```
POST ${SYNC_SERVER_INTROSPECT_URL}/internal/sync/introspect
Content-Type: application/json
x-internal-admin-token: ${SYNC_SERVER_INTROSPECT_TOKEN}

{ "token": "<client token>", "docId": "<requested doc id>" }
```

Expected response:

```json
{ "ok": true, "userId": "u1", "orgId": "o1", "role": "editor", "sessionId": "..." }
```

Introspection results are cached in-memory for `SYNC_SERVER_INTROSPECT_CACHE_MS`, so token revocation
may not take effect immediately.

## Persistence backends

Select with:

- `SYNC_SERVER_PERSISTENCE_BACKEND=leveldb` (default)
- `SYNC_SERVER_PERSISTENCE_BACKEND=file`

Additional knobs:

- `SYNC_SERVER_DATA_DIR` (default: `./.sync-server-data`)
- `SYNC_SERVER_PERSIST_COMPACT_AFTER_UPDATES` (default: `200`)

If `y-leveldb` is not installed, the server falls back to file persistence in non-production environments.

## At-rest encryption (KeyRing)

Enable with:

- `SYNC_SERVER_PERSISTENCE_ENCRYPTION=keyring`

Provide keys via one of:

- `SYNC_SERVER_ENCRYPTION_KEYRING_JSON` (KeyRing JSON)
- `SYNC_SERVER_ENCRYPTION_KEYRING_PATH` (path to KeyRing JSON)
- `SYNC_SERVER_PERSISTENCE_ENCRYPTION_KEY_B64` (base64-encoded 32-byte key; implies keyring mode)

KeyRing CLI helpers:

```bash
pnpm --filter @formula/sync-server keyring:generate
pnpm --filter @formula/sync-server keyring:rotate
pnpm --filter @formula/sync-server keyring:validate
```

## Retention & tombstones

Retention is optional (disabled by default):

- `SYNC_SERVER_RETENTION_TTL_MS` (document TTL; `0` disables retention)
- `SYNC_SERVER_RETENTION_SWEEP_INTERVAL_MS` (how often to sweep; `0` disables the periodic sweeper)
- `SYNC_SERVER_TOMBSTONE_TTL_MS` (how long to keep tombstones; defaults to the retention TTL if set)

Documents can also be tombstoned via the internal admin API; tombstoned docs are not served again.

## Internal admin endpoints

Internal endpoints are disabled by default. Enable with:

- `SYNC_SERVER_INTERNAL_ADMIN_TOKEN`

Then call endpoints with header:

```
x-internal-admin-token: <token>
```

Endpoints:

- `GET /internal/stats` – operational stats (connections, persistence stats, etc.)
- `POST /internal/retention/sweep` – trigger a retention/tombstone sweep
- `DELETE /internal/docs/:docId` – tombstone + purge a document and terminate active sockets
- `GET /internal/metrics` – same metrics as `/metrics`, but token-protected

## Prometheus metrics

The server exposes Prometheus text format at:

- `GET /metrics` (public)

Notable metrics (prefix `sync_server_`):

- `sync_server_ws_connections_total` / `sync_server_ws_connections_current`
- `sync_server_ws_connections_rejected_total{reason=...}`
- `sync_server_ws_messages_rate_limited_total`
- `sync_server_ws_messages_too_large_total`
- `sync_server_ws_closes_total{code=...}`
- `sync_server_retention_docs_purged_total{sweep=...}`
- `sync_server_retention_sweep_errors_total{sweep=...}`
- `sync_server_persistence_info{backend=...,encryption=...}`

## Limits & hardening

- Hard maximum websocket message size:

- `SYNC_SERVER_MAX_MESSAGE_BYTES` (default: `2097152` / 2 MiB)

- Per-connection message rate limiting:

  - `SYNC_SERVER_MAX_MESSAGES_PER_WINDOW`
  - `SYNC_SERVER_MESSAGE_WINDOW_MS`

- Per-document message rate limiting:

  - `SYNC_SERVER_MAX_MESSAGES_PER_DOC_WINDOW`
  - `SYNC_SERVER_DOC_MESSAGE_WINDOW_MS`

- Awareness payload limits:

  - `SYNC_SERVER_MAX_AWARENESS_STATE_BYTES`
  - `SYNC_SERVER_MAX_AWARENESS_ENTRIES`

- Optional JWT cell-range restriction enforcement (fail-closed):

  - `SYNC_SERVER_ENFORCE_RANGE_RESTRICTIONS` (default: `true` in production)

Limits are enforced both at the `ws` server (`maxPayload`) and defensively in message handlers.

## Optional HTTPS/WSS mode

If you deploy without a reverse proxy and want the server to terminate TLS itself, set:

- `SYNC_SERVER_TLS_CERT_PATH`
- `SYNC_SERVER_TLS_KEY_PATH`

When set, the server listens with HTTPS (and accepts WSS websocket upgrades).
