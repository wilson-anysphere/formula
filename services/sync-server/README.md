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
- One of:
  - `SYNC_SERVER_AUTH_TOKEN` (opaque token), or
  - `SYNC_SERVER_JWT_SECRET` (HS256 JWT)

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
- `sync_server_retention_docs_purged_total{sweep=...}`
- `sync_server_persistence_info{backend=...,encryption=...}`

## Message size limits

Set a hard maximum websocket message size with:

- `SYNC_SERVER_MAX_MESSAGE_BYTES` (default: `2097152` / 2 MiB)

Enforced at the `ws` server (`maxPayload`) and defensively in the message handler.

## Optional HTTPS/WSS mode

If you deploy without a reverse proxy and want the server to terminate TLS itself, set:

- `SYNC_SERVER_TLS_CERT_PATH`
- `SYNC_SERVER_TLS_KEY_PATH`

When set, the server listens with HTTPS (and accepts WSS websocket upgrades).
