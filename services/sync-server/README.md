# Sync Server

`services/sync-server` is a production-oriented [y-websocket](https://github.com/yjs/y-websocket) server with:

- Opaque token or JWT (HS256) authentication
- Role-based enforcement:
  - `viewer`: read-only (drops Yjs update writes)
  - `commenter`: comment-only (allows updates to the `comments` root; rejects other roots)
- Reserved-root mutation guard for internal versioning/branching metadata (see below)
- Awareness anti-spoofing / identity sanitization
- Connection attempt + message rate limiting
- Persistence to disk (file or LevelDB) with optional at-rest encryption
- Health endpoints (`/healthz`, `/readyz`) and internal admin endpoints under `/internal/*`
- Prometheus metrics (`/metrics`)

## Running

```bash
pnpm --filter @formula/sync-server dev
```

## Deployment

This server is designed to run in production behind a reverse proxy / load balancer and to persist
documents to local disk.

Code references:

- Env var parsing + defaults: [`src/config.ts`](./src/config.ts)
- HTTP/WebSocket handling and health endpoints: [`src/server.ts`](./src/server.ts)
- Formula collaboration semantics (reserved roots + message size limits):
  - [`docs/06-collaboration.md`](../../docs/06-collaboration.md)
  - [`instructions/collaboration.md`](../../instructions/collaboration.md)

### Horizontal scaling (y-websocket rooms)

`y-websocket` does **not** provide cross-instance fanout by default. That means:

- All websocket clients for a given document (`docId` / room name) must connect to the **same**
  sync-server instance.
- If clients for the same `docId` land on different instances, they will not see each other's
  updates (each instance maintains its own in-memory room state and persistence).

If you run multiple replicas, you must configure **sticky routing by `docId`** (i.e. the websocket
URL path `/<docId>`). Cookie-based or client-IP stickiness is **not** sufficient; the affinity must
be per-document.

If you need to freely load balance connections for the same `docId`, you must add an external
pubsub/broker layer (e.g. Redis) and modify the server to publish/subscribe document updates across
instances.

### Persistence + storage

The built-in persistence backends (`SYNC_SERVER_PERSISTENCE_BACKEND=file|leveldb`) write to
`SYNC_SERVER_DATA_DIR` on the **local filesystem**.

Implications:

- The data directory is **not shared** across pods/instances. If you scale horizontally without
  docId-sticky routing, different replicas will load and persist different versions of the same
  document.
- Do **not** mount the same `SYNC_SERVER_DATA_DIR` into multiple replicas at once (LevelDB is not a
  multi-writer database; sync-server also uses a lock file — see “Data directory locking” below).
- In Kubernetes, use a per-pod PersistentVolume (often a StatefulSet with `ReadWriteOnce`) if you
  need durability across pod restarts.

### Reverse proxy / ingress (WebSocket)

Your proxy must support WebSockets and long-lived connections:

- Forward websocket upgrade headers (`Upgrade` + `Connection: upgrade`) and use HTTP/1.1.
- Increase idle/read timeouts (many load balancers default to ~60s).
- If the proxy terminates TLS, run sync-server over HTTP/WS internally and expose WSS externally
  (or use the built-in TLS options described below).

Example nginx snippet:

```nginx
location / {
  proxy_http_version 1.1;
  proxy_set_header Upgrade $http_upgrade;
  proxy_set_header Connection "upgrade";

  # WebSockets are long-lived; bump timeouts above the defaults.
  proxy_read_timeout 3600s;
  proxy_send_timeout 3600s;

  proxy_pass http://sync_server_upstream;
}
```

### `SYNC_SERVER_TRUST_PROXY`

When running behind a **trusted** reverse proxy, set `SYNC_SERVER_TRUST_PROXY=1` so rate limiting
and per-IP connection limits use the first `X-Forwarded-For` address (instead of the proxy’s IP).

Only enable this when your proxy/ingress:

- overwrites or strips any client-supplied `X-Forwarded-*` headers, and
- is the only network path to the sync-server

Otherwise clients can spoof `X-Forwarded-For` to bypass per-IP limits. See the implementation in
[`src/server.ts`](./src/server.ts) (`pickIp`) and defaults in [`src/config.ts`](./src/config.ts).

### Security knobs

- **Public metrics:** `GET /metrics` is public by default. Disable it with
  `SYNC_SERVER_DISABLE_PUBLIC_METRICS=1` and scrape metrics via:
  - `GET /internal/metrics` protected by `SYNC_SERVER_INTERNAL_ADMIN_TOKEN`, or
  - proxy/network-layer allowlisting.
- **Internal admin endpoints:** set `SYNC_SERVER_INTERNAL_ADMIN_TOKEN` to enable `/internal/*`.
  Treat this as a production secret; avoid exposing these endpoints to the public internet.
- **WebSocket `Origin` allowlist:** set `SYNC_SERVER_ALLOWED_ORIGINS` to a comma-separated list of
  allowed `Origin` header values (e.g. `https://app.example.com,https://staging.example.com`).
  - If `SYNC_SERVER_ALLOWED_ORIGINS` is set and an `Origin` header is present on the websocket
    upgrade request, it must exactly match one of the allowlisted origins (after trimming
    whitespace).
  - If the `Origin` header is missing, the connection is allowed (to support non-browser clients).
  - Rejections use HTTP `403` with body `Origin not allowed` and increment
    `sync_server_ws_connections_rejected_total{reason="origin_not_allowed"}`.
  - If you need to restrict *all* websocket clients (including those without an `Origin` header),
    enforce allowlisting at your reverse proxy/ingress or network layer in addition to auth.

### Kubernetes probes (`/healthz`, `/readyz`)

The server exposes:

- `GET /healthz` – process liveness + a small operational snapshot (always `200` when the process is
  healthy).
- `GET /readyz` – readiness gate (returns `200` only after persistence/tombstones are initialized
  and the data-dir lock is held; otherwise `503`).

Suggested probes:

```yaml
livenessProbe:
  httpGet:
    path: /healthz
    port: 1234
  periodSeconds: 10
  timeoutSeconds: 2
readinessProbe:
  httpGet:
    path: /readyz
    port: 1234
  periodSeconds: 10
  timeoutSeconds: 2
```

For large datasets, consider adding a `startupProbe` so Kubernetes doesn't restart the pod during
initial persistence initialization.

### Graceful shutdown / drain mode

On `SIGTERM`/`SIGINT` (or when `server.stop()` is called), sync-server enters **drain mode** to
support rolling deploys without mass disconnects:

- `GET /readyz` returns `503` with JSON `{ "reason": "draining" }` so load balancers stop routing new
  connections to the instance.
- New websocket upgrades are rejected with HTTP `503` and increment
  `sync_server_ws_connections_rejected_total{reason="draining"}`.
- Existing websocket clients are allowed to remain connected for up to
  `SYNC_SERVER_SHUTDOWN_GRACE_MS` (default: `10000`). After the grace period expires, any remaining
  sockets are force-terminated and the process exits.

Monitor `sync_server_shutdown_draining_current` (set to `1` while draining).

### Formula deployment gotchas (reserved roots + message size)

These are the two most common collaboration misconfigurations:

- **Reserved roots:** The reserved-root mutation guard defaults to **enabled** when
  `NODE_ENV=production` and rejects writes to `versions`, `versionsMeta`, and `branching:*` (close
  code `1008`). If you use Yjs-based versioning/branching stores, disable it or use out-of-doc
  stores instead. See “Reserved root mutation guard” below and the deployment notes in
  [`docs/06-collaboration.md`](../../docs/06-collaboration.md).
- **Message size:** `SYNC_SERVER_MAX_MESSAGE_BYTES` defaults to **2 MiB**; large branching commits
  *and* large version snapshots (when stored in-doc) can exceed this and cause close code `1009`.
  See “Limits & hardening” below and [`docs/06-collaboration.md`](../../docs/06-collaboration.md)
  for chunking/streaming guidance.

## Stress testing

`services/sync-server` includes a **manual** stress/load harness that starts a local sync-server
instance and hammers it with many concurrent `y-websocket` clients.

Run with defaults:

```bash
pnpm -C services/sync-server stress
```

Common overrides:

```bash
# More clients + multiple docs (rooms)
pnpm -C services/sync-server stress -- --clients 200 --docs 10 --durationMs 30000

# Fixed number of ops per client, spread across a duration
pnpm -C services/sync-server stress -- --clients 50 --opsPerClient 2000 --durationMs 60000

# Use JWT auth instead of opaque tokens
pnpm -C services/sync-server stress -- --authMode jwt

# Force a smaller max message size (useful for validating 1009 behavior)
pnpm -C services/sync-server stress -- --maxMessageBytes 65536
```

Notes:

- The harness **does not** run under `pnpm test` (manual only).
- It spawns a sync-server child process using `src/index.ts` (via `tsx`), waits for `/healthz`,
  then connects `y-websocket` `WebsocketProvider` clients.
- Workload:
  - baseline “cell-ish” updates against a `cells` `Y.Map`
  - periodic awareness updates (presence churn)
  - a small fraction of writes to reserved-ish roots (`branching:*`, `versions`) to exercise those paths
- At the end it waits for **convergence** (all clients observe a final per-client counter map),
  prints a summary, and exits **non-zero** if convergence fails or if any disconnects / unexpected
  websocket close codes are observed.

Interpreting output:

- `throughput`: rough ops/sec based on successful local ops attempted by clients during the workload
- `convergenceTime`: how long it took for all clients to observe the final counters after the workload
- `wsCloseCodes`: counts of websocket close codes observed (watch for `1008`, `1009`, `1013`, `1006`)
- `metrics snapshot`: selected `/metrics` lines at the end of the run (if available)

Required environment variables:

- `SYNC_SERVER_HOST` (default: `127.0.0.1`)
- `SYNC_SERVER_PORT` (default: `1234`; must be between `0` and `65535`)
- `SYNC_SERVER_TRUST_PROXY` (default: `false`) – when running behind a reverse proxy, set this to
  `true` so rate limiting and per-IP connection limits use the `x-forwarded-for` header.
  **Only enable this when the proxy is trusted** (otherwise clients can spoof their IP).
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
- Optional strict claim enforcement (recommended in production):
  - `SYNC_SERVER_JWT_REQUIRE_SUB`
    - If `true`, require a non-empty `sub` claim (user id).
    - Defaults to **`true` in production** (`NODE_ENV=production`), **`false` otherwise**.
    - Missing/empty `sub` rejects the websocket upgrade with HTTP `403`.
  - `SYNC_SERVER_JWT_REQUIRE_EXP`
    - If `true`, require an `exp` claim (expiry time, unix seconds).
    - Defaults to **`true` in production** (`NODE_ENV=production`), **`false` otherwise**.
    - Missing `exp` rejects the websocket upgrade with HTTP `401`.
    - Note: `jsonwebtoken` validates expiration *if present*; this flag enforces that the claim exists
      at all.

JWT payload claims:

- `docId` (document id / room name)
- `role` (`owner|admin|editor|commenter|viewer`) – defaults to `editor` if omitted
- `sub` (user id) – required when `SYNC_SERVER_JWT_REQUIRE_SUB=1` (see above)
- `exp` (expiry time) – required when `SYNC_SERVER_JWT_REQUIRE_EXP=1` (see above)

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

#### Optional JWT session revalidation (introspection)

When using JWT auth, the sync-server can optionally call the API internal introspection endpoint
on each websocket upgrade to ensure the issuing session is still active and permissions haven't
been revoked.

Set:

- `SYNC_SERVER_INTROSPECTION_URL` (base URL like `https://api.internal.example.com` **or** full URL
  to `/internal/sync/introspect`)
- `SYNC_SERVER_INTROSPECTION_TOKEN` (sent as `x-internal-admin-token`)
- Optional: `SYNC_SERVER_INTROSPECTION_CACHE_TTL_MS` (default: `15000`; set to `0` to disable caching)
- Optional: `SYNC_SERVER_INTROSPECTION_MAX_CONCURRENT` (default: `50`; set to `0` to disable the limit)

The request/response format is the same as the token introspection auth mode described below.

The sync-server forwards `clientIp` and `userAgent` to the introspection endpoint when available.

Note: This runs **in addition to** local JWT verification. A JWT that passes signature/claims checks
can still be rejected if the introspection endpoint reports the session as inactive.

Caching notes:

- Introspection results are cached in-memory for `SYNC_SERVER_INTROSPECTION_CACHE_TTL_MS` (keyed by
  `(token, docId, clientIp)`), so session revocations may take up to the TTL to take effect.

Concurrency limiting / over-capacity behavior:

- If `SYNC_SERVER_INTROSPECTION_MAX_CONCURRENT` is non-zero and the server already has that many
  in-flight HTTP requests to the introspection endpoint, the websocket upgrade is rejected
  immediately with HTTP `503` ("Introspection over capacity"). This is intended to fail-fast and
  protect the API introspection endpoint during spikes.

Metrics (for monitoring/alerting):

- `sync_server_introspection_over_capacity_total` – count of websocket upgrades rejected due to the
  max-concurrent limit.
- `sync_server_introspection_requests_total{result=ok|inactive|error}` – count of introspection
  calls by result (watch `result="error"`).

### Token introspection (auth mode)

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

{
  "token": "<client token>",
  "docId": "<requested doc id>",
  "clientIp": "<client ip address>",
  "userAgent": "<user agent header>"
}
```

`clientIp` and `userAgent` are optional. When present, they allow the introspection service to enforce org
IP allowlists and record richer audit information.

Expected response:

```json
{
  "ok": true,
  "userId": "user-123",
  "orgId": "o1",
  "role": "editor",
  "sessionId": "...",
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

Optional fields:

- `sessionId` (string) – forwarded into the websocket `AuthContext` for audit/diagnostics.
- `rangeRestrictions` (array) – optional per-cell edit permissions, using the **same schema** as the
  JWT `rangeRestrictions` claim above (see `packages/collab/permissions/normalizeRestriction`).

  When `SYNC_SERVER_ENFORCE_RANGE_RESTRICTIONS=1` (default: `true` in production) and a non-empty
  `rangeRestrictions` array is provided, the sync-server validates that **incoming Yjs updates do not
  modify unauthorized cells** (violations close the websocket with code `1008`).

For compatibility, `{ "active": true, ... }` is also accepted (and `active: false` is treated as
inactive). When returning an inactive response, provide a string `reason` (or `error`) so the
sync-server can map it to an HTTP status code (`401` for invalid/expired sessions, `403` for access
denied).

Introspection results are cached in-memory for `SYNC_SERVER_INTROSPECT_CACHE_MS`, so token revocation
may not take effect immediately.

Cache keys are scoped per `(token, docId, clientIp)` so a token cannot be replayed from a different
client IP during the cache window.

## Persistence backends

Select with:

- `SYNC_SERVER_PERSISTENCE_BACKEND=leveldb` (default)
- `SYNC_SERVER_PERSISTENCE_BACKEND=file`

Additional knobs:

- `SYNC_SERVER_DATA_DIR` (default: `./.sync-server-data`)
- `SYNC_SERVER_PERSIST_COMPACT_AFTER_UPDATES` (default: `200`)
- `SYNC_SERVER_LEVELDB_DOCNAME_HASHING` (default: `false`) – hash `docName` before writing LevelDB keys to avoid
  storing raw document ids in the database.

If `y-leveldb` is not installed, the server falls back to file persistence in non-production environments.

## Data directory locking

By default the server creates a lock file (`.sync-server.lock`) in `SYNC_SERVER_DATA_DIR` to prevent multiple
processes from using the same persistence directory.

You can disable this (unsafe for multi-process deployments) with:

- `SYNC_SERVER_DISABLE_DATA_DIR_LOCK=true`

When locking is disabled, `/readyz` will return `503` with reason `data_dir_lock_disabled`.

## At-rest encryption (KeyRing)

Enable with:

- `SYNC_SERVER_PERSISTENCE_ENCRYPTION=keyring`

Provide keys via one of:

- `SYNC_SERVER_ENCRYPTION_KEYRING_JSON` (KeyRing JSON)
- `SYNC_SERVER_ENCRYPTION_KEYRING_PATH` (path to KeyRing JSON)
- `SYNC_SERVER_PERSISTENCE_ENCRYPTION_KEY_B64` (base64-encoded 32-byte key; implies keyring mode)

Optional:

- `SYNC_SERVER_PERSISTENCE_ENCRYPTION_STRICT` (default: `true` in production) – when `true`, legacy plaintext reads
  are rejected. When `false`, legacy plaintext documents can still be read to allow in-place migration.

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

- `GET /metrics` (public by default; set `SYNC_SERVER_DISABLE_PUBLIC_METRICS=1` to disable)
- Metrics responses include `Cache-Control: no-store` to discourage caching by proxies/CDNs.

Notable metrics (prefix `sync_server_`):

- `sync_server_ws_connections_total` / `sync_server_ws_connections_current`
- `sync_server_ws_active_docs_current` / `sync_server_ws_unique_ips_current`
- `sync_server_ws_connections_rejected_total{reason=...}`
- `sync_server_ws_message_bytes_total` / `sync_server_ws_message_bytes_rejected_total`
- `sync_server_ws_messages_rate_limited_total`
- `sync_server_ws_messages_too_large_total`
- `sync_server_ws_message_handler_errors_total{stage=...}`
- `sync_server_ws_reserved_root_mutations_total`
- `sync_server_ws_reserved_root_inspection_fail_closed_total{reason=...}`
- `sync_server_ws_awareness_spoof_attempts_total` / `sync_server_ws_awareness_client_id_collisions_total`
- `sync_server_ws_closes_total{code=...}`
- `sync_server_introspection_over_capacity_total`
- `sync_server_introspection_requests_total{result=ok|inactive|error}`
- `sync_server_introspection_request_duration_ms{path=...,result=...}`
- `sync_server_retention_docs_purged_total{sweep=...}`
- `sync_server_retention_sweep_errors_total{sweep=...}`
- `sync_server_persistence_info{backend=...,encryption=...}`
- `sync_server_process_resident_memory_bytes` / `sync_server_process_heap_used_bytes` / `sync_server_event_loop_delay_ms`

## Limits & hardening

- Hard maximum websocket message size:

  - `SYNC_SERVER_MAX_MESSAGE_BYTES` (default: `2097152` / 2 MiB). Set to `0` to disable (not recommended).

- WebSocket upgrade URL/token size limits (defense-in-depth against oversized request lines / auth headers):

  - `SYNC_SERVER_MAX_URL_BYTES` (default: `8192`; set to `0` to disable).
    - If exceeded, the websocket upgrade is rejected with HTTP `414` and increments
      `sync_server_ws_connections_rejected_total{reason="url_too_long"}`.
  - `SYNC_SERVER_MAX_TOKEN_BYTES` (default: `4096`; set to `0` to disable).
    - Applies to both `?token=<token>` and `Authorization: Bearer <token>`.
    - If exceeded, the websocket upgrade is rejected with HTTP `414` and increments
      `sync_server_ws_connections_rejected_total{reason="token_too_long"}`.

- Hard maximum document id (room name) size:
  - 1024 bytes (UTF-8). Requests exceeding this are rejected with HTTP `414` and increment
    `sync_server_ws_connections_rejected_total{reason="doc_id_too_long"}`.

- Connection limits:
  - `SYNC_SERVER_MAX_CONNECTIONS`
  - `SYNC_SERVER_MAX_CONNECTIONS_PER_IP`
  - `SYNC_SERVER_MAX_CONNECTIONS_PER_DOC` (default: `0` / unlimited)
  - Set any of these to `0` to disable.
  - Connection attempt rate limiting:
    - `SYNC_SERVER_MAX_CONN_ATTEMPTS_PER_WINDOW`
    - `SYNC_SERVER_CONN_ATTEMPT_WINDOW_MS`
    - Set either to `0` to disable.

- Per-connection message rate limiting:

  - `SYNC_SERVER_MAX_MESSAGES_PER_WINDOW`
  - `SYNC_SERVER_MESSAGE_WINDOW_MS`
  - Set either to `0` to disable.

- Per-IP aggregate message rate limiting (across all websocket connections):

  - `SYNC_SERVER_MAX_MESSAGES_PER_IP_WINDOW`
  - `SYNC_SERVER_IP_MESSAGE_WINDOW_MS`
  - Set either to `0` to disable.

- Per-document message rate limiting:

  - `SYNC_SERVER_MAX_MESSAGES_PER_DOC_WINDOW`
  - `SYNC_SERVER_DOC_MESSAGE_WINDOW_MS`
  - Set either to `0` to disable.

- Awareness payload limits:

  - `SYNC_SERVER_MAX_AWARENESS_STATE_BYTES`
  - `SYNC_SERVER_MAX_AWARENESS_ENTRIES`
  - Setting either to `0` effectively disables awareness updates (they will be dropped).

- Optional cell-range restriction enforcement (fail-closed):

  - `SYNC_SERVER_ENFORCE_RANGE_RESTRICTIONS` (default: `true` in production)
  - Applies to `rangeRestrictions` provided via:
    - JWT claims (`rangeRestrictions`)
    - Token introspection responses (`rangeRestrictions`)

### Reserved root mutation guard (versioning/branching roots)

`services/sync-server` can defensively reject Yjs updates that attempt to mutate **reserved top-level roots** inside the shared `Y.Doc`.

This exists because some deployments store **version history** and/or **branching metadata** outside the Yjs document (e.g. in an API/DB or local SQLite), and do not want arbitrary clients to be able to write to those internal roots.

Configuration:

- `SYNC_SERVER_RESERVED_ROOT_GUARD_ENABLED`
  - If set to `1`/`true`, the guard is enabled.
  - If set to `0`/`false`, the guard is disabled.
  - If **unset**, it defaults to:
    - **enabled** when `NODE_ENV=production`
    - **disabled** otherwise
- `SYNC_SERVER_RESERVED_ROOT_NAMES`
  - Optional comma-separated list of **exact** reserved root names.
  - Defaults to: `versions,versionsMeta`
- `SYNC_SERVER_RESERVED_ROOT_PREFIXES`
  - Optional comma-separated list of reserved root **prefixes**.
  - Defaults to: `branching:`

Default reserved roots:

- Exact root names:
  - `versions`
  - `versionsMeta`
- Root prefixes:
  - `branching:` (so `branching:branches`, `branching:commits`, `branching:meta`, etc.)

Behavior:

- When the guard detects a reserved-root write, the server closes the websocket with code `1008` and reason `"reserved root mutation"`.
- Implementation/config:
  - guard wiring + env var default: [`src/server.ts`](./src/server.ts)
  - update inspection + enforcement: [`src/ywsSecurity.ts`](./src/ywsSecurity.ts)

Implications for Formula collaboration deployments:

- `@formula/collab-versioning` defaults to `YjsVersionStore` (history stored *inside* the Y.Doc under `versions`/`versionsMeta`). This requires the guard to be **disabled**, otherwise version snapshot/checkpoint writes will be rejected.
- Formula’s Yjs-backed branching store (`YjsBranchStore`, graph stored *inside* the Y.Doc under `branching:*`) also requires the guard to be **disabled**.
- If you customize the branching root name (non-default `rootName`), note that the guard only blocks configured prefixes; you may need to extend `SYNC_SERVER_RESERVED_ROOT_PREFIXES` accordingly.
- If you want to keep the guard enabled in production, use non-Yjs stores instead (where applicable), for example:
  - `ApiVersionStore` (cloud DB/API): `packages/versioning/src/store/apiVersionStore.js`
  - `SQLiteVersionStore` (local desktop): `packages/versioning/src/store/sqliteVersionStore.js`
  - `SQLiteBranchStore` (local desktop): `packages/versioning/branches/src/store/SQLiteBranchStore.js`

Limits are enforced both at the `ws` server (`maxPayload`) and defensively in message handlers.

Websocket compression (`permessage-deflate`) is disabled by default as defense-in-depth against
compression bombs.

## Optional HTTPS/WSS mode

If you deploy without a reverse proxy and want the server to terminate TLS itself, set:

- `SYNC_SERVER_TLS_CERT_PATH`
- `SYNC_SERVER_TLS_KEY_PATH`

When set, the server listens with HTTPS (and accepts WSS websocket upgrades).

## Reserved roots (versions / branching metadata)

The sync-server treats some Yjs roots as "reserved" (`versions`, `versionsMeta`, `branching:*`) so
untrusted clients cannot write unbounded metadata into the shared document.

By default, the reserved-root guard is:

- **enabled in production**
- **disabled in dev/test** (to keep local workflows compatible)

Override with:

- `SYNC_SERVER_RESERVED_ROOT_GUARD_ENABLED` (`true|false`)

If you disable the reserved-root guard (e.g. to allow Yjs-based versioning/branching stores),
sync-server can enforce per-document quotas to prevent unbounded growth:

- `SYNC_SERVER_MAX_VERSIONS_PER_DOC` (default: `500` in production; `0` otherwise)
- `SYNC_SERVER_MAX_BRANCHING_COMMITS_PER_DOC` (default: `5000` in production; `0` otherwise)

Set either value to `0` to disable that limit.
