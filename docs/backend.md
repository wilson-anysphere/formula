# Backend (API + Sync server) local development

This repo contains a minimal-but-extensible **enterprise/cloud backend foundation**:

- `services/api`: Fastify + Postgres API service (auth, orgs, docs, RBAC, audit, sync token issuance)
- `services/sync-server`: Production Yjs sync server (`y-websocket`) that validates short-lived collaboration tokens issued by the API

## Quickstart

Prereqs:
- Docker + Docker Compose (use `docker compose`)

Start the local stack:

```bash
docker compose up --build
```

Services (default ports shown; see “Port overrides” below):
- API: http://localhost:3000
- Sync WS: ws://localhost:1234/<docId>?token=... (y-websocket protocol)
- Sync health: http://localhost:1234/healthz
- Postgres: localhost:5432 (user/pass/db = `postgres` / `postgres` / `formula`)

The API automatically runs SQL migrations on startup.

## Security-related configuration

### Local docker-compose vs production

`docker-compose.yml` is intended for **local development** and runs the API with:

- `NODE_ENV=development`
- `COOKIE_SECURE=false`
- development secrets for `SYNC_TOKEN_SECRET` / secret store keyring unless overridden

The production API Docker image sets `NODE_ENV=production`, and the API will **fail fast**
on insecure defaults (for example `COOKIE_SECURE!=true` or known dev secrets). In particular,
production requires:

- `COOKIE_SECURE=true`
- `SYNC_TOKEN_SECRET` set to a non-dev value
- secret store keyring configured via `SECRET_STORE_KEYS` (recommended) or `SECRET_STORE_KEYS_JSON` (also supported)
  - legacy: `SECRET_STORE_KEY` is still supported for smooth upgrades
- (Legacy-only) `LOCAL_KMS_MASTER_KEY` set to a non-dev value if you need to decrypt/migrate historical encrypted rows
- (Optional, AWS KMS) `AWS_KMS_ENABLED=true` + `AWS_REGION` when using `org_settings.kms_provider = 'aws'`
  - Ensure `@aws-sdk/client-kms` is installed in the API runtime image

For SSO deployments (OIDC/SAML), also set:

- `PUBLIC_BASE_URL` to the canonical external API origin (must be `https://...` in production, e.g. `https://api.example.com`)

### CORS

The API uses an **allowlist** for CORS:

- `CORS_ALLOWED_ORIGINS` — comma-separated list of allowed origins (e.g. `https://app.example.com`)
  - In production: defaults to **no allowed origins** unless explicitly set.
  - In dev/test: defaults to common localhost origins (`http://localhost:5173`, `http://localhost:3000`, etc).

### Client IP (rate limiting + IP allowlists)

Several protections depend on the derived client IP (`request.ip`):

- auth rate limiting (brute-force protection)
- org `ip_allowlist` enforcement (enterprise)
- OIDC redirect URI derivation (when deployed behind a reverse proxy)
- OIDC and SAML SSO endpoint rate limiting (`/auth/oidc/*`, `/auth/saml/*`): returns `429` + `Retry-After` when limited

If the API is deployed behind a trusted reverse proxy / load balancer, set:

- `TRUST_PROXY=true`

so Fastify will honor forwarding headers (e.g. `X-Forwarded-For`). Do **not** enable this
unless the proxy strips spoofed headers.

### Port overrides

If you already have something running on these ports, you can override the published ports:

```bash
API_PORT=3001 SYNC_WS_PORT=1235 POSTGRES_PORT=5433 docker compose up --build
```

Defaults are:

- `API_PORT=3000`
- `SYNC_WS_PORT=1234`
- `POSTGRES_PORT=5432`

The examples below also read `SYNC_WS_PORT` (default `1234`) so they keep working if you override ports.
The curl examples use `API_BASE` derived from `API_PORT` (default `3000`).

## Token configuration (API ↔ sync-server)

The API issues short-lived JWT sync tokens via `POST /docs/:docId/sync-token`.

- API signing secret: `SYNC_TOKEN_SECRET`
- Token lifetime: `SYNC_TOKEN_TTL_SECONDS` (in seconds)
- Sync-server verification secret: `SYNC_SERVER_JWT_SECRET` (must match `SYNC_TOKEN_SECRET`)
- JWT audience: `SYNC_SERVER_JWT_AUDIENCE` (must match the token `aud`, default: `formula-sync`)

In `docker-compose.yml`, the sync server is configured to reuse the API secret by default, so you can override both by setting `SYNC_TOKEN_SECRET`:

```bash
SYNC_TOKEN_SECRET=my-local-sync-secret docker compose up --build
```

### Sync token introspection (revocation + permission revalidation)

When the sync server verifies JWTs locally (`SYNC_SERVER_AUTH_MODE=jwt-hs256`), token claims can become stale (e.g. if
sessions are revoked or document permissions change).

To have the sync server revalidate tokens against server-side state, use the API internal introspection endpoint:

- API: `POST /internal/sync/introspect` (requires `INTERNAL_ADMIN_TOKEN` via `x-internal-admin-token`)

Sync server options:

- **Auth mode** (recommended for production deployments): `SYNC_SERVER_AUTH_MODE=introspect`
  - `SYNC_SERVER_INTROSPECT_URL=<api base url>`
  - `SYNC_SERVER_INTROSPECT_TOKEN=<same as INTERNAL_ADMIN_TOKEN>`
- **Optional revalidation** (keep JWT auth, but revalidate during upgrade): set
  - `SYNC_SERVER_INTROSPECTION_URL=<api base url>`
  - `SYNC_SERVER_INTROSPECTION_TOKEN=<same as INTERNAL_ADMIN_TOKEN>`

In both cases the sync server forwards `clientIp` and `userAgent` so the API can enforce org IP allowlists.

## Secret store configuration

The API includes a small database-backed encrypted secret store (`secrets` table). It is used for things like:

- OIDC client secrets
- SIEM auth configuration (see [`docs/siem.md`](./siem.md))

### Key configuration

The secret store supports multiple keys for rotation without downtime.

Configuration options (highest priority first):

- `SECRET_STORE_KEYS_JSON`: JSON encoded keyring.
- `SECRET_STORE_KEYS` (recommended): comma-separated list of `<keyId>:<base64>` entries (the last entry is current).
- legacy `SECRET_STORE_KEY`: a single secret which is hashed with SHA-256 to derive the AES-256 key.

Examples:

`SECRET_STORE_KEYS`:

```bash
SECRET_STORE_KEYS="k2025-12:<base64(32-byte key)>,k2026-01:<base64(32-byte key)>"
```

`SECRET_STORE_KEYS_JSON`:

```json
{
  "currentKeyId": "k2026-01",
  "keys": {
    "k2025-12": "<base64(32-byte key)>",
    "k2026-01": "<base64(32-byte key)>"
  }
}
```

Note: `SECRET_STORE_KEYS_JSON` also accepts `{ current, keys }`, direct maps like `{ "k1": "...", "k2": "..." }`
(current defaults to the last entry), or arrays like `[ { "id": "k1", "key": "..." }, ... ]`.

To generate a new 32-byte key value:

```bash
node -e 'console.log(require("crypto").randomBytes(32).toString("base64"))'
```

Key ids are application-defined (e.g. a date string), but must be non-empty and must not include `:` (the delimiter used in the on-disk encoding).

If neither `SECRET_STORE_KEYS_JSON` nor `SECRET_STORE_KEYS` is set, the API falls back to legacy single-key mode using `SECRET_STORE_KEY` (compatible with older deployments).

Note: legacy deployments derive the actual AES-256 key as `sha256(SECRET_STORE_KEY)`. When migrating from `SECRET_STORE_KEY` to a multi-key ring (`SECRET_STORE_KEYS` / `SECRET_STORE_KEYS_JSON`), include the derived key as one of the entries so existing secrets remain decryptable.

### Rotation tooling

After adding a new key (and making it current), re-encrypt existing rows with:

```bash
pnpm -C services/api secrets:rotate
```

This scans the `secrets` table and re-encrypts any secrets not using the current key id into the latest format. It is safe to run while the API is online.

Optional env vars:

- `PREFIX="oidc:<orgId>:"` to scope rotation to a literal secret-name prefix.
- `BATCH_SIZE=250` to control pagination.

## Persistence (local docker compose)

The docker compose stack configures the sync server with:

- `SYNC_SERVER_PERSISTENCE_BACKEND=file`
- `SYNC_SERVER_DATA_DIR=/data`
- persistence stored in the `sync_server_data` named volume (mounted at `/data`)

To wipe local sync persistence, run `docker compose down -v`.

## API overview

### Authentication

- `POST /auth/register` → creates a user + a personal org + a session cookie
- `POST /auth/login` → session cookie
- `POST /auth/logout`
- `GET /me` → current user + org memberships

### Documents + RBAC

- `POST /docs` (requires auth) → create a document in an org
- `POST /docs/:docId/invite` → invite an existing user by email (document-level role)
- `POST /docs/:docId/sync-token` → issue a short-lived collaboration token for the sync server

Roles: `owner | admin | editor | commenter | viewer`

### Range restrictions (optional)

The API can optionally include `rangeRestrictions` in sync JWTs (and/or return a `rangeRestrictions`
array in sync token introspection responses when using `SYNC_SERVER_AUTH_MODE=introspect`). When enabled,
the sync server enforces these restrictions **server-side** so a malicious client cannot bypass UI checks
by sending crafted Yjs updates.

Enable enforcement:

- `SYNC_SERVER_ENFORCE_RANGE_RESTRICTIONS=1` (default: `true` in `NODE_ENV=production`, `false` otherwise)

If a client attempts to write to a protected cell, the sync server closes the connection with
policy violation (`1008`) and logs `permission_violation`.

Opaque token auth (`SYNC_SERVER_AUTH_TOKEN`) cannot carry range restrictions (JWT and
`SYNC_SERVER_AUTH_MODE=introspect` can).

The claim shape matches `packages/collab/permissions.normalizeRestriction`:

```json
{
  "rangeRestrictions": [
    {
      "range": { "sheetId": "Sheet1", "startRow": 0, "endRow": 0, "startCol": 0, "endCol": 0 },
      "editAllowlist": ["user-id"],
      "readAllowlist": ["user-id"]
    }
  ]
}
```

### Audit logging

- `GET /orgs/:orgId/audit` (org admin) → query audit events
- `GET /orgs/:orgId/audit/export` (org admin) → export audit events (`format=json|cef|leef`, default: `json`)

### SIEM integration

For SIEM configuration APIs and details on the background export worker, see [`docs/siem.md`](./siem.md).

### Retention / residency scaffolding

Org settings include:
- `dataResidencyRegion`
- retention windows for audit log + document versions

The API runs a periodic retention sweep (configurable via `RETENTION_SWEEP_INTERVAL_MS`).

## Example: register → create doc → invite → sync token → connect

If you overrode `API_PORT`, set `API_BASE` so the commands below keep working:

```bash
API_PORT=${API_PORT:-3000}
API_BASE="http://localhost:${API_PORT}"
```

1) Register two users:

```bash
curl -i "$API_BASE/auth/register" \
  -H 'content-type: application/json' \
  -d '{"email":"alice@example.com","password":"password1234","name":"Alice","orgName":"Acme"}'

curl -i "$API_BASE/auth/register" \
  -H 'content-type: application/json' \
  -d '{"email":"bob@example.com","password":"password1234","name":"Bob"}'
```

2) Create a document (use Alice's `Set-Cookie` session):

```bash
curl -i "$API_BASE/docs" \
  -H 'content-type: application/json' \
  -H 'cookie: formula_session=...' \
  -d '{"orgId":"<alice-org-id>","title":"Q1 Plan"}'
```

3) Invite Bob to the document:

```bash
curl -i "$API_BASE/docs/<doc-id>/invite" \
  -H 'content-type: application/json' \
  -H 'cookie: formula_session=...' \
  -d '{"email":"bob@example.com","role":"editor"}'
```

4) Bob requests a sync token (use Bob's session cookie):

```bash
curl -s -X POST "$API_BASE/docs/<doc-id>/sync-token" \
  -H 'content-type: application/json' \
  -H 'cookie: formula_session=...' \
  -d '{}'
```

5) Connect to the sync server with the token using a Yjs client (`y-websocket`):

```bash
TOKEN=... DOC_ID=... pnpm -C services/sync-server exec node --input-type=module - <<'NODE'
import * as Y from "yjs";
import { WebsocketProvider } from "y-websocket";

const token = process.env.TOKEN;
const docId = process.env.DOC_ID;
if (!token || !docId) throw new Error("Missing TOKEN or DOC_ID env var");

const wsPort = process.env.SYNC_WS_PORT ?? "1234";
const wsUrl = `ws://localhost:${wsPort}`;

// Prefer Node's built-in WebSocket (Node 20+). Fall back to the `ws` package if needed.
const WebSocketPolyfill = globalThis.WebSocket ?? (await import("ws")).default;

const ydoc = new Y.Doc();

const provider = new WebsocketProvider(wsUrl, docId, ydoc, {
  WebSocketPolyfill,
  disableBc: true,
  params: { token }
});

provider.on("status", (event) => console.log("ws status:", event.status));
provider.on("sync", (isSynced) => console.log("synced:", isSynced));

// Send a trivial update once we're connected.
const onSync = (isSynced) => {
  if (!isSynced) return;
  ydoc.getText("t").insert(0, "hello");
  provider.off("sync", onSync);
};
provider.on("sync", onSync);

setTimeout(() => {
  provider.destroy();
  ydoc.destroy();
}, 2_000);
NODE
```

Note: the Yjs example uses `pnpm ... exec` so the required JS dependencies (`yjs`, `y-websocket`, `ws`) are available. If you haven’t installed dependencies yet, run `pnpm install` first (or skip to the minimal WebSocket auth check below).

If you just want to verify the token authorizes the WebSocket upgrade (without speaking the `y-websocket` protocol), you can also do a minimal connect/disconnect:

```bash
TOKEN=... DOC_ID=... node - <<'NODE'
const token = process.env.TOKEN
const docId = process.env.DOC_ID
if (!token || !docId) throw new Error('Missing TOKEN or DOC_ID env var')

const wsPort = process.env.SYNC_WS_PORT ?? '1234'

if (typeof WebSocket !== 'function') {
  throw new Error('WebSocket is not available (need Node 20+)')
}

const ws = new WebSocket(`ws://localhost:${wsPort}/${docId}?token=${encodeURIComponent(token)}`)
ws.addEventListener('open', () => {
  console.log('connected')
  ws.close()
})
ws.addEventListener('close', (ev) => console.log('closed', ev.code))
ws.addEventListener('error', (ev) => console.error('error', ev))
NODE
```
