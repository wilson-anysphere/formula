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

## Persistence (local docker-compose)

The docker-compose stack configures the sync server with:

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

### Audit logging

- `GET /orgs/:orgId/audit` (org admin) → query audit events
- `GET /orgs/:orgId/audit/export` (org admin) → NDJSON export

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
