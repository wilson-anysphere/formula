# Formula — AI-Native Spreadsheet (Excel-Compatible)

Formula is a next-generation spreadsheet with a **desktop-first** product strategy (Tauri) and an **optional web target** used to keep the core engine/UI portable.

## Python scripting

Formula supports modern scripting with a stable Python API (`import formula`) designed to mirror the macro compatibility spec in [`docs/08-macro-compatibility.md`](./docs/08-macro-compatibility.md).

- Python package (in-repo): `python/formula_api/`
- Runtimes / bridges (JS): `packages/python-runtime/`
  - Native Python subprocess (desktop/Node)
  - Pyodide-in-Worker (web/webview)

## Platform strategy (high level)

- **Primary target:** Tauri desktop app (Windows/macOS/Linux).
- **Secondary target:** Web build for development, demos, and long-term optional deployment.
- **Core principle:** The calculation engine runs behind a **WASM boundary** and is executed off the UI thread (Worker). Platform-specific integrations live in thin host adapters.

For details, see:
- [`docs/adr/ADR-0001-platform-target.md`](./docs/adr/ADR-0001-platform-target.md)
- [`docs/adr/ADR-0002-engine-execution-model.md`](./docs/adr/ADR-0002-engine-execution-model.md)

## Development

### Built-in function catalog (tab completion + hints)

The set of built-in Excel-like functions is owned by the Rust formula engine (`crates/formula-engine`)
via an inventory-backed registry (`FunctionSpec`). JavaScript/TypeScript features like tab-completion
and formula-bar signature hints consume a **generated** catalog committed into the repo:

- `shared/functionCatalog.json` (canonical artifact)
- `shared/functionCatalog.mjs` (ESM wrapper for runtime import compatibility)
- `shared/functionCatalog.mjs.d.ts` (TypeScript typings for the wrapper)

To regenerate after adding/removing Rust functions (requires a Rust toolchain):

```bash
pnpm generate:function-catalog
```

### Rust/WASM engine (formula-wasm)

The JS `@formula/engine` package loads the Rust engine via `wasm-bindgen` artifacts generated from `crates/formula-wasm`.

Build the WASM artifacts (requires `wasm-pack` on your `PATH`):

```bash
pnpm build:wasm
```

The web + desktop Vite entrypoints run this automatically via `predev`/`prebuild` so `createEngineClient()` can load the engine without extra manual steps.

Smoke-check that the generated wrapper (`packages/engine/pkg/formula_wasm.js`) exists:

```bash
pnpm smoke:wasm
```

### Web (preview target)

```bash
pnpm install
pnpm dev:web
```

Build:

```bash
pnpm build:web
```

The web target renders the shared Canvas grid (`@formula/grid`) with a mock data provider and brings up a Worker-based engine client (`@formula/engine`). The Worker boundary is where the Rust/WASM engine will execute as it is implemented.

### Collaboration sync server (Yjs)

This repo includes a production-ready **Yjs sync server** (WebSocket, `y-websocket` protocol) with **LevelDB persistence**, auth, basic rate limiting, and health checks.

Run locally:

```bash
pnpm dev:sync
```

Defaults:

- WebSocket: `ws://127.0.0.1:1234/<documentId>?token=<token>`
- Health: `http://127.0.0.1:1234/healthz`

Persistence is stored under `SYNC_SERVER_DATA_DIR` (defaults to `./.sync-server-data/`).

You can switch persistence backends:

- `SYNC_SERVER_PERSISTENCE_BACKEND=leveldb` (default)
- `SYNC_SERVER_PERSISTENCE_BACKEND=file` (portable fallback)

#### Retention / purge (LevelDB only)

When using `SYNC_SERVER_PERSISTENCE_BACKEND=leveldb` with `y-leveldb` installed, the sync server stores per-document metadata:

- `lastSeenMs`: updated when a document is loaded and on subsequent updates (throttled).

You can purge old documents from the LevelDB store without external bookkeeping:

- Enable internal admin endpoints:
  - `SYNC_SERVER_INTERNAL_ADMIN_TOKEN=<token>`
- Configure retention TTL (required for purging):
  - `SYNC_SERVER_RETENTION_TTL_MS=<milliseconds>`
- (Optional) Run periodic sweeps in the background:
  - `SYNC_SERVER_RETENTION_SWEEP_INTERVAL_MS=<milliseconds>`

Trigger a sweep manually:

```bash
curl -X POST \
  -H "x-internal-admin-token: $SYNC_SERVER_INTERNAL_ADMIN_TOKEN" \
  http://127.0.0.1:1234/internal/retention/sweep
```

#### Encryption at rest (file persistence)

The `file` persistence backend supports **encryption at rest** (AES-256-GCM) for persisted `.yjs` documents.

Enable:

- `SYNC_SERVER_PERSISTENCE_BACKEND=file`
- `SYNC_SERVER_PERSISTENCE_ENCRYPTION=keyring`
- Provide key material via **one** of:
  - `SYNC_SERVER_ENCRYPTION_KEYRING_JSON` (KeyRing JSON string), or
  - `SYNC_SERVER_ENCRYPTION_KEYRING_PATH` (path to a JSON file containing KeyRing JSON)

When encryption is enabled, existing legacy plaintext `.yjs` files in `SYNC_SERVER_DATA_DIR` are migrated to the encrypted, append-only format **on startup** (atomic per file).

Key rotation is operator-managed by replacing the KeyRing JSON (bumping `currentVersion` and adding a new key while keeping old key versions available for decryption).

#### Auth (dev default)

If no auth env vars are provided, the server starts with a **development token**:

- token: `dev-token`

For production, set **one** of:

- `SYNC_SERVER_AUTH_TOKEN` (opaque shared token), or
- `SYNC_SERVER_JWT_SECRET` (HMAC JWT secret; HS256)

If using JWTs, tokens are verified with:

- algorithm: `HS256`
- audience: `SYNC_SERVER_JWT_AUDIENCE` (defaults to `formula-sync`) — tokens must include a matching `aud`

JWT claims:

- **Preferred:** `docId: string` — must match the websocket document path exactly (e.g. connecting to `ws://.../my-document-id` requires `docId: "my-document-id"`).
- `role: owner|admin|editor|commenter|viewer` (optional; defaults to `editor`)
  - `viewer` and `commenter` are enforced as **read-only** at the Yjs protocol layer.
- `orgId: string` (optional; included in API-issued tokens)

Legacy (backwards-compatible) allowlisting is still supported if `docId` is not present:

- `doc: string`, or
- `docs: string[]` (e.g. `["my-document-id"]` or `["*"]`)

Token minting:

- The Formula API (`services/api`) can mint compatible sync tokens via `POST /docs/:docId/sync-token`.
  - These tokens include `docId`, `orgId`, `role`, and `aud=formula-sync`.
  - To use them with `services/sync-server`, configure the secrets to match (`SYNC_TOKEN_SECRET` in the API must equal `SYNC_SERVER_JWT_SECRET` in the sync server).

Awareness hardening:

- The server sanitizes awareness updates to prevent presence spoofing; presence `id` is forced to the JWT `sub`.

#### Connect from a Yjs client

```ts
import * as Y from "yjs";
import { WebsocketProvider } from "y-websocket";

const doc = new Y.Doc();
const provider = new WebsocketProvider(
  "ws://127.0.0.1:1234",
  "my-document-id",
  doc,
  { params: { token: "dev-token" } }
);
```

The service lives in:

- `services/sync-server/`

## License

Licensed under the Apache License 2.0. See [`LICENSE`](./LICENSE) and [`NOTICE`](./NOTICE).
