# Formula — AI-Native Spreadsheet (Excel-Compatible)

> **A Cursor product.** All AI features are powered by Cursor's backend—no local models, no API keys, no provider configuration.

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
- [`docs/adr/ADR-0003-engine-protocol-parity.md`](./docs/adr/ADR-0003-engine-protocol-parity.md)

## Development

### Data model / DAX engine (Power Pivot)

The in-repo DAX engine lives in `crates/formula-dax`. Contributor docs:

- [`docs/21-dax-engine.md`](./docs/21-dax-engine.md) — supported DAX syntax/functions, relationships + filter propagation, calculated columns, and the pivot API.

### Built-in function catalog (tab completion + hints)

The set of built-in Excel-like functions is owned by the Rust formula engine (`crates/formula-engine`)
via an inventory-backed registry (`FunctionSpec`). JavaScript/TypeScript features like tab-completion
and formula-bar signature hints consume a **generated** catalog committed into the repo:

- `shared/functionCatalog.json` (canonical artifact)
- `shared/functionCatalog.mjs` (ESM wrapper for runtime import compatibility)
- `shared/functionCatalog.d.mts` (TypeScript typings for the wrapper)

Each entry includes (at minimum): `name`, `min_args`, `max_args`, `volatility`, `return_type`, and a
best-effort `arg_types` array derived from the Rust `FunctionSpec` metadata.

To regenerate after adding/removing Rust functions (requires a Rust toolchain; pinned via `rust-toolchain.toml`):

```bash
pnpm generate:function-catalog
```

### AI tab-completion latency benchmark

To guard against performance regressions in the JavaScript tab-completion engine (`TabCompletionEngine`),
you can run a lightweight micro-benchmark locally:

```bash
pnpm bench:tab-completion
```

### Rust/WASM engine (formula-wasm)

The JS `@formula/engine` package loads the Rust engine via `wasm-bindgen` artifacts generated from `crates/formula-wasm`.

Build the WASM artifacts (requires `wasm-pack` on your `PATH`):

```bash
pnpm build:wasm
```

This command:

- builds deterministic wasm-pack output into `packages/engine/pkg/`
- copies runtime assets into `apps/web/public/engine/` and `apps/desktop/public/engine/` so the worker can import them from a stable URL: `/engine/formula_wasm.js`

The web + desktop Vite entrypoints run this automatically via `predev`/`prebuild` so `createEngineClient()` can load the engine without extra manual steps.

#### Formula editor tooling (lexing + partial parse)

`@formula/engine` also exposes **workbook-independent** editor tooling helpers that run in the same Worker-backed WASM module:

- `engine.lexFormula(formula, options?)` → token DTOs for syntax highlighting
- `engine.parseFormulaPartial(formula, cursor?, options?)` → best-effort partial parse + function-call context for autocomplete/signature help

`options` is a small JS-friendly object (`FormulaParseOptions`):

- `localeId?: string` (e.g. `"en-US"`, `"de-DE"`) — affects argument separators, decimal separators, localized function names, etc.
- `referenceStyle?: "A1" | "R1C1"`

Note: spans and cursor positions are expressed as **UTF-16 code unit offsets** (matching JS string indexing).

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

For deployment details (auth modes, limits & hardening, metrics, etc.), see [`services/sync-server/README.md`](./services/sync-server/README.md).

Run locally:

```bash
pnpm dev:sync
```

Defaults:

- WebSocket: `ws://127.0.0.1:1234/<documentId>?token=<token>`
- Health: `http://127.0.0.1:1234/healthz`
- Ready: `http://127.0.0.1:1234/readyz`

Persistence is stored under `SYNC_SERVER_DATA_DIR` (defaults to `./.sync-server-data/`).

The sync server is **single-writer per data directory**. On startup it creates an exclusive lock file:

- `${SYNC_SERVER_DATA_DIR}/.sync-server.lock`

If the lock file already exists, the server will refuse to start to avoid on-disk corruption. If the server crashes and leaves a stale lock file behind, delete it manually after confirming no other `sync-server` process is using that directory. Note: stale-lock cleanup is **host-local**; if `SYNC_SERVER_DATA_DIR` is shared across machines, a lock from another host will not be auto-removed—stop the other instance and delete the lock file manually. (You can disable locking with `SYNC_SERVER_DISABLE_DATA_DIR_LOCK=true`, but this is not recommended outside of testing.)

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

#### LevelDB docName hashing (metadata privacy)

By default, `y-leveldb` includes the raw `documentId` in **LevelDB keys** (e.g. `["v1", docName, "update", clock]`), which means document names appear in plaintext in SSTables/LOG.

To reduce metadata leakage, you can enable doc-name hashing:

- `SYNC_SERVER_LEVELDB_DOCNAME_HASHING=1` (default `0` for backcompat)

When enabled, the server derives a persistent name for LevelDB keys:

```
persistedName = sha256(documentId) // hex
```

Migration / backcompat:

- On first load of a document, the server will **also** look for legacy (unhashed) keys.
- If legacy keys are present, the server will migrate the document (and any per-doc metas) into the hashed namespace on the next flush (when the last client disconnects) and then delete the legacy keys.

#### Encryption at rest (file persistence)

The `file` persistence backend supports **encryption at rest** (AES-256-GCM) for persisted `.yjs` documents.

Enable:

- `SYNC_SERVER_PERSISTENCE_BACKEND=file`
- Enable keyring encryption via **either**:
  - `SYNC_SERVER_PERSISTENCE_ENCRYPTION=keyring`, or
  - `SYNC_SERVER_PERSISTENCE_ENCRYPTION_KEY_B64=<base64>` (shorthand; implies `keyring` mode)
- Provide key material via **one** of:
  - `SYNC_SERVER_ENCRYPTION_KEYRING_JSON` (KeyRing JSON string), or
  - `SYNC_SERVER_ENCRYPTION_KEYRING_PATH` (path to a JSON file containing KeyRing JSON)
  - `SYNC_SERVER_PERSISTENCE_ENCRYPTION_KEY_B64` (base64-encoded **32-byte** key; convenience option that creates a single-version KeyRing)
  - (Optional shorthand) `SYNC_SERVER_PERSISTENCE_ENCRYPTION_KEY_B64` (base64 32-byte key; creates a single-version keyring)

The server will refuse to start if encryption is enabled but key material is missing (always in production; also in dev/test when `SYNC_SERVER_PERSISTENCE_ENCRYPTION=keyring` is set).

When encryption is enabled, existing legacy plaintext `.yjs` files in `SYNC_SERVER_DATA_DIR` are migrated to the encrypted, append-only format **on startup** (atomic per file).

Manage KeyRing material (generate / rotate / validate):

```bash
# Generate a new keyring (write to a secret file, or inject via env)
pnpm -C services/sync-server -s keyring:generate --out keyring.json

# Validate and inspect a keyring
pnpm -C services/sync-server -s keyring:validate --in keyring.json

# Rotate (adds a new key version; keeps old keys)
pnpm -C services/sync-server -s keyring:rotate --in keyring.json --out keyring.json
```

The keyring JSON contains **secret key material**. Store it in your secret manager or lock down file permissions (e.g. `chmod 600 keyring.json`).

In a built deployment you can run the compiled entrypoint directly:

```bash
node services/sync-server/dist/keyring-cli.js generate --out keyring.json
```

Key rotation is operator-managed by replacing the KeyRing JSON (bumping `currentVersion` and adding a new key while keeping old key versions available for decryption).

#### Encryption at rest (LevelDB persistence)

When using `SYNC_SERVER_PERSISTENCE_BACKEND=leveldb` (default) with `y-leveldb` installed, the sync server can encrypt all **LevelDB values** (updates, state vectors, metadata) at rest using **AES-256-GCM**.

Enable:

- `SYNC_SERVER_PERSISTENCE_BACKEND=leveldb`
- Enable keyring encryption via **either**:
  - `SYNC_SERVER_PERSISTENCE_ENCRYPTION=keyring`, or
  - `SYNC_SERVER_PERSISTENCE_ENCRYPTION_KEY_B64=<base64>` (shorthand; implies `keyring` mode)
- Provide key material via **one** of:
  - `SYNC_SERVER_ENCRYPTION_KEYRING_JSON`, or
  - `SYNC_SERVER_ENCRYPTION_KEYRING_PATH`
  - `SYNC_SERVER_PERSISTENCE_ENCRYPTION_KEY_B64` (base64-encoded **32-byte** key; convenience option that creates a single-version KeyRing)
  - (Optional shorthand) `SYNC_SERVER_PERSISTENCE_ENCRYPTION_KEY_B64` (base64 32-byte key; creates a single-version keyring)

The server will refuse to start if encryption is enabled but key material is missing.

Optional migration strictness:

- `SYNC_SERVER_PERSISTENCE_ENCRYPTION_STRICT=1|0`
  - default: `1` in production, `0` in dev/test
  - strict (`1`): rejects legacy plaintext values (no `FMLLDB01` header)
  - non-strict (`0`): allows reading legacy plaintext values for migration (new writes are always encrypted)

Migration recommendation (LevelDB):

- Safest: start with a fresh `SYNC_SERVER_DATA_DIR`.
- Otherwise: run with `SYNC_SERVER_PERSISTENCE_ENCRYPTION_STRICT=0` until legacy values are rewritten (e.g. via `flushDocument` on connect/last disconnect), then switch back to strict mode.

Note: LevelDB **keys** (including `docName`) remain plaintext unless doc-name hashing is enabled.

#### Internal admin API (purge persisted docs)

The sync server exposes a small internal HTTP API intended for retention and
operational workflows (purging documents, etc). These endpoints are **disabled
by default**.

Enable by setting:

- `SYNC_SERVER_INTERNAL_ADMIN_TOKEN`

All internal endpoints require:

- header: `x-internal-admin-token: <token>`

Purge a persisted Yjs document (disconnects active clients for that document):

- `DELETE /internal/docs/<docName>` → `{ ok: true }`

`<docName>` is the same document id used in the WebSocket URL and may contain
slashes (URL-encode as needed).

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

#### Docker

Build from the repo root:

```bash
docker build -f services/sync-server/Dockerfile -t formula-sync-server .
```

Force native modules (e.g. `leveldown`) to compile from source (useful when validating the node-gyp toolchain / no-prebuild scenarios):

```bash
docker build --build-arg npm_config_build_from_source=true -f services/sync-server/Dockerfile -t formula-sync-server .
```

Run (auth is required in production):

```bash
docker run --rm -p 1234:1234 \
  -e SYNC_SERVER_AUTH_TOKEN=dev-token \
  formula-sync-server
```

Persist data (recommended for `leveldb`):

```bash
docker run --rm -p 1234:1234 \
  -e SYNC_SERVER_AUTH_TOKEN=dev-token \
  -v sync-server-data:/app/services/sync-server/.sync-server-data \
  formula-sync-server
```

## License

Licensed under the Apache License 2.0. See [`LICENSE`](./LICENSE) and [`NOTICE`](./NOTICE).
