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

#### Auth (dev default)

If no auth env vars are provided, the server starts with a **development token**:

- token: `dev-token`

For production, set **one** of:

- `SYNC_SERVER_AUTH_TOKEN` (opaque shared token), or
- `SYNC_SERVER_JWT_SECRET` (HMAC JWT secret; HS256)

If using JWTs, include either:

- `docs: string[]` (e.g. `["my-document-id"]` or `["*"]`), or
- `doc: string`

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
