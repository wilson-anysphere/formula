# Workstream E: Collaboration

> **⛔ STOP. READ [`AGENTS.md`](../AGENTS.md) FIRST. FOLLOW IT COMPLETELY. THIS IS NOT OPTIONAL. ⛔**
>
> This document is supplementary to AGENTS.md. All rules, constraints, and guidelines in AGENTS.md apply to you at all times. Memory limits, build commands, design philosophy—everything.

---

## Mission

Build real-time collaboration that doesn't break at scale. Seamless, conflict-free, offline-first using **CRDTs (Conflict-free Replicated Data Types)** via Yjs.

**The goal:** Google Docs-level collaboration for spreadsheets, with Git-like version control.

---

## Scope

### Your Code

| Location | Purpose |
|----------|---------|
| `packages/collab` | CRDT implementation, Yjs integration |
| `packages/versioning` | Version history, checkpoints, diff |
| `services/sync-server` | WebSocket sync server (y-websocket) |
| `apps/desktop/src/collab/` | Collaboration UI integration |
| `apps/desktop/src/versioning/` | Version history UI |
| `apps/desktop/src/comments/` | Cell-level commenting |

### Your Documentation

- **Primary:** [`docs/06-collaboration.md`](../docs/06-collaboration.md) — implementation-backed wiring for session + binder + presence + versioning/branching

---

## Key Requirements

### CRDT Data Model (Yjs)

```typescript
interface SpreadsheetDoc {
  doc: Y.Doc;
  // Sheet list. Each entry is a Y.Map containing (at minimum) `{ id, name }` plus
  // optional shared per-sheet state like:
  // - `view` (frozen panes + row/col sizes)
  // - `visibility`, `tabColor` (sheet metadata)
  // - formatting defaults (`defaultFormat`, `rowFormats`, `colFormats`, `formatRunsByCol`)
  sheets: Y.Array<Y.Map<any>>;
  // Cell map keyed by canonical cell keys `${sheetId}:${row}:${col}` (0-based row/col).
  cells: Y.Map<Y.Map<any>>;
  metadata: Y.Map<any>;              // Workbook metadata
  namedRanges: Y.Map<any>;           // Named range definitions
  // Optional comments root. Canonical schema uses a Map keyed by comment id, but
  // legacy docs may store an Array.
  comments?: Y.Map<any> | Y.Array<Y.Map<any>>;
}
```

Important nuance (formula clears): for deterministic conflict detection (when
`FormulaConflictMonitor` is enabled), represent clears with an explicit
`formula = null` marker rather than deleting the `formula` key. Yjs map deletes do not create
Items; a `null` marker preserves causal history used by conflict detection.

Related: “empty” cells may still exist as marker-only `Y.Map`s in Yjs (e.g. to carry the
`formula = null` clear marker for deterministic delete-vs-overwrite detection).

See [`docs/06-collaboration.md`](../docs/06-collaboration.md) for:

- desktop binder wiring (`packages/collab/binder/index.js`, `bindYjsToDocumentController`)
- sheet `view` state storage (`sheets[i].get("view")`)
- presence rendering (`apps/desktop/src/grid/presence-renderer/`)
- comments (`@formula/collab-comments`) and conflict monitors (`@formula/collab-conflicts`)
- collaborative versioning (`@formula/collab-versioning`) and branching (`packages/collab/branching/index.js`)
- ADRs (shared sheet view state + undo semantics): [`docs/adr/ADR-0004-collab-sheet-view-and-undo.md`](../docs/adr/ADR-0004-collab-sheet-view-and-undo.md)

### Sync Server Features

- **WebSocket sync** (y-websocket protocol)
- **LevelDB persistence** with optional encryption at rest
- **Authentication:** JWT or shared token
- **Rate limiting** and health checks
- **Presence** with awareness hardening (anti-spoofing)

### Presence System

- Show who's editing what (cursor positions, selections)
- User avatars/colors
- Real-time cursor movement
- Awareness protocol sanitization

### Version History

- **Named checkpoints** ("Q3 Budget Approved")
- **Cell-by-cell diff** with color coding
- **Formula diff** showing specific changes
- **Branch-and-merge** for scenario analysis
- **Side-by-side comparison** of conflicting cells

### Conflict Resolution

CRDTs handle conflicts automatically, but we need:
- Last-write-wins for simple cell edits
- Merge strategies for structural changes (row/column insert)
- UI for manual resolution when needed

---

## Sync Server

### Run Locally

```bash
pnpm dev:sync
```

- WebSocket: `ws://127.0.0.1:1234/<documentId>?token=<token>`
- Health: `http://127.0.0.1:1234/healthz`
- Dev token: `dev-token`

### Configuration

| Env Var | Purpose |
|---------|---------|
| `SYNC_SERVER_AUTH_TOKEN` | Shared auth token |
| `SYNC_SERVER_JWT_SECRET` | JWT secret (HS256) |
| `SYNC_SERVER_PERSISTENCE_BACKEND` | `leveldb` or `file` |
| `SYNC_SERVER_PERSISTENCE_ENCRYPTION` | `keyring` for encryption at rest |
| `SYNC_SERVER_MAX_MESSAGE_BYTES` | Max websocket message size (defaults to 2MB; see close code `1009`) |

Note: large in-doc history payloads (e.g. version snapshots stored under `versions*` and branch commits stored under `branching:*`) can exceed this limit if written as a single Yjs update. `@formula/collab-versioning` now defaults to streaming snapshot writes (chunked across multiple Yjs transactions/updates) to avoid 1009 disconnects for large workbooks. See [`docs/06-collaboration.md`](../docs/06-collaboration.md) for tuning via `yjsStoreOptions`.

### Client Connection

```typescript
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

---

## Offline Support

1. All edits applied locally first (optimistic)
2. Queue operations when offline
3. Automatic sync on reconnection
4. CRDT merge handles conflicts

---

## Performance Targets

| Metric | Target |
|--------|--------|
| Sync latency | <100ms for local edits to appear remotely |
| Offline queue | Support 1000+ pending operations |
| Presence update | <50ms |

---

## Build & Run

```bash
# Install dependencies
pnpm install

# Run sync server
pnpm dev:sync

# Run tests
pnpm -C services/sync-server test

# Build Docker image
docker build -f services/sync-server/Dockerfile -t formula-sync-server .
```

---

## Coordination Points

- **UI Team:** Presence indicators, cursor display, version history UI
- **Core Engine Team:** Cell operations map to CRDT operations
- **File I/O Team:** Serialization for persistence
- **AI Team:** AI actions in collaborative context

---

## Security

- **Auth context:** `docId`, `sub`/`userId` (user id), `role` (owner/admin/editor/commenter/viewer), optional `rangeRestrictions` (from JWT claims or token introspection responses)
- **Read-only enforcement:** viewer role enforced at the sync-server (drops writes); commenter role is comment-only (comments allowed, workbook edits rejected). Desktop mirrors this in the UX (binder/UI disable edits; comments gated via `session.canComment()`).
- **Awareness sanitization:** presence `id` forced to JWT `sub` (desktop should set its local presence id to `sub` so identities are stable)
- **Encryption at rest:** AES-256-GCM for persisted documents

---

## Reference

- Yjs documentation: https://docs.yjs.dev/
- y-websocket: https://github.com/yjs/y-websocket
- CRDT papers: https://crdt.tech/
