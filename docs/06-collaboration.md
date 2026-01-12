# Collaboration (Yjs)

This document is the implementation-backed reference for how Formula wires together:

- **Sync + offline persistence** (`@formula/collab-session`)
- **Desktop workbook binding** (`packages/collab/binder/index.js`)
- **Presence** (`@formula/collab-presence` + desktop `PresenceRenderer`)
- **Version history** (`@formula/collab-versioning`)
- **Branching/merging** (`packages/collab/branching/index.js` + `packages/versioning/branches`)

If you are editing collaboration code, start here and keep this doc in sync with the implementation.

## Design decisions

- [ADR-0004: Collaboration semantics for sheet view state and undo](./adr/ADR-0004-collab-sheet-view-and-undo.md)

---

## Key modules (source of truth)

- Session orchestration: [`packages/collab/session/src/index.ts`](../packages/collab/session/src/index.ts) (`createCollabSession`)
- Workbook roots/schema helpers: [`packages/collab/workbook/src/index.ts`](../packages/collab/workbook/src/index.ts) (`getWorkbookRoots`, `ensureWorkbookSchema`)
- Cell key helpers: [`packages/collab/session/src/cell-key.js`](../packages/collab/session/src/cell-key.js) (`makeCellKey`, `parseCellKey`, `normalizeCellKey`)
- Desktop binder: [`packages/collab/binder/index.js`](../packages/collab/binder/index.js) (`bindYjsToDocumentController`)
- Presence (Awareness wrapper): [`packages/collab/presence/src/presenceManager.js`](../packages/collab/presence/src/presenceManager.js) (`PresenceManager`)
- Desktop presence renderer: [`apps/desktop/src/grid/presence-renderer/`](../apps/desktop/src/grid/presence-renderer/) (`PresenceRenderer`)
- Comments (Yjs `comments` root helpers): [`packages/collab/comments/src/manager.ts`](../packages/collab/comments/src/manager.ts) (`CommentManager`, `createCommentManagerForSession`, `migrateCommentsArrayToMap`)
- Conflict monitors: [`packages/collab/conflicts/index.js`](../packages/collab/conflicts/index.js) (`FormulaConflictMonitor`, `CellConflictMonitor`, `CellStructuralConflictMonitor`)
- Collab version history glue: [`packages/collab/versioning/src/index.ts`](../packages/collab/versioning/src/index.ts) (`createCollabVersioning`)
- Version store kept *inside the Y.Doc*: [`packages/versioning/src/store/yjsVersionStore.js`](../packages/versioning/src/store/yjsVersionStore.js) (`YjsVersionStore`)
- Branching glue: [`packages/collab/branching/index.js`](../packages/collab/branching/index.js) (`CollabBranchingWorkflow`)
- Branch graph store kept *inside the Y.Doc*: [`packages/versioning/branches/src/store/YjsBranchStore.js`](../packages/versioning/branches/src/store/YjsBranchStore.js) (`YjsBranchStore`)
- BranchService + snapshot adapter: [`packages/versioning/branches/src/`](../packages/versioning/branches/src/) (`BranchService`, `yjsDocToDocumentState`, `applyDocumentStateToYjsDoc`)

---

## Yjs workbook schema (roots + conventions)

### Root types

The collaborative workbook is a single shared `Y.Doc` with these primary roots:

- `cells`: `Y.Map<unknown>` keyed by **canonical cell keys**
- `sheets`: `Y.Array<Y.Map<unknown>>` where each entry is a sheet metadata map
- `metadata`: `Y.Map<unknown>` (workbook-level metadata)
- `namedRanges`: `Y.Map<unknown>` (named range definitions)
- `comments`: optional (see `@formula/collab-comments`; supports legacy schemas)

`@formula/collab-workbook` (`getWorkbookRoots`, `ensureWorkbookSchema`) is the canonical place that defines/normalizes these roots.

### Canonical cell keys

**Always use** the canonical cell key format:

```ts
const cellKey = `${sheetId}:${row}:${col}`; // 0-based row/col
```

Implementation:

- [`packages/collab/session/src/cell-key.js`](../packages/collab/session/src/cell-key.js) (`makeCellKey`)

Compatibility:

- The stack can *read* legacy keys like `${sheetId}:${row},${col}` and `r{row}c{col}` (tests / historical encodings), but new code should only *write* canonical keys.

### Cell value schema (`cells.get(cellKey)`)

Each cell is stored as a `Y.Map` with the following relevant fields:

- `value`: any JSON-serializable scalar/object (only when not a formula)
- `formula`: string (normalized; when present, `value` is set to `null`)
- `format`: JSON object for cell formatting (interned into `DocumentController.styleTable` on desktop)
- `enc`: optional encrypted payload (see “Cell encryption” below)
- `modified`: `number` (ms since epoch; best-effort)
- `modifiedBy`: `string` (best-effort user id)

Formula normalization is consistent across the stack: formula strings are treated as canonical when they start with `"="` and have no leading/trailing whitespace (see binder/session implementations).

### Cell encryption (`enc`) (protected ranges)

Because Yjs updates are broadcast to **all** collaborators (and the sync server), the server cannot filter per-cell content per connection. For true confidentiality of protected ranges, Formula supports optional **end-to-end encryption** of cell contents *before* they are written into the shared CRDT.

Encrypted cell data is stored on the per-cell Y.Map under the `enc` field:

```ts
// Stored under `cells.get(cellKey).get("enc")`
type EncryptedCellPayloadV1 = {
  v: 1;
  alg: "AES-256-GCM";
  keyId: string;
  ivBase64: string;
  tagBase64: string;
  ciphertextBase64: string;
};
```

Implementation:

- Encryption codec: [`packages/collab/encryption/src/index.node.js`](../packages/collab/encryption/src/index.node.js) (`encryptCellPlaintext`, `decryptCellPlaintext`, `isEncryptedCellPayload`)
- Session + binder both treat **any** `enc` presence as “encrypted” (even if malformed) to avoid accidentally falling back to plaintext duplicates under legacy cell-key encodings.

Usage:

```ts
import { createCollabSession } from "@formula/collab-session";

const session = createCollabSession({
  connection: { wsUrl, docId, token },
  encryption: {
    keyForCell: ({ sheetId, row, col }) => {
      // Return a per-range key (or null for unencrypted cells).
      return null;
    },
    // Optional: force encryption for some cells even if a key exists.
    // shouldEncryptCell: (cell) => boolean
  },
});
```

Notes:

- Plaintext is JSON `{ value, formula, format? }` and is bound to `{ docId, sheetId, row, col }` via AES-GCM Additional Authenticated Data (AAD) to prevent replay across docs/cells.
- When `enc` is present, plaintext `value`/`formula` fields are omitted.
- If a collaborator does not have the right key, `@formula/collab-session` and the desktop binder will surface a masked value and **refuse plaintext writes** into that cell.

### Sheet schema (`sheets` array entries)

Each entry in `doc.getArray("sheets")` is a `Y.Map` with (at least):

```ts
type SheetViewState = {
  frozenRows: number;
  frozenCols: number;
  colWidths?: Record<string, number>;
  rowHeights?: Record<string, number>;
};

type Sheet = {
  id: string;
  name: string | null;
  view?: SheetViewState;
};
```

The `SheetViewState` shape matches:

- BranchService `SheetViewState`: [`packages/versioning/branches/src/types.js`](../packages/versioning/branches/src/types.js)
- Desktop `DocumentController` `SheetViewState`: [`apps/desktop/src/document/documentController.js`](../apps/desktop/src/document/documentController.js)

---

## Creating a collaboration session (`@formula/collab-session`)

`createCollabSession` is the top-level API for constructing a Yjs doc + sync provider + optional integrations.

Source: [`packages/collab/session/src/index.ts`](../packages/collab/session/src/index.ts)

Typical usage:

```ts
import { createCollabSession } from "@formula/collab-session";
import { IndexedDbCollabPersistence } from "@formula/collab-persistence/indexeddb";

const session = createCollabSession({
  connection: { wsUrl, docId, token },

  // Optional offline-first persistence (applied before connecting the provider).
  persistence: new IndexedDbCollabPersistence(),

  // Optional presence (builds a PresenceManager and exposes it on `session.presence`).
  presence: {
    user: { id: userId, name: userName, color: "#4c8bf5" },
    activeSheet: "Sheet1",
  },

  // Optional collaborative undo/redo (Y.UndoManager-backed).
  undo: { captureTimeoutMs: 750 },

  // Optional end-to-end cell encryption.
  encryption: { keyForCell },
});

// Ensure persisted updates are applied into the doc before syncing.
await session.whenLocalPersistenceLoaded();
```

Notes:

- If `persistence` is provided, `connection.docId` (or `options.docId`) must be set; otherwise the session throws.
- When `presence` is enabled, `session.presence` is a `PresenceManager` instance.

---

## Binding Yjs to the desktop workbook model

**Goal:** keep the desktop workbook state machine (`DocumentController`) in sync with the shared Yjs workbook.

The binder lives at:

- [`packages/collab/binder/index.js`](../packages/collab/binder/index.js) (`bindYjsToDocumentController`)

It synchronizes **cell value/formula/format** between:

- `Y.Doc` → `cells` root (`Y.Map`)
- Desktop `DocumentController` (see [`apps/desktop/src/document/documentController.js`](../apps/desktop/src/document/documentController.js))

> Note: the binder currently **does not** sync the `sheets` array (sheet creation/order/rename) or per-sheet `view` state. Those are stored in Yjs (see below) and are used by branching/versioning, but desktop live-binding is currently cell-focused.

### End-to-end wiring example

```ts
import { createCollabSession } from "@formula/collab-session";
import { IndexedDbCollabPersistence } from "@formula/collab-persistence/indexeddb";

import { DocumentController } from "../apps/desktop/src/document/documentController.js";
import { bindYjsToDocumentController } from "../packages/collab/binder/index.js";

const session = createCollabSession({
  connection: { wsUrl, docId, token },
  persistence: new IndexedDbCollabPersistence(),
  presence: { user: { id: userId, name: userName, color: "#4c8bf5" }, activeSheet: "Sheet1" },
  undo: { captureTimeoutMs: 750 },
  encryption: { keyForCell }, // optional
});

await session.whenLocalPersistenceLoaded();

const documentController = new DocumentController();

// Role + range restrictions come from your auth layer / server.
const permissions = { role: "editor", restrictions: [], userId };

const binder = bindYjsToDocumentController({
  ydoc: session.doc,
  documentController,

  // Optional: makes DocumentController->Yjs writes go through the collaborative UndoManager.
  // (Only present when `createCollabSession({ undo: ... })` is enabled.)
  undoService: session.undo,

  userId,
  permissions,
  encryption: session.getEncryptionConfig(),
});

// Later:
// binder.destroy();
// session.destroy();
```

### Convenience helper: `bindCollabSessionToDocumentController`

If you already have a `CollabSession`, `@formula/collab-session` also exports a small glue helper that calls the binder with sensible defaults:

```ts
import { bindCollabSessionToDocumentController } from "@formula/collab-session";

// Optional: enable role-based permissions on the session.
// session.setPermissions({ role: "editor", rangeRestrictions: [], userId });

const binder = await bindCollabSessionToDocumentController({
  session,
  documentController,
  userId,

  // Optional: pass an explicit undoService (not defaulted by the helper).
  // undoService: session.undo,
});
```

### Echo suppression (origins) + UndoManager exception

When you bind two reactive systems, you must avoid doing this:

1. Local edit (DocumentController) writes to Yjs
2. Yjs observer fires (because the doc changed)
3. Binder applies the same change back into DocumentController (wasted work + can cause feedback loops)

The binder prevents this using **Yjs transaction origins**:

- Local DocumentController→Yjs writes are wrapped in a Yjs transaction with a stable `origin` token (`binderOrigin` or `undoService.origin`).
- The Yjs→DocumentController observer checks `transaction.origin` and **ignores** transactions that originated from those known local origins.

However, when collaborative undo/redo is enabled, Yjs uses a `Y.UndoManager` instance as an origin for undo/redo application. If we treated the UndoManager itself as “local” for echo suppression, then:

- Undo/redo would mutate Yjs
- …but the binder would ignore it
- …and the desktop UI would not update

So the binder intentionally:

- ignores **origin tokens** used for local writes
- but **does not** ignore the UndoManager instance origin

This is implemented by filtering `undoService.localOrigins` and excluding any value that looks like an UndoManager (see `isUndoManager(...)` in the binder).

Practical rule of thumb:

- When the binder is active, treat `DocumentController` as the “source of truth” for local edits.
- If you also call `session.setCellValue` / `session.setCellFormula` directly, be careful when passing `undoService: session.undo`: those direct session writes use the same origin token and may be echo-suppressed from applying back into `DocumentController`. (This is why the helper `bindCollabSessionToDocumentController` does **not** default to `session.undo`.)

### Undo/redo semantics in collaborative mode

The desktop `DocumentController` maintains its own local history stack, but in collaborative mode it is **not** the canonical user-facing undo stack.

In a shared Yjs session you generally want undo/redo to:

- only revert the **local user’s** edits
- never undo remote collaborators’ changes

That behavior is provided by Yjs’ `UndoManager` (via `@formula/collab-undo`, exposed as `session.undo` when `createCollabSession({ undo: ... })` is enabled).

See: [`docs/adr/ADR-0004-collab-sheet-view-and-undo.md`](./adr/ADR-0004-collab-sheet-view-and-undo.md)

---

## Sheet view state storage and syncing

Per-sheet view state (frozen panes + row/col size overrides) is stored on each sheet entry in the `sheets` array:

- `doc.getArray("sheets").get(i).get("view")`

The `view` object uses the same `SheetViewState` shape as BranchService (and desktop `DocumentController`):

```ts
{
  frozenRows: 2,
  frozenCols: 1,
  colWidths: { "0": 120 },
  rowHeights: { "1": 40 }
}
```

Compatibility note:

- Some historical/experimental docs stored `frozenRows` / `frozenCols` as **top-level** fields directly on the sheet map.
- BranchService’s Yjs adapter explicitly supports both shapes (see `branchStateFromYjsDoc` in [`packages/versioning/branches/src/yjs/branchStateAdapter.js`](../packages/versioning/branches/src/yjs/branchStateAdapter.js)), but **new writes should use `view`**.

Because `sheets` is part of the shared Y.Doc, any edits to `view` will sync like any other Yjs update (subject to your provider/persistence).

Semantics note:

- `sheets[].view` is **shared workbook state** (similar to formatting), not per-user transient UI state.
- Per-user ephemeral state (cursor/selection/active sheet/viewport) should use Awareness/presence.

See: [`docs/adr/ADR-0004-collab-sheet-view-and-undo.md`](./adr/ADR-0004-collab-sheet-view-and-undo.md)

---

## Presence integration (Awareness → desktop renderer)

### PresenceManager

`@formula/collab-session` can construct a `PresenceManager` automatically when `presence: {...}` is provided.

- `session.presence` is a `PresenceManager` (or `null` if not enabled)
- `PresenceManager` writes/reads the Yjs Awareness state field `"presence"`
- `PresenceManager.subscribe()` emits **remote** presences on the active sheet (it intentionally ignores local cursor/selection changes to reduce re-render pressure)

Source: [`packages/collab/presence/src/presenceManager.js`](../packages/collab/presence/src/presenceManager.js)

### Desktop rendering (`apps/desktop/src/grid/presence-renderer/`)

The desktop grid has a canvas overlay renderer:

- [`apps/desktop/src/grid/presence-renderer/`](../apps/desktop/src/grid/presence-renderer/)

The renderer expects the same `cursor` + `selections` shapes that `PresenceManager` emits, so you can typically pass the array through directly:

```ts
import { PresenceRenderer } from "../apps/desktop/src/grid/presence-renderer/index.js";

const renderer = new PresenceRenderer();

// `getCellRect(row, col)` must return the cell’s pixel rect in the overlay canvas coordinate space.
const unsubscribe = session.presence.subscribe((presences) => {
  renderer.clear(overlayCtx);
  renderer.render(overlayCtx, presences, { getCellRect });
});
```

Also remember to drive the local presence from UI events:

- `session.presence.setActiveSheet(sheetId)`
- `session.presence.setCursor({ row, col })`
- `session.presence.setSelections([{ startRow, startCol, endRow, endCol }, ...])`

---

## Comments (shared `comments` root)

Formula supports shared cell comments stored in the collaborative Y.Doc under the `comments` root.

Package:

- `@formula/collab-comments` (implementation: [`packages/collab/comments/src/`](../packages/collab/comments/src/))

The canonical schema is a `Y.Map` keyed by comment id. For backwards compatibility, some historical docs store comments as a `Y.Array<Y.Map>`.

Important rule: **don’t call** `doc.getMap("comments")` blindly on unknown documents; instantiating a legacy Array root as a Map can make the array content inaccessible. Instead, use `CommentManager` (which supports both) or run a migration first.

Recommended integration in collaborative mode:

```ts
import { createCommentManagerForSession } from "@formula/collab-comments";

// Ensures comment edits are wrapped in session.transactLocal(...) so they participate
// correctly in collaborative undo (when enabled).
const comments = createCommentManagerForSession(session);

const commentId = comments.addComment({
  cellRef: "A1",
  kind: "threaded",
  content: "Hello",
  author: { id: userId, name: userName },
});

comments.addReply({
  commentId,
  content: "First reply",
  author: { id: userId, name: userName },
});
```

If you need to normalize legacy Array-backed docs to the canonical Map schema:

- `migrateCommentsArrayToMap(doc)` (see [`packages/collab/comments/src/manager.ts`](../packages/collab/comments/src/manager.ts))

---

## Conflict monitoring (optional)

`@formula/collab-session` can attach conflict monitors that detect *true* offline/concurrent edits that overwrite one another.

These are implemented in `@formula/collab-conflicts` (see [`packages/collab/conflicts/index.js`](../packages/collab/conflicts/index.js)) and are enabled by passing options into `createCollabSession`:

```ts
import { createCollabSession } from "@formula/collab-session";

const session = createCollabSession({
  connection: { wsUrl, docId, token },

  // Detect same-cell formula conflicts (and optionally formula-vs-value conflicts).
  formulaConflicts: {
    localUserId: userId,
    mode: "formula", // or "formula+value"
    onConflict: (conflict) => {
      // conflict.remoteUserId is best-effort and may be an empty string.
      console.log("formula conflict", conflict);
    },
  },

  // Detect move/delete-vs-edit conflicts using the shared `cellStructuralOps` log.
  cellConflicts: {
    localUserId: userId,
    onConflict: (conflict) => console.log("structural conflict", conflict),
  },

  // Optional: detect value-vs-value conflicts when you are NOT using formula+value mode above.
  cellValueConflicts: {
    localUserId: userId,
    onConflict: (conflict) => console.log("value conflict", conflict),
  },
});
```

Notes:

- Concurrency detection is **causal** and based on Yjs map entry Item origin ids (not wall-clock timestamps).
- `remoteUserId` attribution is best-effort and may be empty if the overwriting writer did not update `modifiedBy`.

---

## Version history (collab mode)

Use `@formula/collab-versioning` to store and restore workbook snapshots in a collaborative session.

Source: [`packages/collab/versioning/src/index.ts`](../packages/collab/versioning/src/index.ts)

### “Use this in collaborative mode” snippet

```ts
import { createCollabVersioning } from "@formula/collab-versioning";

const versioning = createCollabVersioning({
  session,
  user: { userId, userName },
  // store: defaults to YjsVersionStore (history inside the shared Y.Doc)
});

await versioning.createSnapshot({ description: "Auto" });
await versioning.createCheckpoint({ name: "Q3 Approved", annotations: "Signed off" });

const versions = await versioning.listVersions();
await versioning.restoreVersion(versions[0].id);
```

### How history is stored (YjsVersionStore)

By default, `createCollabVersioning` uses `YjsVersionStore`, which stores all versions **inside the same shared Y.Doc** under these roots:

- `versions` (records)
- `versionsMeta` (ordering + metadata)

Important detail: when the store is `YjsVersionStore`, `CollabVersioning` excludes `versions` and `versionsMeta` from snapshots/restores to avoid:

- recursive snapshots (history containing itself)
- restores rolling back version history

See the `excludeRoots` logic in [`packages/collab/versioning/src/index.ts`](../packages/collab/versioning/src/index.ts).

---

## Branching + merging (collab mode)

Formula supports Git-like **branch + merge** workflows for collaborative spreadsheets by storing the branch/commit graph **inside the shared Y.Doc**.

### Required Yjs roots (default `rootName = "branching"`)

The Yjs-backed branch store reserves:

- `branching:branches` (`Y.Map<branchName, Y.Map>`)
- `branching:commits` (`Y.Map<commitId, Y.Map>`)
- `branching:meta` (`Y.Map`, including `rootCommitId` and `currentBranchName`)

Source: [`packages/versioning/branches/src/store/YjsBranchStore.js`](../packages/versioning/branches/src/store/YjsBranchStore.js)

### “Use this in collaborative mode” snippet (CollabBranchingWorkflow)

```ts
import { CollabBranchingWorkflow } from "../packages/collab/branching/index.js";
import { BranchService, YjsBranchStore } from "../packages/versioning/branches/src/index.js";

const store = new YjsBranchStore({ ydoc: session.doc }); // writes under branching:*
const branchService = new BranchService({ docId, store });

// Collab wrapper applies checkout/merge results back into the *live* Y.Doc.
const workflow = new CollabBranchingWorkflow({ session, branchService });

// Initialize (creates main branch if needed).
await branchService.init({ userId, role: "owner" }, /* initialState */ { sheets: {} });

await workflow.createBranch({ userId, role: "owner" }, { name: "scenario-a" });
await workflow.checkoutBranch({ userId, role: "owner" }, { name: "scenario-a" });

// Snapshot current shared workbook into a commit.
await workflow.commitCurrentState({ userId, role: "owner" }, "Try new model");
```

### Global checkout semantics (current design)

Branch checkout and merge currently mutate the **shared workbook roots** (`cells`, `sheets`, `metadata`, `namedRanges`, `comments`) of the live document. That means:

- all collaborators observe the checkout/merge result
- branching acts like a shared “scenario mode”, not a per-user view

This is enforced by `CollabBranchingWorkflow` calling `applyDocumentStateToYjsDoc(session.doc, ...)` after checkout/merge.

---

## Related docs

- Workstream overview: [`instructions/collaboration.md`](../instructions/collaboration.md)
- ADR (sheet view + undo semantics): [`docs/adr/ADR-0004-collab-sheet-view-and-undo.md`](./adr/ADR-0004-collab-sheet-view-and-undo.md)
