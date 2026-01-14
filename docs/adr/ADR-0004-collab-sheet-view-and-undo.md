# ADR-0004: Collaboration semantics for sheet view state and undo

- **Status:** Accepted
- **Date:** 2026-01-12

## Context

The collaboration stack uses **Yjs** as the shared CRDT, with the desktop UI still
largely modeled around the local `DocumentController`.

During development we ended up with two implicit assumptions that were not
written down:

1. **Sheet “view state”** (frozen panes, column widths, row heights, etc) is
   stored in the shared Yjs workbook under `sheets[].view`.
2. **Undo/redo** exists in two places:
   - `DocumentController` maintains a local history stack.
   - Collaborative sessions can enable Yjs’ `UndoManager` via `@formula/collab-undo`.

Without an explicit decision, downstream implementers may accidentally:

- store per-user UI preferences (scroll position, selection, zoom) in `sheets[].view`,
  making them globally shared and noisy,
- or wire UI undo/redo to `DocumentController` while in collaborative mode, causing
  incorrect undo semantics and subtle divergence.

This ADR clarifies the intended semantics.

## Decision

### 1) Sheet view state is **shared workbook state** stored under `sheets[].view`

The canonical home for per-sheet view/layout state is the **shared** Yjs workbook
schema:

- Root: `sheets` (`Y.Array<Y.Map>`)
- Per-sheet key: `view`
- Canonical path: `sheets[i].get("view")` / `sheets[].view`

This `view` object contains (at minimum) the same fields as the desktop
`DocumentController`’s `SheetViewState`:

- `frozenRows: number`
- `frozenCols: number`
- `colWidths?: Record<string, number>` (sparse overrides keyed by 0-based col index)
- `rowHeights?: Record<string, number>` (sparse overrides keyed by 0-based row index)
- `mergedRanges?: Array<{ startRow, endRow, startCol, endCol }>` (Excel-style merged cells; inclusive end coords)

In addition, desktop clients may store other sheet-level layout metadata under
`sheets[].view` when it is intended to be **shared workbook state**, e.g.:

- `drawings?: any[]` (pictures/shapes/charts placement metadata; stable-sorted by `zOrder`)

**These fields are globally shared across collaborators.**

Rationale:

- Frozen panes and row/col sizes are part of the *workbook’s persisted layout*
  (similar to formatting), not a transient UI preference.
- Branching/versioning snapshots already treat `sheets[].view` as part of the
  shared document state (see `branchStateFromYjsDoc` reading `view`).
- Keeping layout deterministic across clients simplifies collaborative editing
  and ensures exports / re-imports are consistent.

Compatibility note:

- Some historical/experimental documents may store `frozenRows` / `frozenCols` as
  **top-level keys on the sheet map**. Treat those as legacy input; new code
  should write to `sheets[].view` going forward.

#### Non-goal: per-user sheet preferences in `sheets[].view`

Per-user, ephemeral UI state MUST NOT be stored in shared workbook state. Examples:

- scroll position / viewport
- current selection / active cell
- active sheet tab
- local UI-only affordances (hover, edit mode)

Those belong in **Awareness** (presence) or other local-only storage.

Future path for persistent per-user preferences:

- Use **Awareness** for real-time per-user state (cursor/selection/etc).
- Persist per-user preferences outside the shared workbook (e.g. app storage keyed
  by `{docId, userId}`), or add a dedicated “user preferences” layer that is not
  included in exports/versioning snapshots. Do not overload `sheets[].view`.

---

### 2) In collaborative mode, the canonical user-facing undo stack is Yjs UndoManager

#### Single-user/local mode

When not collaborating (no `CollabSession` / no shared Yjs doc), the canonical
state machine is:

- `DocumentController` owns the workbook state and its history.
- UI undo/redo is wired to the `DocumentController`’s internal undo stack.

#### Collaborative mode (Yjs-backed)

When collaborating (shared Yjs document is the source of truth):

- **User-facing undo/redo MUST be backed by Yjs’ `UndoManager`**, exposed via
  `@formula/collab-undo` (e.g. `CollabSession.undo`).
- `DocumentController`’s internal undo/history is **not** the canonical undo stack
  and must not be wired to user undo UI in collaborative mode.

Rationale:

- Collaborative undo must only revert the *local user’s* changes; it must never
  undo other collaborators’ edits. Yjs `UndoManager` tracks changes by origin to
  enforce this.
- A local-only undo stack in `DocumentController` cannot correctly model merged
  remote edits and will not be consistent across clients.

#### Required transaction semantics (`CollabSession.origin` / `session.transactLocal`)

To ensure edits are undoable (and classified as local vs remote):

- All shared mutations must run in a local-origin Yjs transaction.
- The canonical API is `session.transactLocal(() => { ... })`, which uses
  `CollabSession.origin` (and, when enabled, the `@formula/collab-undo` wrapper).

For feature code that writes to Yjs directly (not through a helper manager):

```ts
session.transactLocal(() => {
  // mutate session.doc / session.cells / session.sheets / ...
});
```

For feature code that uses managers, prefer helpers that bind a manager to the
session and already wrap mutations in `session.transactLocal`, e.g.:

- `createSheetManagerForSession(session)` (`@formula/collab-workbook`)
- `createNamedRangeManagerForSession(session)` (`@formula/collab-workbook`)
- `createMetadataManagerForSession(session)` (`@formula/collab-workbook`)
- `createCommentManagerForSession(session)` (`@formula/collab-comments`)

Desktop note: the desktop app sometimes uses a *binder-origin* undo scope (tracking
DocumentController→Yjs transactions under a distinct origin token, instead of
`session.origin`). In that setup, comment mutations must run inside the
binder-origin transact wrapper (e.g. `createCommentManagerForDoc({ doc: session.doc, transact: undoService.transact })`)
so they participate in the same undo stack as cell edits.

---

### 3) Interaction with `bindYjsToDocumentController` and “local-only undo”

`bindYjsToDocumentController` binds cell edits (and sheet view state deltas) between:

- the shared Yjs CRDT, and
- the desktop `DocumentController` (as a local projection for the UI).

In collaborative mode:

1. **Do not expose `DocumentController` undo/redo to the user.**
2. Wire UI undo/redo to the collaborative undo service:
   - `session.undo.undo()` / `session.undo.redo()` (or the underlying `UndoService`)
3. Ensure DocumentController-origin edits are captured by Yjs UndoManager:
   - Pass an `undoService` (from `@formula/collab-undo`) into `bindYjsToDocumentController`.

`bindYjsToDocumentController` prevents feedback loops between the two reactive
systems using internal guards (not by blanket-filtering `transaction.origin`):

- While applying a **DocumentController → Yjs** write it sets `applyingLocal`.
  The Yjs `observeDeep` handlers early-return when this flag is set, so the
  deep observer event that immediately follows the binder’s own write is ignored.
- While applying a **Yjs → DocumentController** update it sets `applyingRemote`.
  The DocumentController `"change"` handler ignores events while this flag is set,
  so remote-applied deltas are not written back into Yjs.

This is intentionally not implemented as “ignore all local origins”, because the
stack reuses local origins for programmatic operations (branch checkout/merge
apply, version restore, direct `session.setCell*`) that must still update the
desktop projection.

Origins still matter for **undo/conflict semantics**:

- If you pass `undoService: session.undo` (or use `bindCollabSessionToDocumentController`,
  which defaults to `session.transactLocal(...)`), DocumentController edits use
  `session.origin`, participate in collaborative undo, and are treated as local by
  conflict monitors.
- If you do not, the binder uses its own origin token and those edits will not be
  undoable by collaborative undo tracking.

Practical guidance:

- **Preferred (desktop UI):** make cell edits via `DocumentController` methods and
  let the binder write them to Yjs.
- For non-cell shared state (comments, metadata, sheets, `sheets[].view`, etc),
  mutate Yjs directly but do so inside `session.transactLocal`.

## Consequences

- Frozen panes and row/col sizing are shared and persist as part of the workbook.
- Per-user UI state (selection/cursor/viewport) is not stored in shared workbook
  state; it uses Awareness/local storage.
- In collaborative mode, the only user-facing undo/redo stack is the Yjs-based
  undo service (`@formula/collab-undo`). `DocumentController` undo is single-user only.

## Current implementation pointers

- Shared workbook roots: `@formula/collab-workbook` (`packages/collab/workbook/src/index.ts`)
- Collaborative session + origins + `transactLocal`: `@formula/collab-session`
  (`packages/collab/session/src/index.ts`)
- Collaborative undo service: `@formula/collab-undo`
  (`packages/collab/undo/src/undo-service.js`, `yjs-undo-service.js`)
- DocumentController sheet view state (`SheetViewState`): `apps/desktop/src/document/documentController.js`
- DocumentController↔Yjs binder: `bindYjsToDocumentController`
  (`packages/collab/binder/index.js`)
- Branch snapshot adapter reading `sheets[].view`: `packages/versioning/branches/src/yjs/branchStateAdapter.js`
