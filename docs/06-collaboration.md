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
- Workbook metadata managers: [`packages/collab/workbook/src/index.ts`](../packages/collab/workbook/src/index.ts) (`SheetManager`, `MetadataManager`, `NamedRangeManager`, `createSheetManagerForSession`, `createMetadataManagerForSession`, `createNamedRangeManagerForSession`)
- Cell key helpers: [`packages/collab/session/src/cell-key.js`](../packages/collab/session/src/cell-key.js) (`makeCellKey`, `parseCellKey`, `normalizeCellKey`)
- Desktop binder: [`packages/collab/binder/index.js`](../packages/collab/binder/index.js) (`bindYjsToDocumentController`)
- Desktop sheet view binder: [`apps/desktop/src/collab/sheetViewBinder.ts`](../apps/desktop/src/collab/sheetViewBinder.ts) (`bindSheetViewToCollabSession`)
- Collaborative undo: [`packages/collab/undo/index.js`](../packages/collab/undo/index.js) (`createUndoService`, `REMOTE_ORIGIN`)
- Cell encryption: [`packages/collab/encryption/src/index.node.js`](../packages/collab/encryption/src/index.node.js) (`encryptCellPlaintext`, `decryptCellPlaintext`)
- Encrypted/protected range metadata + policy: [`packages/collab/encrypted-ranges/src/index.ts`](../packages/collab/encrypted-ranges/src/index.ts) (`EncryptedRangeManager`, `createEncryptionPolicyFromDoc`, `createEncryptedRangeManagerForSession`)
- Presence (Awareness wrapper): [`packages/collab/presence/src/presenceManager.js`](../packages/collab/presence/src/presenceManager.js) (`PresenceManager`)
- Desktop presence renderer: [`apps/desktop/src/grid/presence-renderer/`](../apps/desktop/src/grid/presence-renderer/) (`PresenceRenderer`)
- Permissions + masking: [`packages/collab/permissions/index.js`](../packages/collab/permissions/index.js) (`getCellPermissions`, `maskCellValue`)
- Local persistence implementations: [`packages/collab/persistence/src/`](../packages/collab/persistence/src/) (`IndexedDbCollabPersistence`, `FileCollabPersistence`)
- Comments (Yjs `comments` root helpers): [`packages/collab/comments/src/manager.ts`](../packages/collab/comments/src/manager.ts) (`CommentManager`, `createCommentManagerForSession`, `createCommentManagerForDoc`, `migrateCommentsArrayToMap`)
- Conflict monitors: [`packages/collab/conflicts/index.js`](../packages/collab/conflicts/index.js) (`FormulaConflictMonitor`, `CellConflictMonitor`, `CellStructuralConflictMonitor`)
- Collab version history glue: [`packages/collab/versioning/src/index.ts`](../packages/collab/versioning/src/index.ts) (`createCollabVersioning`)
- Version store kept *inside the Y.Doc*: [`packages/versioning/src/store/yjsVersionStore.js`](../packages/versioning/src/store/yjsVersionStore.js) (`YjsVersionStore`)
- Version store backed by the Formula API (cloud DB): [`packages/versioning/src/store/apiVersionStore.js`](../packages/versioning/src/store/apiVersionStore.js) (`ApiVersionStore`)
- Version store backed by local SQLite (Node/desktop): [`packages/versioning/src/store/sqliteVersionStore.js`](../packages/versioning/src/store/sqliteVersionStore.js) (`SQLiteVersionStore`)
- Branching glue: [`packages/collab/branching/index.js`](../packages/collab/branching/index.js) (`CollabBranchingWorkflow`)
- Branch graph store kept *inside the Y.Doc*: [`packages/versioning/branches/src/store/YjsBranchStore.js`](../packages/versioning/branches/src/store/YjsBranchStore.js) (`YjsBranchStore`)
- Branch graph store backed by local SQLite (Node/desktop): [`packages/versioning/branches/src/store/SQLiteBranchStore.js`](../packages/versioning/branches/src/store/SQLiteBranchStore.js) (`SQLiteBranchStore`)
- BranchService + browser-safe entrypoint: [`packages/versioning/branches/src/browser.js`](../packages/versioning/branches/src/browser.js) (`BranchService`, `YjsBranchStore`, `yjsDocToDocumentState`, `applyDocumentStateToYjsDoc`)

---

## Performance regression benchmarks (binder + session)

Collab binder/session code is performance sensitive (it may process **tens of thousands** of cell updates when hydrating a doc, applying a version restore, or merging branches). To help catch accidental **O(N²)** regressions, the repo includes **opt-in** Node-based perf benchmarks:

- Binder-only: [`packages/collab/binder/test/perf/binder-perf.test.js`](../packages/collab/binder/test/perf/binder-perf.test.js)
  - README: [`packages/collab/binder/test/perf/README.md`](../packages/collab/binder/test/perf/README.md)
- Session + binder (exercises `bindCollabSessionToDocumentController`): [`packages/collab/session/test/perf/session-binder-perf.test.js`](../packages/collab/session/test/perf/session-binder-perf.test.js)
  - README: [`packages/collab/session/test/perf/README.md`](../packages/collab/session/test/perf/README.md)

These tests are **skipped by default** and only run when the relevant env var is set:

```bash
# Binder perf
FORMULA_RUN_COLLAB_BINDER_PERF=1 NODE_OPTIONS=--expose-gc FORMULA_NODE_TEST_CONCURRENCY=1 pnpm test:node packages/collab/binder/test/perf/binder-perf.test.js

# Session + binder perf
FORMULA_RUN_COLLAB_SESSION_BINDER_PERF=1 NODE_OPTIONS=--expose-gc FORMULA_NODE_TEST_CONCURRENCY=1 pnpm test:node packages/collab/session/test/perf/session-binder-perf.test.js
```

Tip: set `PERF_JSON=1` to emit structured JSON metrics per scenario (useful for CI parsing), and optionally set `PERF_MAX_TOTAL_MS_*` / `PERF_MAX_PEAK_*` env vars to enforce budgets.

Common knobs:

- `PERF_CELL_UPDATES`, `PERF_BATCH_SIZE`, `PERF_COLS` (workload sizing)
- `PERF_SCENARIO=yjs-to-dc|dc-to-yjs|all` (run one direction only)
- `PERF_KEY_ENCODING=canonical|legacy|rxc` (Yjs→DC runs; benchmark key normalization)
- `PERF_INCLUDE_FORMAT=1` / `PERF_FORMAT_VARIANTS` (exercise formatting paths)
- `PERF_INCLUDE_GUARDS=0` (binder perf only; disable canRead/canEdit hooks)

There is also a manual GitHub Actions workflow to run these in CI: [`.github/workflows/collab-perf.yml`](../.github/workflows/collab-perf.yml).

## Yjs workbook schema (roots + conventions)

### Root types

The collaborative workbook is a single shared `Y.Doc` with these primary roots:

- `cells`: `Y.Map<unknown>` keyed by **canonical cell keys**
- `sheets`: `Y.Array<Y.Map<unknown>>` where each entry is a sheet metadata map
- `metadata`: `Y.Map<unknown>` (workbook-level metadata)
- `namedRanges`: `Y.Map<unknown>` (named range definitions)
- `comments`: optional (see `@formula/collab-comments`; supports legacy schemas)

Note: even “empty” cells may still exist in Yjs as marker-only `Y.Map`s (for example, to
record a causal `formula = null` clear for deterministic conflict detection; see “Conflict
monitoring” below). `CollabSession.getCell()` treats these marker-only cells as empty UI
state (returns `null`) even though the underlying Yjs map entry is preserved for causality.
If you don’t need marker-only causality history (for example, when conflict monitors are
disabled), you can optionally prune these entries with `CollabSession.compactCells(...)`
(see “Cell-map compaction” below).

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

- `value`: any JSON-serializable scalar/object (only when not a formula; when setting a formula string, set `value` to `null`)
- `formula`: `string | null` (may be written non-canonically; most readers normalize to a canonical `=...` form. Clears should be represented as `null` — older writers may omit/delete the key, which readers should treat as `null`)
- `format`: JSON object for cell formatting (interned into `DocumentController.styleTable` on desktop)
- `enc`: optional encrypted payload (see “Cell encryption” below)
- `modified`: `number` (ms since epoch; best-effort)
- `modifiedBy`: `string` (best-effort user id; in some deployments the sync-server may rewrite this to the authenticated user id for touched cells when range restriction enforcement is enabled)

Important nuance (conflict monitoring): for deterministic delete-vs-overwrite detection,
`@formula/collab-conflicts`’ `FormulaConflictMonitor` expects local formula clears to be
written as `cell.set("formula", null)` (not `cell.delete("formula")`). Yjs map deletes do
not create a new Item, which makes causality ambiguous.

See: [`packages/collab/conflicts/src/formula-conflict-monitor.js`](../packages/collab/conflicts/src/formula-conflict-monitor.js)

Implementation detail:

- When `createCollabSession({ formulaConflicts: ... })` is enabled, `session.setCellFormula(...)`
  delegates to `FormulaConflictMonitor.setLocalFormula(...)`, which writes the `null` marker for clears.
- Value writes clear formulas via a `null` marker (even when value conflicts are not being monitored) so later formula writes can causally reference the clear via Yjs Item `origin`. This is important for cross-client determinism when some collaborators run conflict monitors and others do not.

Formula normalization: most consumers normalize formula strings to a canonical **text** form (`=...` with surrounding whitespace stripped). The desktop binder normalizes formulas on read/write, and branching/versioning adapters normalize formulas when producing snapshots/diffs. Direct `session.setCellFormula(...)` writes preserve the input aside from basic trimming, so UI code should prefer writing canonical formula text.

Example (direct Yjs write):

```ts
import * as Y from "yjs";

class CollaborativeDocument {
  constructor(readonly doc: Y.Doc) {}

  setCell(sheetId: string, row: number, col: number, value: unknown, formula?: string | null): void {
    this.doc.transact(() => {
      const cells = this.doc.getMap("cells");
      const cellKey = `${sheetId}:${row}:${col}`;

      let cellData = cells.get(cellKey);
      if (!(cellData instanceof Y.Map)) {
        cellData = new Y.Map();
        cells.set(cellKey, cellData);
      }

      if (formula) {
        cellData.set("formula", formula);
        cellData.set("value", null);
      } else {
        // Don't delete: map deletes do not create Items; null markers preserve causality
        // for deterministic delete-vs-overwrite conflict detection.
        cellData.set("formula", null);
        cellData.set("value", value);
      }
    });
  }
}
```

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
- The WebCrypto implementation caches imported `CryptoKey`s by `keyId` in a bounded LRU cache (default max 256) to avoid unbounded growth and to limit retention of sensitive key material.
  - Configure max size via `globalThis.__FORMULA_ENCRYPTION_KEY_CACHE_MAX_SIZE__` (and in Node, `FORMULA_ENCRYPTION_KEY_CACHE_MAX_SIZE`).
  - Set max size to `0` to disable caching entirely.
  - Call `clearEncryptionKeyCache()` (exported by `@formula/collab-encryption`) during teardown/tests to release cached keys.
- Session + binder both treat **any** `enc` presence as “encrypted” (even if malformed) to avoid accidentally falling back to plaintext duplicates under legacy cell-key encodings.
- Binder write guard: when a cell already has an `enc` payload, `bindYjsToDocumentController` requires that the key resolver returns a **matching `keyId`** (and treats unknown payload schemas as non-writable). This prevents older clients from clobbering encrypted content they cannot decrypt (including future encryption versions) just because they have some unrelated key material. Desktop surfaces an actionable “unsupported format” toast when such writes are rejected.
- Session write guard: `CollabSession` enforces the same invariant for direct `setCell*` APIs — if a cell already has an `enc` payload, callers must provide a key with a matching `keyId` and a supported payload schema, otherwise the write is rejected. This prevents accidental key-id mismatches and avoids clobbering ciphertext from newer clients.
- Versioning diffs (`packages/versioning/src/yjs/*` + `semanticDiff`) treat `enc` as meaningful cell content for modified/moved/format-only detection without requiring decryption, and surface only minimal metadata (e.g. key id) in diff records.

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
    //
    // Optional: encrypt per-cell formatting (`format`) alongside value/formula.
    // Defaults to false for backwards compatibility (format remains plaintext).
    // encryptFormat: true,
  },
});
```

#### Shared encrypted-range metadata (`metadata.encryptedRanges`) (`@formula/collab-encrypted-ranges`)

For end-to-end encryption to be safe in collaborative mode, **all clients must agree on which cells are “protected” and must be written encrypted**, even if a given client does not have the actual key bytes. Formula stores this *policy metadata* in the shared workbook `metadata` root:

- Location: `getWorkbookRoots(doc).metadata.get("encryptedRanges")` (aka `doc.getMap("metadata").get("encryptedRanges")`)
- When you derive `encryption.shouldEncryptCell` from this metadata (see below), clients without keys can still refuse plaintext writes into protected cells.
- Important: **only metadata is shared** (range rectangles + `keyId`). **Secret key material must never be stored in the Y.Doc.**

Canonical schema (new writes should always use this):

- `metadata.encryptedRanges`: `Y.Array<Y.Map<unknown>>`
- Each entry is a `Y.Map` with:

```ts
type EncryptedRange = {
  id: string;
  sheetId: string; // stable workbook sheet id (not the display name)
  startRow: number; // 0-based, inclusive
  startCol: number; // 0-based, inclusive
  endRow: number; // 0-based, inclusive
  endCol: number; // 0-based, inclusive
  keyId: string; // key identifier (not key bytes)
  createdAt?: number;
  createdBy?: string;
};
```

Notes:

- `keyId` is a **stable identifier** for out-of-band key material. Do not “rotate” a key by overwriting the bytes for an existing `keyId`; existing ciphertext references the old bytes and would become undecryptable.

Legacy schemas tolerated (read support):

- `metadata.encryptedRanges` stored as `Y.Map<id, Y.Map>` (map key is treated as the range id)
- Array entries missing an explicit `id` (plain objects or `Y.Map`s); ids are derived deterministically as `legacy:...` so policy lookup still works
- Older fields such as `sheetName` / `sheet` instead of stable `sheetId`

Normalization + dedupe (write support):

- `EncryptedRangeManager` normalizes `metadata.encryptedRanges` into the canonical `Y.Array` + local `Y.Map` entry constructors **before** applying undo-tracked edits. This avoids Yjs `instanceof` pitfalls when a doc was hydrated using a different Yjs module instance (ESM vs CJS), and ensures collaborative undo only captures the user’s explicit change.
- During normalization it drops malformed entries and dedupes duplicates, including identical ranges with different ids (e.g. from concurrent inserts).
- If `metadata.encryptedRanges` is present but in an unknown schema (neither `Y.Array`, `Y.Map`, nor a plain JS array), `EncryptedRangeManager` throws rather than clobbering potentially-newer data.

APIs (source: [`packages/collab/encrypted-ranges/src/index.ts`](../packages/collab/encrypted-ranges/src/index.ts)):

- `EncryptedRangeManager`
  - `list(): EncryptedRange[]` (deterministic ordering)
    - Throws if `metadata.encryptedRanges` is present but in an unsupported schema (rather than silently returning `[]` and risking plaintext writes).
  - `add(range: { sheetId, startRow, startCol, endRow, endCol, keyId, createdAt?, createdBy? }): string`
    - `sheetId` should be the stable workbook sheet id; as a best-effort convenience, `EncryptedRangeManager` will resolve a sheet *display name* to its id when possible (useful for legacy/UI-driven callsites).
  - `update(id: string, patch: Partial<...>): void`
    - Like `add`, `patch.sheetId` is best-effort resolved from sheet display name → stable id when possible.
  - `remove(id: string): void`
    - Note: removing an encrypted range only updates the shared **policy metadata**; it does **not** decrypt cells that are already stored with an `enc` payload. Those cells remain encrypted until rewritten by a client with the correct key.
- `createEncryptedRangeManagerForSession(session)` → `EncryptedRangeManager`
  - Uses `session.transactLocal(...)` so range edits participate in the session’s local-origin collaborative undo scope (when undo is enabled).
  - If your application uses a separate Yjs `UndoManager` / origin token, prefer `new EncryptedRangeManager({ doc, transact })` with a `transact` that uses that origin so range edits are undoable.
- `createEncryptionPolicyFromDoc(doc)` → `{ shouldEncryptCell(cell): boolean; keyIdForCell(cell): string | null }`
  - Reads `metadata.encryptedRanges` (including legacy schemas) to answer:
    - **should this cell be encrypted on write?** (`shouldEncryptCell`)
    - **which key id applies?** (`keyIdForCell`)
  - Fail-closed: if `metadata.encryptedRanges` is present but in an unknown schema, `shouldEncryptCell` returns `true` for all valid cells (so keyless clients refuse plaintext writes), and `keyIdForCell` returns `null`.
    - Desktop: when this condition is detected in collaboration mode, the app surfaces an error toast since editing may be blocked across the workbook until the client is updated.
  - Sheet matching is case-insensitive and tolerates callers passing either a stable sheet id or (when sheets metadata is available) a sheet display name.
  - Overlap rule: when multiple ranges match:
    - canonical array schema (`Y.Array`): the most recently added match wins (last entry wins)
    - legacy map schema (`Y.Map<id, ...>`): the lexicographically greatest key wins (deterministic ordering)

#### Wiring pattern: derive `shouldEncryptCell` from shared metadata

Recommended pattern (allows keyless clients to refuse plaintext writes into protected ranges):

```ts
import * as Y from "yjs";
import { createCollabSession } from "@formula/collab-session";
import {
  createEncryptionPolicyFromDoc,
  createEncryptedRangeManagerForSession,
} from "@formula/collab-encrypted-ranges";

const doc = new Y.Doc({ guid: docId });
const policy = createEncryptionPolicyFromDoc(doc);

const session = createCollabSession({
  doc,
  connection: { wsUrl, docId, token },
  encryption: {
    shouldEncryptCell: policy.shouldEncryptCell,
    keyForCell: (cell) => {
      const keyId = policy.keyIdForCell(cell);
      if (!keyId) return null;

      // Resolve key bytes out-of-band (KMS/keyring). DO NOT store these in the doc.
      const keyBytes = keyring.getAes256KeyBytes(keyId);
      return keyBytes ? { keyId, keyBytes } : null;
    },
  },
});

// Mutate the shared metadata via a manager that uses transactLocal for undo scope.
const ranges = createEncryptedRangeManagerForSession(session);
ranges.add({ sheetId: "Sheet1", startRow: 0, startCol: 0, endRow: 9, endCol: 9, keyId: "k1" });
```

Tests worth reading:

- Policy enforcement in session (keyless clients block plaintext writes): [`packages/collab/session/test/encrypted-ranges.policy.test.js`](../packages/collab/session/test/encrypted-ranges.policy.test.js)
- Undo + foreign-Yjs normalization behavior: [`packages/collab/session/test/encrypted-ranges.undo.test.js`](../packages/collab/session/test/encrypted-ranges.undo.test.js)

Reminder: encryption is **orthogonal** to collaboration permissions. Viewer/commenter role enforcement and comment-only restrictions are handled separately (see “Permissions + masking” below); encrypted ranges only define confidentiality + “must-encrypt” write policy.

Desktop dev/testing toggle:

- The desktop app supports a **dev-only** URL param toggle to exercise encrypted cell payloads end-to-end:
  - `?collabEncrypt=1` enables deterministic encryption for a demo range
  - `?collabEncryptRange=Sheet1!A1:C10` overrides the encrypted range (default `Sheet1!A1:C10`)
    - The sheet qualifier uses formula-style sheet names; desktop resolves these to stable sheet ids when possible.
- This is implemented in `apps/desktop/src/collab/devEncryption.ts` and is intended for manual verification (two clients: one with the key, one without).
- The key is derived deterministically from `docId` + a hardcoded dev salt (not production key management).

Desktop encryption commands (Command Palette):

- `collab.encryptSelectedRange` — encrypt the current selection and generate a shareable key string.
  - If the entered `keyId` already exists in the local key store, the desktop UI will reuse it (it does not silently overwrite key bytes).
  - If the key store is unavailable and the UI cannot verify whether the key id already exists, it warns in the confirmation prompt before generating a new key.
  - If the `keyId` is already used by an existing encrypted range but the key is not imported, the UI refuses to generate a new key for that id (import the key first or choose a different id).
  - If the `keyId` already appears in existing encrypted cell payloads (even if the encrypted-range metadata was removed), the UI refuses to generate a new key for that id unless the key is already imported.
  - If the encrypted range metadata is unreadable (unsupported `metadata.encryptedRanges` schema), the UI aborts early to avoid generating/storing orphaned key material.
  - If a new key is generated/stored but adding the encrypted range fails, the UI best-effort deletes the new key (when it can verify the key id was previously missing) to avoid orphaned key material.
- `collab.removeEncryptedRange` — remove encrypted range *metadata* overlapping the current selection.
  - Note: removing a range does **not** decrypt cells that already have an `enc` payload.
  - If multiple ranges overlap, the desktop UI can remove a single chosen range or all overlaps.
  - Uses stable sheet ids when possible; name-based matching is best-effort and avoids sheet id/name ambiguity.
  - If the encrypted range metadata is unreadable (unsupported schema), the command aborts with an error toast.
- `collab.listEncryptedRanges` — list all encrypted ranges in the workbook and jump to one (selects it).
  - Resolves legacy ranges stored with a sheet display name (instead of a stable `sheetId`) when possible, and avoids sheet id/name ambiguity (also canonicalizes stable ids via their display name to tolerate case-mismatched legacy ids).
  - If the encrypted range metadata is unreadable (unsupported schema), the command aborts with an error toast.
- `collab.exportEncryptionKey` — export the key for the active cell’s encrypted range.
  - Prefers the `keyId` embedded in an existing encrypted cell payload (if present), otherwise falls back to policy metadata.
  - If the key bytes are missing locally, it prompts to import the key first.
  - If the encrypted range policy metadata is unreadable (unsupported `metadata.encryptedRanges` schema) and the active cell does not already contain an `enc` payload, the UI cannot determine the key id and surfaces an error.
- `collab.importEncryptionKey` — import a shared key string into the local key store.
  - If the key id already exists with different bytes, the UI prompts before overwriting.
  - If the UI cannot verify whether the key id already exists (key store unavailable), it prompts before importing.
  - After importing, the desktop app best-effort rehydrates the collab binder so already-encrypted cells can decrypt.

Notes:

- Desktop `CollabSession` wiring (`apps/desktop/src/app/spreadsheetApp.ts`) resolves encryption keys with this precedence:
  1) If the cell already has a valid `enc` payload and the referenced key is available locally, use `enc.keyId`.
  2) Otherwise, fall back to the shared encrypted-range policy (`metadata.encryptedRanges`) `keyIdForCell`.
  - This allows clients to keep decrypting existing encrypted cells even if the policy metadata is missing/out-of-sync, while still supporting key rotation/overwrite flows when the old ciphertext key is unavailable.
- Plaintext is JSON `{ value, formula, format? }` and is bound to `{ docId, sheetId, row, col }` via AES-GCM Additional Authenticated Data (AAD) to prevent replay across docs/cells.
  - The encryption codec supports an optional `format` field.
  - By default (`encryption.encryptFormat` unset/false), `@formula/collab-session` + the desktop binder only encrypt `value`/`formula` and leave per-cell formatting stored separately under the shared plaintext `format` key (legacy behavior).
  - When `encryption.encryptFormat=true`, per-cell formatting is included in the encrypted plaintext and the plaintext `format` key is removed from the Yjs cell map (as well as the legacy `style` alias key, if present). This prevents the sync server and unauthorized collaborators from learning cell-specific formatting metadata.
    - Backwards compatibility: encrypted documents created by older clients may have plaintext `format` (and/or no `format` inside the encrypted payload). For confidentiality, `encryptFormat=true` does **not** fall back to using plaintext `format` when `enc` is present; such cells render with default per-cell style until rewritten by a client that re-encrypts them with `encryptFormat=true`.
    - Note: sheet/row/col formatting defaults and compressed `formatRunsByCol` range-run formatting are stored outside the per-cell map and are not currently encrypted by this flag.
    - Expected impact: because AES-GCM encryption uses a random IV, any update that re-encrypts the cell (including a formatting-only change) will change the `enc` payload bytes even if the underlying value/formula did not change. Snapshot/diff tooling that treats `enc` as an opaque blob may therefore report “cell changed” without being able to attribute it specifically to a formatting-only change. Versioning/diff UX may need follow-up work to surface these changes more semantically for authorized users.
- When `enc` is present, plaintext `value`/`formula` fields are omitted.
- If a collaborator does not have the right key, `@formula/collab-session` and the desktop binder will surface a masked value and **refuse plaintext writes** into that cell.

### Permissions + masking (roles + range restrictions)

Formula supports a simple role model plus optional per-range allowlists:

- Roles: `owner | admin | editor | commenter | viewer`
- Range restrictions: each restriction can include a `readAllowlist` and/or `editAllowlist` for a rectangular range

Implementation: `@formula/collab-permissions` (see [`packages/collab/permissions/index.js`](../packages/collab/permissions/index.js)).

Important: **masking is not confidentiality**. Masking is a UX/access-control measure; for true confidentiality use end-to-end encryption (`enc`) as described above.

#### Session-level permissions (`CollabSession`)

`CollabSession` exposes permission-aware helpers:

- `session.setPermissions({ role, rangeRestrictions, userId })`
- `session.getPermissions()`
- `session.onPermissionsChanged((permissions) => { ... })` (calls immediately; returns `unsubscribe()`)
- `session.canReadCell({ sheetId, row, col })`
- `session.canEditCell({ sheetId, row, col })`
- Convenience role helpers: `session.getRole()`, `session.isReadOnly()`, `session.canComment()`, `session.canShare()`

`setPermissions` validates and normalizes `rangeRestrictions` eagerly (using `@formula/collab-permissions`), so misconfigured restrictions fail fast with an actionable error message (e.g. `rangeRestrictions[3] invalid: ...`).

These checks also incorporate encryption invariants (e.g. refusing writes to encrypted cells when no key is available).

Default behavior when permissions are unset:

- If `setPermissions(...)` has never been called, `session.getRole()` returns `null` and role helpers default to:
  - `session.isReadOnly() === false` (editable by default)
  - `session.canComment() === false`, `session.canShare() === false`
- `canReadCell` / `canEditCell` default to permissive role behavior, but still enforce **encryption invariants**
  (encrypted cells require a key to read/write).

#### Desktop binder permissions (`bindYjsToDocumentController`)

The desktop binder can enforce permissions and masking at the UI projection layer:

- Unreadable cells are masked in `DocumentController` (default mask is `"###"`).
- Optionally, per-cell formatting for masked cells can also be suppressed with `maskCellFormat: true` (defaults to `false`; clears the per-cell `styleId` to `0` when a cell is masked due to permissions or missing encryption keys).
- Disallowed edits are rejected and reverted (optionally surfaced via `onEditRejected`).
- Shared-state writes (sheet view + sheet-level formatting defaults) can be gated via `canWriteSharedState` (used by `bindCollabSessionToDocumentController` to suppress these writes for read-only roles).

The binder accepts permission info either as:

- a function: `permissions(cell) -> { canRead, canEdit }`, or
- a role-based object: `{ role, restrictions, userId }`

In addition (legacy / composition):

- `canReadCell(cell) -> boolean` and `canEditCell(cell) -> boolean` can also be provided as separate callbacks.
- If both `permissions` *and* `canReadCell`/`canEditCell` are provided, they are **ANDed** (all checks must allow).
  This is used by `bindCollabSessionToDocumentController` as defense-in-depth: it provides role/range metadata via
  `permissions`, while `session.canReadCell` / `session.canEditCell` also incorporate encryption invariants.

Note the naming difference:

- `CollabSession.setPermissions` uses `rangeRestrictions`
- `bindYjsToDocumentController` expects `restrictions`

#### Desktop permission wiring (JWT-derived; best-effort)

In desktop collaboration mode, the sync-server `token` may be either:

- an **opaque** shared token (dev), or
- a **JWT** (typical production), or
- an **opaque** token that must be introspected server-side (depending on deployment).

When the token is a JWT, the desktop app **decodes the JWT payload without verifying it** and derives *client-side* permissions + identity:

- `sub` → used (when present) as:
  - the `PresenceManager` user id (so presence ids match what the sync-server will enforce; the sync-server sanitizes awareness identities to the authenticated user id), and
  - forwarded as `userId` in `CollabSession.setPermissions(...)` (so `modifiedBy` metadata attribution is stable).
- `role` → forwarded as `role` in `CollabSession.setPermissions(...)` (defaults to `"editor"` if missing/invalid)
- `rangeRestrictions` → forwarded as `rangeRestrictions` in `CollabSession.setPermissions(...)` (defaults to `[]` if missing/invalid)

This decode is intentionally best-effort and does **not** replace server-side verification/authorization. The sync-server is the source of truth; the desktop decode exists so the UI can:

- mask unreadable cells immediately (before remote updates arrive), and
- proactively disable edits that the server would drop anyway.

Because the JWT payload is unverified and therefore untrusted, desktop treats `rangeRestrictions` defensively:
`CollabSession.setPermissions(...)` will throw if any restriction fails validation, and the desktop app will
fall back to dropping invalid restrictions (and, as a last resort, continuing with safe defaults) rather than
crashing on startup.

Sync-server note: in `jwt-hs256` auth mode, the server may be configured to require a non-empty `sub`
(`SYNC_SERVER_JWT_REQUIRE_SUB=1`, recommended). When `sub` is omitted and the server allows it, the
sync-server will treat the authenticated user id as the shared fallback `"jwt"`, which means presence ids
and `modifiedBy` attribution will not distinguish between collaborators.

Fallback for opaque / non-JWT tokens:

- If the token is missing or does not look like a JWT payload (or decoding fails), desktop treats the token as **opaque** and falls back to:
  - the locally chosen collab identity (for presence),
  - `{ role: "editor", rangeRestrictions: [] }` for `CollabSession.setPermissions(...)` (client-side gating becomes permissive; the sync-server still enforces, including any `rangeRestrictions` supplied via token introspection in `SYNC_SERVER_AUTH_MODE=introspect` deployments).
  - Note: the sync-server always sanitizes awareness identity fields to the **authenticated** user id. In shared-token auth mode (`SYNC_SERVER_AUTH_TOKEN`), that user id is the constant `"opaque"`, so presence **ids** are not stable per user (dev-only behavior). Clients may still show multiple cursors via distinct awareness clientIDs (and user display names/colors), but you will not get a canonical per-user id for attribution/access control. For stable presence/attribution, prefer JWT tokens with a real `sub` (or provide a stable user id out-of-band in introspect deployments).

Implementation reference: desktop JWT decode helpers live in [`apps/desktop/src/collab/jwt.ts`](../apps/desktop/src/collab/jwt.ts) (`tryDecodeJwtPayload`, `tryDeriveCollabSessionPermissionsFromJwtToken`).

#### Read-only UX behavior (viewer/commenter roles)

Roles apply both at the sync-server layer *and* in the desktop UX.

For `viewer` and `commenter` roles, the desktop app behaves as “read-only” for *shared* workbook mutations:

- **Cell edits are blocked**: the binder installs `DocumentController.canEditCell` guards (via `session.canEditCell(...)`) and will reject/revert disallowed deltas if they slip through (e.g. via programmatic calls).
- **Workbook-editing UI is disabled**: most desktop editing commands should consult `session.isReadOnly()` (or `session.getRole()`) and avoid initiating workbook mutations when the role cannot edit. `commenter` is still allowed to comment (see below).
- **Shared-state writes are suppressed (defense in depth)**: `bindCollabSessionToDocumentController` passes `canWriteSharedState: () => !session.isReadOnly()` into the binder, so sheet-level view/format deltas (freeze panes, row/col sizes, sheet/row/col format defaults, etc) are *not* written into Yjs for read-only roles. Some of these mutations may still apply locally in `DocumentController` (view/UI convenience), but they will not sync to other collaborators and can be overwritten by remote state.
- **Comments:**
  - `commenter` can add/edit replies and resolve threads (see `roleCanComment` in `@formula/collab-permissions`).
  - `viewer` can only read comments.

Note: the sync-server also enforces read-only roles by dropping incoming Yjs write messages for read-only connections, so client-side enforcement is primarily for UX consistency and to avoid “edit → immediate revert” loops.

Sync-server nuance:

- `viewer`: treated as fully read-only; SyncStep2/Update messages are dropped.
- `commenter`: treated as **comment-only**; sync-server allows Yjs updates that touch the `comments` root and rejects updates that touch other roots (cells/sheets/etc).

### Sheet schema (`sheets` array entries)

Each entry in `doc.getArray("sheets")` is a `Y.Map` with (at least):

```ts
type SheetViewState = {
  frozenRows: number;
  frozenCols: number;
  backgroundImageId?: string;
  /**
   * Per-column size overrides in **CSS pixels** (zoom-independent base sizes, i.e. `zoom = 1`).
   *
   * Note: this is a UI/view concern. The core engine stores column widths in Excel "character"
   * units (OOXML `col/@width`) for functions like `CELL("width")`.
   */
  colWidths?: Record<string, number>;
  /**
   * Per-row size overrides in **CSS pixels** (zoom-independent base sizes, i.e. `zoom = 1`).
   */
  rowHeights?: Record<string, number>;
};

// Optional layered formatting defaults (sheet/row/col).
// In Yjs these may be stored either:
// - as top-level keys on the sheet entry (`defaultFormat` / `rowFormats` / `colFormats`), or
// - nested inside `sheets[].view` in some BranchService-style snapshots.
type SheetFormatDefaults = {
  defaultFormat?: Record<string, any>;
  // Canonical (binder) encoding is a Y.Map keyed by string indices ("0", "1", ...),
  // but BranchService-style snapshots may store these as plain objects.
  rowFormats?: Y.Map<any> | Record<string, Record<string, any>>;
  colFormats?: Y.Map<any> | Record<string, Record<string, any>>;
  // Optional compressed formatting runs for large rectangular ranges.
  // Stored as a sparse per-column map.
  // In Yjs, this may be stored either as:
  // - a top-level key on the sheet entry (`formatRunsByCol`), or
  // - nested inside `sheets[].view` in some BranchService-style snapshots.
  //
  // Each run covers the half-open row interval `[startRow, endRowExclusive)`.
  formatRunsByCol?: Y.Map<any> | Record<
    string,
    Array<{
      startRow: number;
      endRowExclusive: number;
      format: Record<string, any>;
    }>
  >;
};

type Sheet = {
  id: string;
  name: string | null;
  view?: SheetViewState;
  // Additional per-sheet metadata tracked by `@formula/collab-workbook` + BranchService.
  visibility?: "visible" | "hidden" | "veryHidden";
  tabColor?: string | null; // 8-digit ARGB hex (e.g. "FFFF0000")
} & SheetFormatDefaults;
```

Implementation references:

- BranchService `SheetViewState` (superset, used for versioning/branching snapshots; includes optional layered formats): [`packages/versioning/branches/src/types.js`](../packages/versioning/branches/src/types.js)
- Desktop `DocumentController` `SheetViewState` (subset: frozen panes + row/col sizes): [`apps/desktop/src/document/documentController.js`](../apps/desktop/src/document/documentController.js)

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
- Desktop/Tauri note: the WebView CSP must allow WebSocket connections (`connect-src ws: wss:`) for the sync provider; see [`docs/11-desktop-shell.md`](./11-desktop-shell.md).

### Sync provider connection + “ready” signals

When you pass `connection: { wsUrl, docId, token }`, `@formula/collab-session` constructs a `y-websocket` `WebsocketProvider` internally and exposes it as `session.provider`.

Useful lifecycle helpers:

- `await session.whenLocalPersistenceLoaded()` — resolves when `options.persistence` has loaded any saved updates into the doc
- `await session.flushLocalPersistence()` — best-effort: forces any pending local persistence work to be durably written (useful before app teardown)
  - Optional: `await session.flushLocalPersistence({ compact: false })` skips compaction for a faster “durability snapshot” write (useful before a hard process exit)
- `await session.whenSynced()` — resolves when the sync provider reports `sync=true`

#### Observability APIs (sync status + local diagnostics)

`CollabSession` exposes a few lightweight observability hooks for **status bars** and **debug tooling**. These APIs are intentionally **best-effort**:

- Not all providers emit `"status"` / `"sync"` events.
- Some providers don’t update `.connected` / `.wsconnected` / `.synced` eagerly.
- These APIs should not be treated as protocol-level guarantees (they can be briefly stale).

APIs:

- `session.getSyncState(): { connected: boolean; synced: boolean }`
  - `connected`: best-effort derived from provider state + `"status"` events.
  - `synced`: best-effort “initial sync complete” signal (forced `false` when disconnected).
- `session.onStatusChange(cb): () => void`
  - Subscribes to `{ connected, synced }` transitions.
  - Does **not** call `cb` immediately; call `getSyncState()` for a snapshot.
- `session.getUpdateStats(): { lastUpdateBytes: number; maxRecentBytes: number; avgRecentBytes: number }`
  - Tracks the size (bytes) of recent **local-origin** Yjs updates (currently last 20).
  - Useful for diagnosing sync-server/websocket message size limits. This is the raw Yjs update payload size (not websocket framing/compression).
- `session.getLocalPersistenceState(): { enabled: boolean; loaded: boolean; lastFlushedAt: number | null }`
  - `loaded` stays `false` if persistence hydration fails.
  - `lastFlushedAt` is set by `flushLocalPersistence()` (best-effort).

Implementation: [`packages/collab/session/src/index.ts`](../packages/collab/session/src/index.ts) (`getSyncState`, `onStatusChange`, `getUpdateStats`, `getLocalPersistenceState` + constructor “Observability hooks”).

#### IndexedDB persistence flush + compaction (`IndexedDbCollabPersistence`)

`IndexedDbCollabPersistence` (implementation: [`packages/collab/persistence/src/indexeddb.ts`](../packages/collab/persistence/src/indexeddb.ts)) is a thin wrapper around `y-indexeddb`.

Two important operational details:

- **`flush(docId)` is implemented as a snapshot update (and compacts by default).** `y-indexeddb` persists incremental Yjs updates asynchronously and does not expose a reliable “await all pending writes” API. To satisfy Formula’s `CollabPersistence.flush` contract, `IndexedDbCollabPersistence.flush` writes a full-document snapshot (`Y.encodeStateAsUpdate(doc)`) into the IndexedDB `updates` object store so the in-memory document state at the time of the call can be recovered on the next load, even if some incremental writes are still in flight. By default, `flush()` calls `compact()` (so repeated flushes do not grow IndexedDB without bound). Implementation note: `IndexedDbCollabPersistence.flush` accepts an optional `{ compact?: boolean }` argument (defaults to `true`); pass `{ compact: false }` to append a snapshot record instead.
- **`compact(docId)` rewrites the update log to keep load time and disk usage bounded.** Without compaction, a long-lived document can accumulate a large number of incremental updates, which increases IndexedDB size and slows down `load()` (replay cost). Compaction replaces many small updates with a single snapshot update by reading the existing update records, merging them with `Y.encodeStateAsUpdate(doc)` (via `Y.mergeUpdates(...)`), clearing the `updates` store, and writing the merged update. This is important to avoid clobbering updates written by other tabs/processes (or in-flight writes) that may not have been applied to the current in-memory `Y.Doc`.
  - Knobs:
    - `new IndexedDbCollabPersistence({ maxUpdates })` (defaults to `500`) enables automatic background compaction once more than `maxUpdates` incremental updates have been observed (`0` disables auto-compaction).
    - `compactDebounceMs` controls how long compaction is debounced during bursts of edits (defaults to `250ms`).

`CollabSession.flushLocalPersistence()` delegates to the underlying persistence `flush(docId)` when present. Desktop typically calls it during teardown (or before closing a window) to reduce the chance of losing the last few edits.

Desktop note: the Tauri desktop app calls `flushLocalPersistence({ compact: false })` as part of the quit/restart flow (best-effort with a short timeout) to reduce the risk of losing recent offline/unsynced edits when the process hard-exits. The quit flow also waits briefly for DocumentController ↔︎ Yjs binder work to settle (including async encryption) before snapshotting.

If you use `persistence` or offline auto-connect gating, the initial WebSocket connection may be delayed until hydration completes (so offline edits are present before syncing).

Implementation: see `scheduleProviderConnectAfterHydration()` in [`packages/collab/session/src/index.ts`](../packages/collab/session/src/index.ts).

### Legacy offline option (`options.offline`, deprecated)

`@formula/collab-session` still accepts a legacy `options.offline` option for backwards compatibility.

This option is now **deprecated**. Prefer the unified `options.persistence` interface (`@formula/collab-persistence`) for all new code.

Current implementation detail (as of `packages/collab/session/src/index.ts`):

- `offline.mode: "indexeddb"` maps to `new IndexedDbCollabPersistence()`
- `offline.mode: "file"` maps to `new FileCollabPersistence(dir)` (and requires `offline.filePath`)
  - includes best-effort migration of historical on-disk logs written by the deprecated `@formula/collab-offline` file backend

For compatibility, passing `options.offline` will still expose `session.offline.{whenLoaded,clear,destroy}` as a convenience layer.

### Schema initialization (default sheet + root normalization)

By default, `createCollabSession` will keep the shared workbook schema well-formed by calling `ensureWorkbookSchema(session.doc, ...)` (from `@formula/collab-workbook`):

- ensures required roots exist (`cells`, `sheets`, `metadata`, `namedRanges`)
- ensures there is at least one sheet, creating a default sheet when appropriate
- prunes duplicate sheet ids created by concurrent initialization

To avoid creating “placeholder” sheets that later race with real hydrated state, schema initialization is *gated*:

- when a sync provider is present (e.g. y-websocket), schema init waits until the first `sync=true` event
- when local persistence (`options.persistence`) is enabled, schema init waits until local state has loaded (legacy `options.offline` follows the same gating)

Advanced control:

- disable entirely with `schema: { autoInit: false }`
- set default sheet id/name via `schema: { defaultSheetId, defaultSheetName }`

Implementation: see `ensureSchema` inside [`packages/collab/session/src/index.ts`](../packages/collab/session/src/index.ts).

### Collaborative undo scope (what gets undone)

When `createCollabSession({ undo: ... })` is enabled, the session creates a Yjs UndoManager-backed undo service (`session.undo`) that only tracks **local-origin** edits.

The default undo scope includes these shared roots:

- `cells`, `sheets`, `metadata`, `namedRanges`, and the `comments` root

You can include additional roots via:

- `undo.scopeNames` (creates maps by name)
- `undo.includeRoots(doc)` (include arbitrary Yjs root types)

Some internal roots are intentionally excluded from undo tracking (e.g. `cellStructuralOps`, the structural conflict monitor log), so conflict detection metadata is never undone.

Mixed-module Yjs note: in some Node/test environments, a `Y.Doc` may contain root types/items created by a different `yjs` module instance (ESM vs CJS). Upstream `Y.UndoManager` uses `instanceof` checks and can warn `[yjs#509] Not same Y.Doc` when scope types fail those checks. `@formula/collab-undo` includes best-effort patching of “foreign” constructors in the undo scope/transactions so collaborative undo works without warnings.

### Transactions + origins (local vs remote)

Yjs transactions have an optional `origin` value (`doc.transact(fn, origin)`), which Formula uses pervasively to distinguish **local** vs **remote** changes.

Practical guidance:

- For any feature that mutates shared state, prefer `session.transactLocal(() => { ... })`. It runs `fn` inside a local-origin transaction so:
  - collaborative undo (when enabled) records it as a local edit
  - conflict monitors can reliably classify it as local
- When applying remote updates manually in tests/tools, ensure they do *not* use the local origin. `@formula/collab-undo` exports a `REMOTE_ORIGIN` token for this purpose.

#### Bulk “time travel” origins (version restore / branch apply)

Some features intentionally perform **bulk rewrites** of workbook state (effectively “time travel”):

- **Version restore** (`@formula/collab-versioning` / `createYjsSpreadsheetDocAdapter.applyState`) uses the origin string `"versioning-restore"`.
- **Branch checkout/merge apply** (`CollabBranchingWorkflow` → `applyDocumentStateToYjsDoc`) uses the origin string `"branching-apply"` by default
  (can be configured to use `session.origin` for undoable behavior).

These origins are **not** “local user edits” and should be treated specially:

- Conflict monitors (`FormulaConflictMonitor`, `CellConflictMonitor`, `CellStructuralConflictMonitor`) accept an `ignoredOrigins` set to completely ignore these transactions
  (no conflicts emitted and no local-edit tracking updates). `createCollabSession` configures this by default.
- Structural conflict monitoring (`CellStructuralConflictMonitor`) only logs operations for origins in its `localOrigins` set; bulk apply origins should not be included.

Example:

```ts
import * as Y from "yjs";
import { REMOTE_ORIGIN } from "@formula/collab-undo";

Y.applyUpdate(session.doc, remoteUpdateBytes, REMOTE_ORIGIN);
```

Implementation references:

- `CollabSession.transactLocal`: [`packages/collab/session/src/index.ts`](../packages/collab/session/src/index.ts)
- `REMOTE_ORIGIN`: [`packages/collab/undo/src/yjs-undo-service.js`](../packages/collab/undo/src/yjs-undo-service.js)

### Batch cell writes (`CollabSession.setCells`)

`CollabSession.setCells(...)` applies multiple cell updates as a **single local transaction**, with semantics aligned to `setCellValue` / `setCellFormula` (permissions, encryption, and conflict monitors).

Signature:

```ts
await session.setCells(
  [{ cellKey: "Sheet1:0:0", value: 123 }, { cellKey: "Sheet1:0:1", formula: "=A1*2" }],
  { ignorePermissions?: boolean },
);
```

Semantics (implementation-backed):

- **Formula wins**: when both `value` and `formula` are provided, the update is treated as a formula write (`value` is ignored). `formula: null` (or empty/whitespace) clears the formula.
- **Atomic / all-or-nothing**:
  - The session validates the entire batch before applying any cell-content writes.
  - All cell mutations are applied inside a single `session.transactLocal(...)` after a preflight check, so the batch does not partially apply (permission denied, missing encryption key, plaintext-to-encrypted violation).
- **Permissions**:
  - If session permissions are configured via `session.setPermissions(...)`, the batch is rejected if *any* target cell is not editable.
  - `{ ignorePermissions: true }` bypasses permission checks (intended for internal tooling/migrations), but does **not** bypass encryption invariants.
- **Encryption invariants**:
  - If a cell is already encrypted (or `shouldEncryptCell` says it should be), the write is stored as `enc` (never plaintext).
  - Plaintext writes are rejected for encrypted cells (checked for the whole batch before applying any update).
  - When `encryption.encryptFormat = true`, the session attempts to preserve existing per-cell formatting by decrypting the previous payload when possible, otherwise falling back to plaintext `format` / legacy `style` stored under existing key aliases, and then removing plaintext `format`/`style` keys when writing the encrypted payload.
- **Conflict monitors**: when enabled, `setCells` delegates to the monitors (`FormulaConflictMonitor` / `CellConflictMonitor`) so null-marker semantics (`formula = null`) stay consistent.

Implementation: [`packages/collab/session/src/index.ts`](../packages/collab/session/src/index.ts) (`CollabSession.setCells`).

### Cell-map compaction (`CollabSession.compactCells`)

Some writers preserve “empty” cell entries in the `cells` root (including marker-only clears used for conflict detection). Over time this can bloat the `cells` map and increase snapshot/update size.

`CollabSession.compactCells(...)` is an **opt-in** maintenance API that deletes prunable entries from `doc.getMap("cells")` while preserving encryption and formatting invariants.

Signature:

```ts
const { scanned, deleted } = session.compactCells({
  dryRun: true,
  maxCellsScanned: 50_000,
  // pruneMarkerOnly: true | false,
});
```

Options (implementation-backed):

- `dryRun?: boolean` — when true, returns counts without mutating the doc.
- `maxCellsScanned?: number` — bounds work; defaults to `Infinity`. If non-finite or `<= 0`, returns `{ scanned: 0, deleted: 0 }`.
- `pruneMarkerOnly?: boolean` — controls whether marker-only entries (`formula = null` clears) are eligible for deletion:
  - **Default:** `false` when *any* conflict monitor is enabled (`formulaConflicts`, `cellConflicts`, or `cellValueConflicts`), otherwise `true`.
- `origin?: unknown` — Yjs transaction origin for the delete transaction (defaults to `"cells-compact"`).

Preservation rules:

- **Never prunes encrypted cells** (`enc` present).
- **Never prunes format-only cells** or explicit format clears (`format` or legacy `style` present).
- Only prunes entries that contain no keys other than `value` / `formula` / `modified` / `modifiedBy`.

When it’s safe to run:

- Safe to run periodically in docs where you are **not relying on marker-only causality** (e.g. conflict monitoring disabled), or when you explicitly set `pruneMarkerOnly: false` to preserve clear markers.

Implementation: [`packages/collab/session/src/index.ts`](../packages/collab/session/src/index.ts) (`CollabSession.compactCells`).

---

## Binding Yjs to the desktop workbook model

**Goal:** keep the desktop workbook state machine (`DocumentController`) in sync with the shared Yjs workbook.

The binder lives at:

- [`packages/collab/binder/index.js`](../packages/collab/binder/index.js) (`bindYjsToDocumentController`)

It synchronizes:

- **Cell contents** (`value` / `formula` / `format`):
  - `Y.Doc` → `cells` root (`Y.Map`)
  - Desktop `DocumentController` (see [`apps/desktop/src/document/documentController.js`](../apps/desktop/src/document/documentController.js))
- **Sheet view state** (frozen panes + row/col sizes):
  - `Y.Doc` → `sheets[].view`
  - Desktop `DocumentController` sheet view state (`applyExternalSheetViewDeltas` / `sheetViewDeltas`)
- **Layered formatting defaults** (sheet/row/col formats):
  - `Y.Doc` → `sheets[].defaultFormat`, `sheets[].rowFormats`, `sheets[].colFormats` (with legacy fallback from `sheets[].view`)
  - Desktop `DocumentController` format state (`applyExternalFormatDeltas` / `formatDeltas`)
- **Range-run formatting** (compressed rectangular formats):
  - `Y.Doc` → `sheets[].formatRunsByCol` (with legacy fallback from `sheets[].view`)
  - Desktop `DocumentController` range-run state (`applyExternalRangeRunDeltas` / `rangeRunDeltas`)

> Note: the binder syncs **cell contents** (`cells`) *and* **sheet view state**
> (`sheets[].view`), but it does **not** implement full sheet list semantics
> (create/delete/rename/reorder) or other sheet metadata syncing (e.g. `visibility`, `tabColor`). If the desktop UI needs live sheet list syncing,
> it should observe the Yjs `sheets` array directly (or use a dedicated binder).

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

  // Optional: opt into write semantics needed for conflict monitors (preserve `formula=null`
  // markers and empty cell maps on clears so delete-vs-overwrite causality is detectable).
  // formulaConflictsMode: "formula+value",
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
  // Optional: also clear per-cell formatting for masked cells.
  // maskCellFormat: true,
  // Optional: opt into conflict-monitor-compatible binder writes.
  // formulaConflictsMode: "formula+value",

  // Optional: override how DocumentController-driven writes are transacted into Yjs.
  //
  // By default, the helper uses `session.transactLocal(...)` (so edits use
  // `session.origin` and participate in collaborative undo/conflict detection when enabled).
  //
  // To opt out and use the binder’s internal origin token instead, pass:
  // undoService: null,
});
```

`bindCollabSessionToDocumentController` passes `maskCellFormat` and `formulaConflictsMode` through to the underlying binder.

Implementation:

- Binder option: [`packages/collab/binder/index.js`](../packages/collab/binder/index.js) (`bindYjsToDocumentController` → `maskCellFormat`)
- Helper pass-through: [`packages/collab/session/src/index.ts`](../packages/collab/session/src/index.ts) (`bindCollabSessionToDocumentController`)

### Echo suppression (feedback-loop prevention)

When you bind two reactive systems, you must avoid doing this:

1. Local edit (DocumentController) writes to Yjs
2. Yjs observer fires (because the doc changed)
3. Binder applies the same change back into DocumentController (wasted work + can cause feedback loops)

The binder prevents feedback loops using internal guards (and a small amount of `payload.source` filtering),
not by filtering `transaction.origin`:

- While applying a local **DocumentController → Yjs** write, the binder sets an internal `applyingLocal` flag.
  The Yjs `observeDeep` handlers early-return when `applyingLocal` is true, so the deep observer event that
  immediately follows the binder’s own write is ignored.
- While applying a **Yjs → DocumentController** update (via `applyExternalDeltas` / `applyExternalSheetViewDeltas`
  / `applyExternalFormatDeltas` / `applyExternalRangeRunDeltas`), the binder sets `applyingRemote`.
  The DocumentController `"change"` handler ignores changes while `applyingRemote` is true, so remote-applied
  deltas are not written back into Yjs.
- In desktop, multiple binders can be attached to the same `DocumentController` (for example, the lightweight
  sheet view binder used by `SpreadsheetApp` alongside the full binder). In that case, one binder may apply a
  remote update into the `DocumentController` while the other binder’s `applyingRemote` flag is **false**.
  To avoid echoing those **external** changes back into Yjs (and polluting collaborative undo with “remote”
  edits), binders treat these sources as external and ignore them in their `"change"` handlers:
  - `payload.source === "collab"` (remote Yjs → DocumentController apply)
  - `payload.source === "applyState"` (snapshot restore / version history hydration)

This is intentionally **not** implemented as “ignore all local origins”, because other parts of the stack
often reuse the same origin token for programmatic mutations (e.g. branch checkout/merge apply, version
restore, or direct `session.setCell*` calls) that *must* update the desktop projection.

Origins still matter for **undo/conflict semantics**, not echo suppression:

- If you pass `undoService: session.undo` (or use `bindCollabSessionToDocumentController`, which defaults to
  `session.transactLocal(...)`), DocumentController-driven writes use `session.origin` and participate in
  collaborative undo + conflict monitoring.
- If you pass `undoService: null`, the binder uses its own `binderOrigin` for DocumentController-driven writes,
  so those edits will generally be treated as **non-local** by collaborative undo/conflict monitors.

### Undo/redo semantics in collaborative mode

The desktop `DocumentController` maintains its own local history stack, but in collaborative mode it is **not** the canonical user-facing undo stack.

In a shared Yjs session you generally want undo/redo to:

- only revert the **local user’s** edits
- never undo remote collaborators’ changes

That behavior is provided by Yjs’ `UndoManager` (via `@formula/collab-undo`, exposed as `session.undo` when `createCollabSession({ undo: ... })` is enabled).

See: [`docs/adr/ADR-0004-collab-sheet-view-and-undo.md`](./adr/ADR-0004-collab-sheet-view-and-undo.md)

---

## Sheet view state storage and syncing

Per-sheet view state (frozen panes + row/col size overrides, plus additional shared sheet-level UI metadata like merged ranges, drawings, and background images) is stored on each sheet entry in the `sheets` array:

- `doc.getArray("sheets").get(i).get("view")`

### `sheets[].view` encoding: plain objects vs `Y.Map`

For compatibility with BranchService snapshots and historical clients, `sheets[].view` may be stored as either:

- a **plain JSON object** (most common in snapshots / branching/versioning adapters), or
- a **`Y.Map`** (and nested keys like `view.colWidths` / `view.rowHeights` may also be `Y.Map`s).

In mixed-module environments (ESM + CJS), a doc may also contain **duck-typed / “foreign” `Y.Map`** instances
whose constructors do not match the app’s `yjs` import (e.g. after applying provider updates created by a
different `yjs` module instance).

Implications:

- Do not assume `sheet.get("view")` is a plain object; prefer `@formula/collab-yjs-utils` (`getYMap`, `yjsValueToJson`)
  or the binder/session helpers that already handle cross-instance types.
- When `view` is stored as a `Y.Map`, desktop binders update it **in-place** to avoid rewriting large unknown keys
  (e.g. `view.drawings`) on small changes like freeze panes or axis resizing.

The `view` object is BranchService-compatible (some snapshots may include
additional keys like layered formatting defaults), but the desktop
binder/`DocumentController` currently consume the subset of fields related to
frozen panes + row/col size overrides (and, in desktop, merged ranges + drawings + background image metadata):

```ts
{
  frozenRows: 2,
  frozenCols: 1,
  backgroundImageId: "img-bg-1",
  colWidths: { "0": 120 },
  rowHeights: { "1": 40 },
  // Optional shared metadata:
  mergedRanges: [{ startRow: 0, endRow: 1, startCol: 0, endCol: 2 }],
  drawings: [{ id: "drawing-1", zOrder: 0, kind: { type: "image", imageId: "img-1" }, anchor: { /* ... */ } }]
}
```

`colWidths` / `rowHeights` values are **base CSS pixels** (zoom-independent, i.e. the size at `zoom = 1`).

### Drawing IDs (namespaces + safety)

Drawing entries in `view.drawings` (and the desktop `DocumentController` sheet view state) are stored as **JSON-serializable objects**. For compatibility with historical snapshots, `drawing.id` is allowed to be either:

- a **string** (preferred for document snapshots / Yjs payloads), or
- a **number** (must be a JS safe integer).

The UI drawing overlay requires a stable numeric key (`DrawingObject.id: number`). Adapters normalize ids as follows:

- **Positive safe integers** (`> 0`) are passed through unchanged.
  - For string ids, this only applies to **canonical base-10 integer strings** (e.g. `"42"`). Non-canonical numeric strings like `"001"`, `"1e3"`, or `"+1"` are treated as opaque ids and hashed to avoid collisions.
- Any other id (missing, non-numeric, non-positive, or unsafe integer) is mapped into a **reserved hashed namespace**: `id <= -2^33`.
  - String ids are **trimmed** before hashing to match `DocumentController` normalization.
  - Very long string ids are hashed from a bounded summary (prefix/middle/suffix + length) to avoid large allocations.
  - Additionally, for collab safety, most consumers treat **string ids longer than 4096 chars as invalid** and will ignore/drop those drawing entries when normalizing sheet view state (desktop `DocumentController`, desktop sheet-view binder, full collab binder upgrade path, workbook schema normalization, and BranchService snapshot normalization).

In addition, ChartStore “canvas charts” are rendered as drawing objects with ids in a separate negative namespace:

- **Canvas chart ids**: `-2^32 <= id < 0` (see `chartIdToDrawingId` / `isChartStoreDrawingId`).

Important: do **not** treat “negative id” as synonymous with “chart”. Workbook drawings can also have negative ids (hashed ids).

Implementation references:

- UI id generation for new drawings: `apps/desktop/src/drawings/types.ts` (`createDrawingObjectId`, random 53-bit safe integer)
- Snapshot id parsing + hashing: `apps/desktop/src/drawings/modelAdapters.ts` (`parseDrawingObjectId`)
- Canvas chart id mapping: `apps/desktop/src/charts/chartDrawingAdapter.ts` (`chartIdToDrawingId`, `isChartStoreDrawingId`)

XLSX/export note: DrawingML’s `<xdr:cNvPr id="...">` is an `xsd:unsignedInt` (u32). UI-layer ids may exceed u32, so any future “export UI-created drawings to XLSX” pipeline must **remap ids at write time** instead of writing the UI id verbatim.

Compatibility note:

- Some historical/experimental docs stored `frozenRows` / `frozenCols` as **top-level** fields directly on the sheet map.
- The desktop sheet view binder also mirrors `drawings` and `backgroundImageId` to top-level `sheet.drawings` / `sheet.backgroundImageId` fields for backwards compatibility with older clients that do not read `view.*`.
- BranchService’s Yjs adapter explicitly supports both shapes (see `branchStateFromYjsDoc` in [`packages/versioning/branches/src/yjs/branchStateAdapter.js`](../packages/versioning/branches/src/yjs/branchStateAdapter.js)), but **new writes should use `view`**.

Because `sheets` is part of the shared Y.Doc, any edits to `view` will sync like any other Yjs update (subject to your provider/persistence).

### Desktop binder integration (`bindYjsToDocumentController`)

When the desktop binder is in use, `bindYjsToDocumentController` also keeps the
desktop projection in sync with the shared view state:

- **Yjs → desktop:** observes `sheets` changes and applies them into
  `DocumentController` via `applyExternalSheetViewDeltas` (or best-effort via
  `setFrozen`/`setColWidth`/`setRowHeight` for older controllers).
- **Desktop → Yjs:** listens for `sheetViewDeltas` emitted by `DocumentController`
  and writes normalized state back into `sheets[i].view`.

Desktop also has a lightweight sheet-view-only binder used by `SpreadsheetApp`:
[`apps/desktop/src/collab/sheetViewBinder.ts`](../apps/desktop/src/collab/sheetViewBinder.ts) (`bindSheetViewToCollabSession`).
It performs the same high-level synchronization (Yjs `sheets` ↔ `DocumentController` frozen panes + row/col size overrides, plus merged ranges + drawings metadata),
and also mirrors legacy top-level `frozenRows`/`frozenCols` fields for backwards compatibility. Like the main binder, it
suppresses writes for read-only roles (`viewer`/`commenter`).

In addition, the binder synchronizes layered formatting defaults (sheet/row/col styles):

- **Yjs → desktop:** reads `defaultFormat` / `rowFormats` / `colFormats` from sheet metadata
  (preferring top-level keys, with fallback to `sheets[].view.*` for legacy snapshots) and applies
  them via `DocumentController.applyExternalFormatDeltas` when available.
- **Desktop → Yjs:** listens for `formatDeltas` emitted by `DocumentController` and writes them
  into the sheet entry using sparse encodings:
  - `defaultFormat` (style object)
  - `rowFormats` / `colFormats` (`Y.Map<string, styleObject>` keyed by string indices)

Branching/merge snapshot note:

- BranchService’s Yjs adapter (`branchStateFromYjsDoc`) reads sheet-formatting metadata
  (layered defaults + range-run formats) from the canonical top-level sheet keys:
  `defaultFormat` / `rowFormats` / `colFormats` / `formatRunsByCol`
  (with fallback to legacy encodings embedded in `sheets[].view`).
- When applying a snapshot back into Yjs (`applyBranchStateToYjsDoc` / `applyDocumentStateToYjsDoc`), BranchService updates
  those top-level keys so desktop binders that prefer them don’t see stale formatting after checkout/merge.

Semantics note:

- `sheets[].view` is **shared workbook state** (similar to formatting), not per-user transient UI state.
- Read-only roles (`viewer`/`commenter`) may still allow local “view tweaks” in the desktop UI, but those
  changes are not persisted into Yjs (so they do not sync and can be overwritten by remote state).
- Per-user ephemeral state (cursor/selection/active sheet/viewport) should use Awareness/presence.

See: [`docs/adr/ADR-0004-collab-sheet-view-and-undo.md`](./adr/ADR-0004-collab-sheet-view-and-undo.md)

---

## Managing shared workbook metadata (`@formula/collab-workbook`)

Edits that aren’t simple cell value/formula writes (e.g. sheet creation/rename/order, named ranges, workbook metadata) should still be treated as **shared document state** and written into the same Yjs doc.

`@formula/collab-workbook` provides small helper managers that encapsulate common mutations and (when you use `createSheetManagerForSession(session)` / `createMetadataManagerForSession(session)` / `createNamedRangeManagerForSession(session)`) run them inside `session.transactLocal(...)` so they:

- sync to other collaborators like any other Yjs change
- participate correctly in collaborative undo (when enabled)

Example:

```ts
import {
  createSheetManagerForSession,
  createNamedRangeManagerForSession,
  createMetadataManagerForSession,
} from "@formula/collab-workbook";

const sheets = createSheetManagerForSession(session);
const namedRanges = createNamedRangeManagerForSession(session);
const metadata = createMetadataManagerForSession(session);

// Sheets
sheets.addSheet({ id: "Sheet2", name: "Plan", index: 1 });
sheets.renameSheet("Sheet2", "Plan v2");
sheets.moveSheet("Sheet2", 0);
// Optional sheet metadata (Excel-like semantics)
sheets.setVisibility("Sheet2", "hidden");
sheets.setTabColor("Sheet2", "FFFF0000");

// Metadata / named ranges
metadata.set("locale", "en-US");
namedRanges.set("Revenue", { sheetId: "Sheet1", startRow: 0, startCol: 0, endRow: 9, endCol: 0 });
```

Note: the desktop `bindYjsToDocumentController` binder syncs **cells + sheet
view state**, but does not project sheet list changes (add/delete/rename/order)
or other per-sheet metadata changes (visibility/tabColor)
into `DocumentController`. If the desktop UI needs live sheet list syncing, it
should observe the Yjs `sheets` array directly (or via a dedicated binder).

---

## Presence integration (Awareness → desktop renderer)

### PresenceManager

`@formula/collab-session` can construct a `PresenceManager` automatically when `presence: {...}` is provided.

- `session.presence` is a `PresenceManager` (or `null` if not enabled)
- `PresenceManager` writes/reads the Yjs Awareness state field `"presence"`
- `PresenceManager.subscribe()` calls its listener immediately with the current **remote** presences on the active sheet, and then on any **remote** awareness update (it intentionally ignores local cursor/selection updates to reduce re-render pressure). Pass `{ includeOtherSheets: true }` to include remote users on non-active sheets.

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
  // In collaborative mode, the desktop UI uses sheet-qualified refs like `Sheet1!A1`.
  // (In non-collab mode some legacy docs use unqualified "A1".)
  cellRef: "Sheet1!A1",
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

Permissions note: `CommentManager` supports an optional `canComment()` guard (and throws on comment mutations when it returns `false`). When you construct a comment manager from a `CollabSession` via `createCommentManagerForSession(session)`, it will use `session.canComment()` so viewer roles remain read-only for comments.

Desktop note: desktop collaboration uses a *binder-origin* undo scope (DocumentController→Yjs),
so comment edits must run inside the binder-origin transact wrapper (not `session.transactLocal`).
Use `createCommentManagerForDoc({ doc: session.doc, transact: undoService.transact })` (see
`apps/desktop/src/collab/documentControllerCollabUndo.ts` / `apps/desktop/src/app/spreadsheetApp.ts`).
If you also want permission enforcement at the comment API layer, pass `canComment: () => session.canComment()` so viewer roles throw on comment mutations (fail-closed).

If you need to normalize legacy Array-backed docs to the canonical Map schema:

- `migrateCommentsArrayToMap(doc)` (see [`packages/collab/comments/src/manager.ts`](../packages/collab/comments/src/manager.ts))

### Automatic legacy schema migration (`comments.migrateLegacyArrayToMap`)

If you construct a `CollabSession` on unknown/historical documents, you can opt into an automatic, best-effort migration of legacy Array-backed comment roots:

```ts
import { createCollabSession } from "@formula/collab-session";

const session = createCollabSession({
  connection: { wsUrl, docId, token },
  comments: { migrateLegacyArrayToMap: true },
});
```

What it migrates:

- Legacy schema: `comments` root stored as `Y.Array<Y.Map>` (a list of comment maps).
- Canonical schema: `comments` root stored as `Y.Map<string, Y.Map>` keyed by comment id.

Migration behavior (implementation-backed):

- Runs **after initial hydration**, queued in a microtask:
  - after local persistence load completes (if `options.persistence` / legacy `options.offline` is enabled)
  - after the provider reports initial `sync=true` (if a provider is enabled)
  - if both local persistence *and* a provider are enabled, it waits for **both** to complete so migration does not run against a partially hydrated document.
- Gated by comment permissions: migration is skipped when `session.canComment()` is false (e.g. viewer role) so read-only clients do not generate Yjs updates. If permissions are later updated to allow comments, migration is retried.
- Uses `migrateCommentsArrayToMap(doc, { origin: "comments-migrate" })`, which:
  - renames the legacy root to `comments_legacy*` (so old content remains accessible), and
  - creates a canonical Map-backed `comments` root and copies entries keyed by comment id.
- Preserves collaborative undo scope: because migration replaces the `comments` root object, `CollabSession` re-adds the new root to all known `UndoManager` scopes (best-effort) so subsequent comment edits remain undoable.
- Never blocks session startup: failures are swallowed (best-effort).

Desktop: the desktop app opts into this by default (see [`apps/desktop/src/app/spreadsheetApp.ts`](../apps/desktop/src/app/spreadsheetApp.ts)).

Implementation:

- Session scheduling + undo-scope repair: [`packages/collab/session/src/index.ts`](../packages/collab/session/src/index.ts) (`scheduleCommentsMigration`)
- Migration transform: [`packages/collab/comments/src/manager.ts`](../packages/collab/comments/src/manager.ts) (`migrateCommentsArrayToMap`)

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
    // Optional: bound growth of the shared op log in long-lived docs.
    // maxOpRecordsPerUser: 2000,
    // maxOpRecordAgeMs: 7 * 24 * 60 * 60 * 1000,
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
- For deterministic delete-vs-overwrite detection, formula clears must be represented as `formula = null`
  (not `cell.delete("formula")`) because Yjs map deletes do not create Items.
- `remoteUserId` attribution is best-effort and may be empty if the overwriting writer did not update `modifiedBy`.
- `cellConflicts` writes causal metadata into a shared `cellStructuralOps` log. Per-user history is bounded via
  `maxOpRecordsPerUser`, and you can additionally enable age-based pruning (`maxOpRecordAgeMs`) to avoid
  unbounded growth in docs with many distinct user ids over time (best-effort).
  - Age pruning is conservative: newly-arriving records are not deleted in the same op-log transaction they are added,
    so late-arriving/offline records can still be ingested and compared before cleanup.
  - Age pruning is incremental: very large logs may take multiple passes to fully prune.
  - Age pruning is additionally conservative relative to the local op log queue: records are only pruned when they are older
    than both the age cutoff and the oldest local op record (queue head), which avoids pruning history that may still be needed
    to compare against long-offline local ops.
- Implementation: [`packages/collab/conflicts/src/formula-conflict-monitor.js`](../packages/collab/conflicts/src/formula-conflict-monitor.js)
- Conflict monitors support an `ignoredOrigins` option to ignore bulk “time travel” transactions such as version restores
  (`"versioning-restore"`) and branch apply operations (`"branching-apply"`). `createCollabSession` wires this by default.

---

## Version history (collab mode)

Use `@formula/collab-versioning` to store and restore workbook snapshots in a collaborative session.

Source: [`packages/collab/versioning/src/index.ts`](../packages/collab/versioning/src/index.ts)

> **Deployment note (sync-server reserved roots):**
> `services/sync-server` includes a **reserved root mutation guard** that can reject Yjs updates touching
> `versions`, `versionsMeta`, or `branching:*`.
>
> - Env var: `SYNC_SERVER_RESERVED_ROOT_GUARD_ENABLED` (defaults to **enabled** when `NODE_ENV=production`, **disabled** otherwise)
> - Optional env vars:
>   - `SYNC_SERVER_RESERVED_ROOT_NAMES` (comma-separated; defaults to `versions,versionsMeta`)
>   - `SYNC_SERVER_RESERVED_ROOT_PREFIXES` (comma-separated; defaults to `branching:`)
> - Implementation: [`services/sync-server/src/server.ts`](../services/sync-server/src/server.ts) + [`services/sync-server/src/ywsSecurity.ts`](../services/sync-server/src/ywsSecurity.ts)
>
> If this guard is enabled, in-doc stores like `YjsVersionStore` and `YjsBranchStore` will not work unless you disable it (or use non-Yjs stores).
> If you customize the branching root name (non-default `rootName`), note that the guard only blocks configured prefixes; you may need to extend `SYNC_SERVER_RESERVED_ROOT_PREFIXES` accordingly.

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

### Workbook diffs (Version History “Compare”)

The desktop Version History “Compare” UI computes a workbook-level diff between a selected version snapshot and the current live Yjs document state via `diffYjsWorkbookVersionAgainstCurrent`:

- Source: [`packages/versioning/src/yjs/versionHistory.js`](../packages/versioning/src/yjs/versionHistory.js)
- Diff implementation: [`packages/versioning/src/yjs/diffWorkbookSnapshots.js`](../packages/versioning/src/yjs/diffWorkbookSnapshots.js)

The returned `WorkbookDiff` includes sheet-level metadata changes in `diff.sheets.metaChanged[]` (in addition to adds/removes/renames/reorders). This list is intentionally limited to a **small subset** of per-sheet state so version history summaries stay fast and compact:

- `visibility` (`visible | hidden | veryHidden`)
- `tabColor`: canonicalized to **8-digit uppercase ARGB** hex (e.g. `FF00FF00`) or `null` when cleared
- frozen panes: `view.frozenRows`, `view.frozenCols`
- background image: `view.backgroundImageId` (optional; only included when non-empty)

Note: large sheet view maps (e.g. row/column size tables) are intentionally excluded from workbook diffs; only frozen pane counts are tracked.

### VersionStore choices (in-doc vs API vs SQLite)

`createCollabVersioning` accepts a `store` option (see defaulting logic in
[`packages/collab/versioning/src/index.ts`](../packages/collab/versioning/src/index.ts)).
The choice affects how version history is persisted and whether it flows through Yjs/sync-server.

- **`YjsVersionStore`** (in-doc, default)
  - Stores history inside the live collaborative `Y.Doc` under the `versions` and `versionsMeta` roots.
  - Version history **syncs via Yjs** (y-websocket/sync-server) and is included in sync-server persistence.
  - Requires the sync-server reserved root guard to be **disabled**, otherwise writes will be rejected.
  - Implementation: [`packages/versioning/src/store/yjsVersionStore.js`](../packages/versioning/src/store/yjsVersionStore.js)
- **`ApiVersionStore`** (cloud DB / Formula API)
  - Stores versions in the API-backed `document_versions` table (out-of-doc).
  - Does **not** mutate `versions*` roots in the shared Yjs document, so it is compatible with the sync-server reserved root guard.
  - Implementation: [`packages/versioning/src/store/apiVersionStore.js`](../packages/versioning/src/store/apiVersionStore.js)
- **`SQLiteVersionStore`** (local desktop / Node)
  - Stores versions in a local SQLite file (out-of-doc).
  - Does **not** mutate `versions*` roots in the shared Yjs document, so it is compatible with the sync-server reserved root guard.
  - Implementation: [`packages/versioning/src/store/sqliteVersionStore.js`](../packages/versioning/src/store/sqliteVersionStore.js)

Note: there are additional stores (e.g. `IndexedDBVersionStore`) under `packages/versioning/src/store/`, but the key deployment split is **in-doc (Yjs)** vs **out-of-doc (API/SQLite)**.

#### Desktop wiring note

The Formula desktop app renders the Version History panel via `createPanelBodyRenderer` (see `apps/desktop/src/panels/panelBodyRenderer.tsx`).
To keep version history usable when `SYNC_SERVER_RESERVED_ROOT_GUARD_ENABLED` is on, the desktop layer can inject an out-of-doc store by providing a factory on `PanelBodyRendererOptions`:

- `createCollabVersioningStore?: (session) => VersionStore` (alias: `createVersionStore`)

### How history is stored (YjsVersionStore)

By default, `createCollabVersioning` uses `YjsVersionStore`, which stores all versions **inside the same shared Y.Doc** under these roots:

- `versions` (records)
- `versionsMeta` (ordering + metadata)

#### Snapshot size + sync-server message limits

Large workbooks can produce snapshot blobs that exceed the sync-server websocket message limit (`SYNC_SERVER_MAX_MESSAGE_BYTES`, default **2MB**). If a single Yjs update exceeds this limit, the server will close the websocket with code **1009** ("Message too big").

To avoid these catastrophic failures, the default `createCollabVersioning` store configuration writes snapshots in **streaming mode** (multiple Yjs transactions/updates):

- `writeMode: "stream"`
- `chunkSize: 64KiB`
- `maxChunksPerTransaction`: bounded so each update stays comfortably below typical message limits (at the default chunk size this is 8 chunks, ~512KiB of snapshot payload/update).

Tradeoff: streaming writes create more Yjs transactions (more websocket messages), but keeps each message small and robust for large documents.

If you need to tune this (for example, when running with a very small `SYNC_SERVER_MAX_MESSAGE_BYTES`), you can override the default store settings without constructing a store manually:

```ts
import { createCollabVersioning } from "@formula/collab-versioning";

const versioning = createCollabVersioning({
  session,
  yjsStoreOptions: {
    // Example: keep each streamed update under ~128KiB.
    chunkSize: 32 * 1024,
    maxChunksPerTransaction: 2,
  },
});
```

### Snapshot/restore isolation (excluded roots)

`CollabVersioning` snapshots/restores are intended to affect **user workbook state**, not internal collaboration metadata stored in the same `Y.Doc`.

Implementation: see `excludeRoots` in [`packages/collab/versioning/src/index.ts`](../packages/collab/versioning/src/index.ts).

CollabVersioning excludes:

- **Always excluded** internal collaboration roots:
  - `cellStructuralOps` (structural conflict op log)
  - default branching graph roots (only when `rootName` is the default `"branching"`):
    - `branching:branches`
    - `branching:commits`
    - `branching:meta`
- **Always excluded** internal versioning roots:
  - `versions`
  - `versionsMeta`

Important details:

- When the store is `YjsVersionStore`, excluding `versions` and `versionsMeta` prevents:
  - recursive snapshots (history containing itself)
  - restores rolling back version history
- Even when the store is out-of-doc (API/SQLite/etc), a document may still contain these reserved roots
  (e.g. from earlier in-doc usage or dev sessions). Restoring must never attempt to rewrite them, both
  to avoid rewinding internal state and to prevent server-side “reserved root mutation” disconnects.

Note: `CollabBranchingWorkflow` supports customizing the branch graph `rootName` (default `"branching"`). If you use a non-default root name, you must also tell `CollabVersioning` to exclude those internal roots; otherwise version snapshots/restores can accidentally include (and later rewind) branch history.

Use `excludeRoots` to extend the built-in exclusions:

```ts
import { createCollabVersioning } from "@formula/collab-versioning";

const rootName = "myBranching";

const versioning = createCollabVersioning({
  session,
  excludeRoots: [`${rootName}:branches`, `${rootName}:commits`, `${rootName}:meta`],
});
```

---

## Branching + merging (collab mode)

Formula supports Git-like **branch + merge** workflows for collaborative spreadsheets by storing the branch/commit graph **inside the shared Y.Doc**.

> **Deployment note (sync-server reserved roots):**
> `YjsBranchStore` writes to `branching:*` roots. When the sync-server reserved root mutation guard is enabled
> (see [`services/sync-server/src/server.ts`](../services/sync-server/src/server.ts) / [`services/sync-server/src/ywsSecurity.ts`](../services/sync-server/src/ywsSecurity.ts)),
> those updates will be rejected unless you disable `SYNC_SERVER_RESERVED_ROOT_GUARD_ENABLED` (or use a non-Yjs store such as `SQLiteBranchStore`).

> **Desktop wiring note:**
> The Formula desktop app can inject a non-Yjs branch store into the Branch Manager panel via `PanelBodyRendererOptions.createCollabBranchStore` (alias: `createBranchStore`) in `apps/desktop/src/panels/panelBodyRenderer.tsx`.

### Required Yjs roots (default `rootName = "branching"`)

The Yjs-backed branch store reserves:

- `branching:branches` (`Y.Map<branchName, Y.Map>`)
- `branching:commits` (`Y.Map<commitId, Y.Map>`)
- `branching:meta` (`Y.Map`, including `rootCommitId` and `currentBranchName`)

Source: [`packages/versioning/branches/src/store/YjsBranchStore.js`](../packages/versioning/branches/src/store/YjsBranchStore.js)

#### Custom branch root name (`rootName`)

Both `YjsBranchStore` and `CollabBranchingWorkflow` support a non-default `rootName` (prefix) for these roots:

```ts
const rootName = "myBranching";
const store = new YjsBranchStore({ ydoc: session.doc, rootName });
const branchService = new BranchService({ docId, store });
const workflow = new CollabBranchingWorkflow({ session, branchService, rootName });
```

If you customize `rootName`:

- update `CollabVersioningOptions.excludeRoots` so version snapshots/restores don’t include (and later rewind) the branch/commit graph (see “Snapshot/restore isolation” above)
- if using sync-server in production, consider extending `SYNC_SERVER_RESERVED_ROOT_PREFIXES` so the reserved-root guard still blocks your custom prefix

### Branch commit size + sync-server message limits

Branching history is stored *inside the shared Yjs document* under `branching:commits`.
Each commit record can include a semantic `patch` (and sometimes a `snapshot`).

For large workbooks (especially when cell payloads are **encrypted** or otherwise high-entropy),
these payloads can become large enough that a single Yjs update exceeds the sync-server
`SYNC_SERVER_MAX_MESSAGE_BYTES` limit (default **2MB**). When that happens, the server will close
the websocket with code **1009** ("Message too big") and branching becomes unusable.

To make large commits robust under realistic message-size limits, configure the Yjs-backed store
to compress + chunk commit payloads:

```js
const store = new YjsBranchStore({
  ydoc: session.doc,
  payloadEncoding: "gzip-chunks",
  chunkSize: 64 * 1024,
  maxChunksPerTransaction: 16,
});
```

Tuning notes:

- `chunkSize` should be comfortably below your `SYNC_SERVER_MAX_MESSAGE_BYTES` (leave room for protocol overhead).
- `maxChunksPerTransaction` bounds the size of each individual Yjs update message.

### “Use this in collaborative mode” snippet (CollabBranchingWorkflow)

```ts
import { CollabBranchingWorkflow } from "../packages/collab/branching/index.js";
import { BranchService, YjsBranchStore, yjsDocToDocumentState } from "../packages/versioning/branches/src/index.js";

const store = new YjsBranchStore({ ydoc: session.doc }); // writes under branching:*
const branchService = new BranchService({ docId, store });

// Collab wrapper applies checkout/merge results back into the *live* Y.Doc.
const workflow = new CollabBranchingWorkflow({ session, branchService });

// Initialize (creates main branch if needed).
await branchService.init({ userId, role: "owner" }, yjsDocToDocumentState(session.doc));

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
