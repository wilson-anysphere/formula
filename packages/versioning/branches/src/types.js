/**
 * @typedef {Record<string, unknown>} JsonObject
 */

/**
 * A spreadsheet cell.
 *
 * `value` and `formula` are mutually exclusive; both are optional because an
 * empty cell is represented as `null`/`undefined`.
 *
 * `enc` optionally stores an encrypted cell payload as written by
 * `@formula/collab-session` (under the Yjs `enc` field). This is treated as an
 * opaque blob by BranchService so encrypted workbooks can be branched/merged
 * without losing ciphertext or leaking plaintext.
 *
 * @typedef {{
 *   value?: unknown,
 *   formula?: string,
 *   format?: JsonObject,
 *   enc?: unknown
 * }} Cell
 */

/**
 * Map of cell address (e.g. "A1") to {@link Cell}. Missing keys represent empty
 * cells.
 *
 * @typedef {Record<string, Cell>} CellMap
 */

/**
 * Multi-sheet document state.
 *
 * Legacy (v0) state shape used by early BranchService versions (cells only).
 *
 * @typedef {{
 *   sheets: Record<string, CellMap>
 * }} LegacyDocumentState
 */

/**
 * Minimal sheet metadata tracked by branching/versioning.
 *
 * `id` is the stable sheet identifier used by collaboration and cell keys.
 * `name` is the user-visible display name (nullable for older/malformed docs).
 * `view` stores per-sheet UI state that should survive undo/redo and semantic merges
 * (e.g. frozen panes).
 * `visibility` controls whether the sheet tab is visible to the user (Excel-like semantics).
 * `tabColor` stores an optional sheet tab color as an 8-digit ARGB hex string.
 *
 * @typedef {{
 *   frozenRows: number,
 *   frozenCols: number,
 *   /**
 *    * Optional worksheet background image id (tiled behind the grid).
 *    *
 *    * Stored as an opaque identifier; image bytes are managed separately.
 *    *
 *    * When present with a `null` value, this represents an explicit "clear background"
 *    * operation (used so semantic merges can distinguish omission from clears).
 *    *\/
 *   backgroundImageId?: string | null,
 *   /**
 *    * Sparse column width overrides (base units, zoom=1), keyed by 0-based column index.
 *    *\/
 *   colWidths?: Record<string, number>,
 *   /**
 *    * Sparse row height overrides (base units, zoom=1), keyed by 0-based row index.
 *    *\/
 *   rowHeights?: Record<string, number>,
 *   /**
 *    * Optional sheet-level merged ranges.
 *    *
 *    * Stored as a list of inclusive rectangle coordinates.
 *    *\/
 *   mergedRanges?: Array<{ startRow: number, endRow: number, startCol: number, endCol: number }>,
 *   /**
 *    * Optional sheet-level drawing metadata (images, shapes, etc).
 *    *
 *    * This is treated as opaque JSON by BranchService; image bytes and other
 *    * large payloads are managed separately.
 *    *\/
 *   drawings?: unknown[],
 *   /**
 *    * Default sheet format applied to all cells unless overridden by a row/column/cell format.
 *    *
 *    * Stored as a style object (not a style id) so BranchService snapshots are self-contained.
 *    *\/
 *   defaultFormat?: JsonObject,
 *   /**
 *    * Sparse row format overrides keyed by 0-based row index.
 *    *
 *    * Values are style objects (not style ids).
 *    *\/
 *   rowFormats?: Record<string, JsonObject>,
 *   /**
 *    * Sparse column format overrides keyed by 0-based column index.
 *    *
 *    * Values are style objects (not style ids).
 *    *\/
 *   colFormats?: Record<string, JsonObject>,
 *   /**
 *    * Range-run formatting (compressed rectangular formatting) stored as a per-column list
 *    * of half-open row intervals: `[startRow, endRowExclusive)`.
 *    *
 *    * This mirrors DocumentController's `formatRunsByCol` snapshot representation and allows
 *    * branches to persist large rectangle format changes without materializing per-cell
 *    * `Cell.format` overrides.
 *    *
 *    * Values are style objects (not style ids) so BranchService snapshots are self-contained.
 *    *\/
 *   formatRunsByCol?: Array<{ col: number, runs: Array<{ startRow: number, endRowExclusive: number, format: JsonObject }> }>,
 * }} SheetViewState
 *
 * @typedef {"visible" | "hidden" | "veryHidden"} SheetVisibility
 *
 * @typedef {{
 *   id: string,
 *   name: string | null,
 *   view?: SheetViewState,
 *   visibility?: SheetVisibility,
 *   /**
 *    * Optional sheet tab color as an 8-digit ARGB hex string (e.g. "FFFF0000").
 *    * `null` indicates an explicit "no color" override; `undefined` means "unknown/untracked".
 *    *\/
 *   tabColor?: string | null,
 * }} SheetMeta
 */

/**
 * Workbook-level sheet state.
 *
 * @typedef {{
 *   order: string[],
 *   metaById: Record<string, SheetMeta>
 * }} SheetsState
 */

/**
 * Workbook document state for BranchService v2.
 *
 * This matches the collaboration/versioning surface area (cells + workbook
 * metadata) so scenario branches can preserve realistic spreadsheet workflows.
 *
 * @typedef {{
 *   schemaVersion: 1,
 *   sheets: SheetsState,
 *   cells: Record<string, CellMap>,
 *   metadata: Record<string, any>,
 *   namedRanges: Record<string, any>,
 *   comments: Record<string, any>
 * }} WorkbookDocumentState
 */

/**
 * Alias used throughout the BranchService package.
 *
 * @typedef {WorkbookDocumentState} DocumentState
 */

/**
 * A semantic merge conflict.
 *
 * For `type: "cell"`, `cell` indicates the conflicting address.
 * For `type: "move"`, `cell` is the *source* address of the moved cell.
 *
 * @typedef {{
 *   type: "cell",
 *   sheetId: string,
 *   cell: string,
 *   reason:
 *     | "content"
 *     | "format"
 *     | "delete-vs-edit"
 *     | "move-destination",
 *   base: Cell | null,
 *   ours: Cell | null,
 *   theirs: Cell | null
 * } | {
 *   type: "move",
 *   sheetId: string,
 *   cell: string,
 *   reason: "move-destination",
 *   base: Cell | null,
 *   ours: { to: string } | null,
 *   theirs: { to: string } | null
 * } | {
 *   type: "sheet",
 *   reason: "rename" | "order" | "presence",
 *   sheetId?: string,
 *   base: any,
 *   ours: any,
 *   theirs: any
 * } | {
 *   type: "metadata",
 *   key: string,
 *   base: any,
 *   ours: any,
 *   theirs: any
 * } | {
 *   type: "namedRange",
 *   key: string,
 *   base: any,
 *   ours: any,
 *   theirs: any
 * } | {
 *   type: "comment",
 *   id: string,
 *   base: any,
 *   ours: any,
 *   theirs: any
 * }} MergeConflict
 */

/**
 * @typedef {{
 *   merged: DocumentState,
 *   conflicts: MergeConflict[]
 * }} MergeResult
 */

/**
 * @typedef {"owner" | "admin" | "editor" | "commenter" | "viewer"} Role
 */

/**
 * @typedef {{
 *   userId: string,
 *   role: Role
 * }} Actor
 */

/**
 * @typedef {{
 *   id: string,
 *   docId: string,
 *   name: string,
 *   createdBy: string,
 *   createdAt: number,
 *   description: string | null,
 *   headCommitId: string
 * }} Branch
 */

/**
 * @typedef {{
 *   id: string,
 *   docId: string,
 *   parentCommitId: string | null,
 *   mergeParentCommitId: string | null,
 *   createdBy: string,
 *   createdAt: number,
 *   message: string | null,
 *   patch: import("./patch.js").Patch
 * }} Commit
 */

export {};
