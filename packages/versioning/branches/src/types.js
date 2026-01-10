/**
 * @typedef {Record<string, unknown>} JsonObject
 */

/**
 * A spreadsheet cell.
 *
 * `value` and `formula` are mutually exclusive; both are optional because an
 * empty cell is represented as `null`/`undefined`.
 *
 * @typedef {{
 *   value?: unknown,
 *   formula?: string,
 *   format?: JsonObject
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
 * @typedef {{
 *   sheets: Record<string, CellMap>
 * }} DocumentState
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
 * }} MergeConflict
 */

/**
 * @typedef {{
 *   merged: DocumentState,
 *   conflicts: MergeConflict[]
 * }} MergeResult
 */

/**
 * @typedef {"owner" | "admin" | "editor" | "viewer"} Role
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

