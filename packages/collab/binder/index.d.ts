import type { Doc } from "yjs";

/**
 * Canonical sheet view state stored in Yjs under `sheets[i].get("view")`.
 *
 * This is normalized by the binder to the shape expected by the desktop
 * `DocumentController`.
 */
export interface SheetViewState {
  frozenRows: number;
  frozenCols: number;
  /**
   * Per-column width overrides keyed by zero-based column index.
   *
   * Keys are serialized as strings to match Yjs map/object encoding.
   */
  colWidths?: Record<string, number>;
  /**
   * Per-row height overrides keyed by zero-based row index.
   *
   * Keys are serialized as strings to match Yjs map/object encoding.
   */
  rowHeights?: Record<string, number>;
}

/**
 * A cell coordinate used throughout collab packages.
 */
export interface CellAddress {
  sheetId: string;
  row: number;
  col: number;
}

export interface CellEncryptionKey {
  keyId: string;
  keyBytes: Uint8Array;
}

export interface EncryptionConfig {
  /**
   * Return an encryption key for a specific cell, or `null` if the cell cannot be
   * encrypted/decrypted by this client.
   */
  keyForCell: (cell: CellAddress) => CellEncryptionKey | null;
  /**
   * Opt-in: when true, per-cell formatting (`format`) is stored inside the
   * encrypted payload (`enc`) instead of being written in plaintext.
   *
   * Defaults to false for backwards compatibility.
   */
  encryptFormat?: boolean;
  /**
   * Determines whether the binder should write encrypted payloads into Yjs for
   * this cell.
   *
   * Note: once a cell is encrypted in Yjs, the binder will keep it encrypted
   * regardless of this predicate to avoid leaking plaintext writes.
   */
  shouldEncryptCell?: (cell: CellAddress) => boolean;
}

export type DocumentRole = "owner" | "admin" | "editor" | "commenter" | "viewer";

export interface CellPermissions {
  canRead: boolean;
  canEdit: boolean;
}

/**
 * A permissive permissions resolver: the binder treats missing/`undefined`
 * properties as `true` and only respects explicit `false`.
 */
export type PermissionsResolver = (
  cell: CellAddress,
) => Partial<CellPermissions> | null | undefined;

export type RangeLike = {
  sheetId?: string;
  sheetName?: string;
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
};

export type RestrictionLike =
  | (RangeLike & {
      id?: string;
      createdAt?: string | number | Date;
      readAllowlist?: string[];
      editAllowlist?: string[];
    })
  | {
      id?: string;
      createdAt?: string | number | Date;
      range: RangeLike;
      readAllowlist?: string[];
      editAllowlist?: string[];
    };

export interface RoleBasedPermissions {
  role: DocumentRole;
  restrictions?: readonly RestrictionLike[] | null;
  /**
   * User id used to evaluate allowlists in restrictions. Defaults to the binder
   * `userId` option when omitted.
   */
  userId?: string | null;
}

export interface UndoServiceLike {
  /**
   * Run a set of Yjs mutations inside an undo-aware transaction.
   *
   * If omitted, the binder will fall back to `ydoc.transact(fn, origin)`.
   */
  transact?: (fn: () => void) => void;
  /**
   * Origin token used for binder-driven transactions (DocumentController → Yjs).
   */
  origin?: unknown;
  /**
   * Set of origins considered local by the undo service.
   */
  localOrigins?: Set<unknown>;
}

export type FormulaConflictsMode = "off" | "formula" | "formula+value";

export interface BindYjsToDocumentControllerOptions {
  ydoc: Doc;
  /**
   * Desktop document controller instance. Intentionally typed structurally/loosely
   * to avoid importing app-internal types from `apps/`.
   */
  documentController: any;

  undoService?: UndoServiceLike | null;
  defaultSheetId?: string;
  userId?: string | null;

  encryption?: EncryptionConfig | null;

  /**
   * Legacy guard: merged with `permissions` when both are provided.
   */
  canReadCell?: ((cell: CellAddress) => boolean) | null;
  /**
   * Legacy guard: merged with `permissions` when both are provided.
   */
  canEditCell?: ((cell: CellAddress) => boolean) | null;

  /**
   * Either a per-cell resolver or a role+restrictions object.
   */
  permissions?: PermissionsResolver | RoleBasedPermissions | null;

  /**
   * Called when the binder needs to mask a cell value (e.g. unreadable or
   * encrypted cells).
   */
  maskCellValue?: ((value: unknown, cell?: CellAddress) => unknown) | null;

  /**
   * Controls whether DocumentController-driven shared-state writes (sheet view
   * state, sheet-level formatting defaults, range-run formatting metadata, etc)
   * are allowed to be persisted into the shared Yjs document.
   *
   * When false, those deltas are ignored so the UI can still support local-only
   * interactions (useful for read-only collaboration roles).
   *
   * Defaults to `true`.
   */
  canWriteSharedState?: boolean | (() => boolean);

  /**
   * When true, the binder will also mask **formatting** (styleId/format objects)
   * for unreadable/encrypted cells, not just their values.
   *
   * Defaults to false.
   */
  maskCellFormat?: boolean;

  /**
   * Called when the binder rejects a DocumentController-driven edit (e.g. due to
   * insufficient permissions).
   */
  onEditRejected?: ((deltas: any[]) => void) | null;

  /**
   * Enables write semantics that are compatible with causal conflict detection
   * (e.g. FormulaConflictMonitor) by writing explicit `null` markers for clears.
   */
  formulaConflictsMode?: FormulaConflictsMode;
}

export interface BindYjsToDocumentControllerBinding {
  destroy(): void;
  /**
   * Re-scan Yjs state and re-apply it into the DocumentController.
   *
   * This is primarily used when local decryption state changes (e.g. a user
   * imports an encryption key) without any Yjs document mutation.
   */
  rehydrate(): void;
  /**
   * Wait for any pending binder work to settle.
   *
   * Useful for teardown flows (e.g. flushing local persistence before a hard
   * process exit).
   */
  whenIdle(): Promise<void>;
}

export function bindYjsToDocumentController(
  options: BindYjsToDocumentControllerOptions,
): BindYjsToDocumentControllerBinding;
