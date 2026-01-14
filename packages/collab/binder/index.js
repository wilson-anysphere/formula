import * as Y from "yjs";
import { getYMap, getYText, yjsValueToJson } from "@formula/collab-yjs-utils";

import { makeCellKey, normalizeCellKey, parseCellKey } from "../session/src/cell-key.js";
import { getWorkbookRoots } from "../workbook/src/index.ts";
import {
  decryptCellPlaintext,
  encryptCellPlaintext,
  isEncryptedCellPayload,
} from "../encryption/src/index.node.js";
import { getCellPermissions, maskCellValue as defaultMaskCellValue } from "../permissions/index.js";

const MASKED_CELL_VALUE = "###";
// Defensive cap: drawings metadata can be remote-authored (sheet view state) and is preserved by
// this binder even though it is not explicitly synced. Keep validation strict to avoid unbounded
// costs when cloning/upgrading legacy sheet entries.
const MAX_DRAWING_ID_STRING_CHARS = 4096;

function stableStringify(value) {
  if (value === undefined) return "undefined";
  if (value == null || typeof value !== "object") return JSON.stringify(value);
  if (Array.isArray(value)) return `[${value.map(stableStringify).join(",")}]`;
  const keys = Object.keys(value).sort();
  const entries = keys.map((k) => `${JSON.stringify(k)}:${stableStringify(value[k])}`);
  return `{${entries.join(",")}}`;
}

function isRecord(value) {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

/**
 * @param {any} raw
 * @returns {any}
 */
function sanitizeDrawingsForPreservation(raw) {
  if (!Array.isArray(raw)) return raw;
  /** @type {any[]} */
  const out = [];
  let changed = false;
  for (const entry of raw) {
    if (!entry || typeof entry !== "object") {
      changed = true;
      continue;
    }
    const rawId = entry?.get?.("id") ?? entry.id;
    if (typeof rawId === "string") {
      if (rawId.length > MAX_DRAWING_ID_STRING_CHARS) {
        changed = true;
        continue;
      }
      if (!rawId.trim()) {
        changed = true;
        continue;
      }
    } else if (typeof rawId === "number") {
      if (!Number.isSafeInteger(rawId)) {
        changed = true;
        continue;
      }
    } else {
      changed = true;
      continue;
    }
    out.push(entry);
  }
  return changed ? out : raw;
}

/**
 * Strip invalid/pathological drawing ids from a sheet view payload.
 *
 * @param {any} view
 * @returns {any}
 */
function sanitizeSheetViewForPreservation(view) {
  if (!isRecord(view)) return view;
  if (!Object.prototype.hasOwnProperty.call(view, "drawings")) return view;
  const sanitized = sanitizeDrawingsForPreservation(view.drawings);
  if (sanitized === view.drawings) return view;
  if (!Array.isArray(sanitized) || sanitized.length === 0) {
    const { drawings: _ignored, ...rest } = view;
    return rest;
  }
  return { ...view, drawings: sanitized };
}

/**
 * @param {any} value
 * @returns {string | null}
 */
function coerceString(value) {
  const text = getYText(value);
  if (text) return yjsValueToJson(text);
  if (typeof value === "string") return value;
  if (value == null) return null;
  return String(value);
}

/**
 * Normalize a potentially user-provided id (string-ish) into a trimmed non-empty string,
 * or `null` if it is unset/invalid.
 *
 * @param {any} value
 * @returns {string | null}
 */
function normalizeOptionalId(value) {
  const raw = coerceString(value);
  if (typeof raw !== "string") return null;
  const trimmed = raw.trim();
  return trimmed ? trimmed : null;
}

/**
 * @param {any} value
 * @returns {number}
 */
function normalizeFrozenCount(value) {
  const num = Number(value);
  if (!Number.isFinite(num)) return 0;
  return Math.max(0, Math.trunc(num));
}

/**
 * Normalize Yjs `sheet.view` state into the {@link SheetViewState} shape expected
 * by DocumentController.
 *
 * @param {any} view
 * @returns {{ frozenRows: number, frozenCols: number, backgroundImageId?: string, colWidths?: Record<string, number>, rowHeights?: Record<string, number> }}
 */
function normalizeSheetViewState(view) {
  const normalizeAxisSize = (value) => {
    const num = Number(value);
    if (!Number.isFinite(num)) return null;
    if (num <= 0) return null;
    return num;
  };

  const normalizeAxisOverrides = (raw) => {
    if (!raw) return null;

    /** @type {Record<string, number>} */
    const out = {};

    if (Array.isArray(raw)) {
      for (const entry of raw) {
        const index = Array.isArray(entry) ? entry[0] : entry?.index;
        const size = Array.isArray(entry) ? entry[1] : entry?.size;
        const idx = Number(index);
        if (!Number.isInteger(idx) || idx < 0) continue;
        const normalized = normalizeAxisSize(size);
        if (normalized == null) continue;
        out[String(idx)] = normalized;
      }
    } else if (typeof raw === "object") {
      for (const [key, value] of Object.entries(raw)) {
        const idx = Number(key);
        if (!Number.isInteger(idx) || idx < 0) continue;
        const normalized = normalizeAxisSize(value);
        if (normalized == null) continue;
        out[String(idx)] = normalized;
      }
    }

    return Object.keys(out).length === 0 ? null : out;
  };

  const colWidths = normalizeAxisOverrides(view?.colWidths);
  const rowHeights = normalizeAxisOverrides(view?.rowHeights);
  const backgroundImageId = normalizeOptionalId(view?.backgroundImageId ?? view?.background_image_id);

  return {
    frozenRows: normalizeFrozenCount(view?.frozenRows),
    frozenCols: normalizeFrozenCount(view?.frozenCols),
    ...(backgroundImageId ? { backgroundImageId } : {}),
    ...(colWidths ? { colWidths } : {}),
    ...(rowHeights ? { rowHeights } : {}),
  };
}

function emptySheetViewState() {
  return { frozenRows: 0, frozenCols: 0 };
}

/**
 * @param {any} a
 * @param {any} b
 */
function sheetViewStateEquals(a, b) {
  if (a === b) return true;
  if (!a || !b) return false;
  if (a.frozenRows !== b.frozenRows) return false;
  if (a.frozenCols !== b.frozenCols) return false;
  if (normalizeOptionalId(a.backgroundImageId) !== normalizeOptionalId(b.backgroundImageId)) return false;

  const axisEquals = (left, right) => {
    if (left === right) return true;
    const leftKeys = left ? Object.keys(left) : [];
    const rightKeys = right ? Object.keys(right) : [];
    if (leftKeys.length !== rightKeys.length) return false;
    leftKeys.sort((x, y) => Number(x) - Number(y));
    rightKeys.sort((x, y) => Number(x) - Number(y));
    for (let i = 0; i < leftKeys.length; i += 1) {
      const key = leftKeys[i];
      if (key !== rightKeys[i]) return false;
      const lv = left[key];
      const rv = right[key];
      if (Math.abs(lv - rv) > 1e-6) return false;
    }
    return true;
  };

  return axisEquals(a.colWidths, b.colWidths) && axisEquals(a.rowHeights, b.rowHeights);
}

function getYMapCell(cellData) {
  return getYMap(cellData);
}

function normalizeFormula(value) {
  const json = yjsValueToJson(value);
  if (json == null) return null;
  const trimmed = String(json).trim();
  const strippedLeading = trimmed.startsWith("=") ? trimmed.slice(1) : trimmed;
  const stripped = strippedLeading.trim();
  if (stripped === "") return null;
  return `=${stripped}`;
}

/**
 * @typedef {{
 *   value: any,
 *   formula: string | null,
 *   formatKey: string | undefined,
 * }} NormalizedCell
 */

/**
 * @typedef {{
 *   value: any,
 *   formula: string | null,
 *   format: any | undefined,
 *   formatKey: string | undefined,
 *   hasEnc: boolean,
 *   maskedForEncryption?: boolean,
 * }} ParsedYjsCell
 */

/**
 * @param {Y.Map<any>} cell
 * @param {{ sheetId: string, row: number, col: number }} cellRef
 * @param {{
 *   keyForCell: (cell: { sheetId: string, row: number, col: number }) => { keyId: string, keyBytes: Uint8Array } | null,
 *   encryptFormat?: boolean,
 * } | null} encryption
 * @param {string} docIdForEncryption
 * @param {(value: unknown, cell?: { sheetId: string, row: number, col: number }) => unknown} maskFn
 * @returns {Promise<ParsedYjsCell>}
 */
async function readCellFromYjs(cell, cellRef, encryption, docIdForEncryption, maskFn = defaultMaskCellValue) {
  let value;
  let formula;
  let hasEnc = false;
  let maskedForEncryption = false;
  const encryptFormat = Boolean(encryption?.encryptFormat);
  /** @type {any | null} */
  let decryptedPlaintext = null;

  // If `enc` is present, treat the cell as encrypted even if the payload is malformed.
  // This avoids accidentally falling back to plaintext fields.
  const encRaw = typeof cell.has === "function" ? (cell.has("enc") ? cell.get("enc") : undefined) : cell.get("enc");
  if (encRaw !== undefined) {
    hasEnc = true;
    if (isEncryptedCellPayload(encRaw)) {
      const key = encryption?.keyForCell?.(cellRef) ?? null;
      if (key && key.keyId === encRaw.keyId) {
        try {
          const plaintext = await decryptCellPlaintext({
            encrypted: encRaw,
            key,
            context: {
              docId: docIdForEncryption,
              sheetId: cellRef.sheetId,
              row: cellRef.row,
              col: cellRef.col,
            },
          });
          decryptedPlaintext = plaintext;

          const decryptedFormula = normalizeFormula(plaintext?.formula ?? null);
          if (decryptedFormula) {
            value = null;
            formula = decryptedFormula;
          } else {
            value = plaintext?.value ?? null;
            formula = null;
          }
        } catch {
          // fall through to masked state
        }
      }
    }

    if (value === undefined && formula === undefined) {
      // `enc` is present but we can't decrypt (missing key, wrong key id, corrupt payload, etc).
      // Always surface a masked placeholder even if `maskFn` is permission-aware and would
      // otherwise return the input value unchanged.
      value = maskFn(MASKED_CELL_VALUE, cellRef);
      formula = null;
      maskedForEncryption = true;
    }
  } else {
    formula = normalizeFormula(cell.get("formula") ?? null);
    if (formula) {
      value = null;
    } else {
      value = cell.get("value") ?? null;
      formula = null;
    }
  }

  let format = undefined;
  let formatKey = undefined;
  if (hasEnc && encryptFormat) {
    // Confidentiality-first semantics:
    // - If `enc` is present, do not read/use plaintext `format`.
    // - Only apply formatting that was successfully decrypted from the encrypted payload.
    const decryptedFormat = decryptedPlaintext?.format ?? null;
    if (decryptedFormat != null) {
      format = decryptedFormat;
      formatKey = stableStringify(format);
    }
  } else {
    // Backwards-compatible alias: some legacy clients stored per-cell formatting under `style`
    // instead of `format`. Prefer `format` when present.
    const raw = cell.get("format");
    const rawOrStyle = raw === undefined ? cell.get("style") : raw;
    if (rawOrStyle !== undefined) {
      format = rawOrStyle ?? null;
      formatKey = stableStringify(format);
    }
  }
  if (formula) {
    return { value: null, formula, format, formatKey, hasEnc, maskedForEncryption };
  }
  return { value, formula: null, format, formatKey, hasEnc, maskedForEncryption };
}

function sameNormalizedCell(a, b) {
  if (!a && !b) return true;
  if (!a || !b) return false;
  return a.formula === b.formula && Object.is(a.value, b.value) && a.formatKey === b.formatKey;
}

/**
 * Bind a Yjs spreadsheet document to a desktop `DocumentController`.
 *
 * This binder currently syncs:
 *
 * - Cell contents: `value`, `formula`, and `format` (cell styles) via the Yjs `cells` root
 * - Sheet view state: `sheets[].view` (frozen panes + row/col size overrides)
 * - Layered formatting defaults: `defaultFormat`, `rowFormats`, `colFormats` (sheet/row/col styles)
 * - Range-run formatting: `formatRunsByCol` (compressed rectangular formatting; used for large ranges)
 *
 * It does **not** currently implement:
 * - full sheet list semantics (create/delete/rename/reorder), or
 * - per-sheet metadata syncing (e.g. `visibility`, `tabColor`)
 *
 * @param {{
 *   ydoc: import("yjs").Doc,
 *   documentController: import("../../../apps/desktop/src/document/documentController.js").DocumentController,
 *   undoService?: { transact?: (fn: () => void) => void, origin?: any, localOrigins?: Set<any> } | null,
 *   defaultSheetId?: string,
 *   userId?: string | null,
 *   encryption?: {
 *     keyForCell: (cell: { sheetId: string, row: number, col: number }) => { keyId: string, keyBytes: Uint8Array } | null,
 *     /**
 *      * Optional override for deciding whether a cell should be encrypted.
 *      * Defaults to `true` when `keyForCell` returns a non-null key.
 *      *
 *      * Tip: to drive encryption policy from shared workbook metadata (so clients
 *      * without keys can still refuse plaintext writes), use
 *      * `createEncryptionPolicyFromDoc(doc)` from `@formula/collab-encrypted-ranges`
 *      * and pass its `shouldEncryptCell` here.
 *      *\/
 *     shouldEncryptCell?: (cell: { sheetId: string, row: number, col: number }) => boolean,
 *     encryptFormat?: boolean,
 *   } | null,
 *   canReadCell?: (cell: { sheetId: string, row: number, col: number }) => boolean,
 *   canEditCell?: (cell: { sheetId: string, row: number, col: number }) => boolean,
 *   permissions?:
 *     | ((cell: { sheetId: string, row: number, col: number }) => { canRead: boolean, canEdit: boolean })
 *     | { role: string, restrictions?: any[], userId?: string | null },
 *   /**
 *    * Controls whether DocumentController-driven shared-state writes (sheet view state,
 *    * sheet-level formatting defaults, etc) are allowed to be persisted into the Yjs
  *    * document.
  *    *
  *    * In non-edit collab roles (viewer/commenter), some UI interactions may still
  *    * emit "shared state" deltas (sheet view state, formatting defaults, etc). When
  *    * this hook returns false, the binder prevents those deltas from being written
  *    * into the shared CRDT. UIs can still allow these mutations to remain local-only
  *    * (e.g. freeze panes / row/col sizing / formatting defaults), while true cell
  *    * edits remain blocked via `canEditCell`.
  *    *
  *    * Defaults to `true`.
  *    *\/
  *   canWriteSharedState?: boolean | (() => boolean),
 *   maskCellValue?: (value: unknown, cell?: { sheetId: string, row: number, col: number }) => unknown,
 *   /**
 *    * When true, suppress per-cell formatting for masked cells (unreadable due to
 *    * permissions, or encrypted without an available key).
 *    *
 *    * This avoids leaking sensitive information via formatting (e.g. number/date
 *    * formats, fills, conditional styling) and makes masked cells visually appear
 *    * "masked".
 *    *
 *    * Defaults to false to preserve existing formatting semantics.
 *    *\/
 *   maskCellFormat?: boolean,
 *   onEditRejected?: (deltas: any[]) => void,
 *   /**
 *    * Enables Yjs write semantics that are compatible with `FormulaConflictMonitor`'s
 *    * causal conflict detection (e.g. representing clears with explicit `null`
 *    * marker Items so later overwrites can reference them via `Item.origin`).
 *    *
 *    * When `"formula"`:
 *    * - formula clears write `cell.set("formula", null)` (instead of `delete("formula")`)
 *    * - empty cell maps created by formula clears are preserved (not deleted) so
 *    *   the clear marker Item remains in the CRDT
 *    *
 *    * When `"formula+value"`:
 *    * - value writes also clear formulas via `cell.set("formula", null)` so
 *    *   formula-vs-value content conflicts are detectable
 *    * - empty cell maps created by value/formula clears are preserved when they
 *    *   would otherwise delete the clear marker Items
 *    *
 *    * Defaults to `"off"` to avoid doc growth/regressions when conflict monitoring
 *    * is not enabled.
 *    *\/
 *   formulaConflictsMode?: "off" | "formula" | "formula+value",
 * }} options
 */
export function bindYjsToDocumentController(options) {
  const {
    ydoc,
    documentController,
    undoService = null,
    defaultSheetId = "Sheet1",
    userId = null,
    encryption = null,
    canReadCell = null,
    canEditCell = null,
    permissions = null,
    canWriteSharedState: canWriteSharedStateRaw = true,
    maskCellValue = defaultMaskCellValue,
    maskCellFormat = false,
    onEditRejected = null,
    formulaConflictsMode: formulaConflictsModeRaw = "off",
  } = options ?? {};

  if (!ydoc) throw new Error("bindYjsToDocumentController requires { ydoc }");
  if (!documentController) throw new Error("bindYjsToDocumentController requires { documentController }");

  /** @type {"off" | "formula" | "formula+value"} */
  const formulaConflictsMode =
    formulaConflictsModeRaw === "formula" || formulaConflictsModeRaw === "formula+value" || formulaConflictsModeRaw === "off"
      ? formulaConflictsModeRaw
      : "off";
  const formulaConflictSemanticsEnabled = formulaConflictsMode !== "off";
  const valueConflictSemanticsEnabled = formulaConflictsMode === "formula+value";

  const canWriteSharedState =
    typeof canWriteSharedStateRaw === "function"
      ? canWriteSharedStateRaw
      : () => canWriteSharedStateRaw !== false;

  const { cells, sheets } = getWorkbookRoots(ydoc);
  const YMapCtor = cells?.constructor ?? Y.Map;
  const docIdForEncryption = ydoc.guid ?? "unknown";

  // Stable origin token for local DocumentController -> Yjs transactions when
  // we don't have a dedicated undo service wrapper.
  const binderOrigin = undoService?.origin ?? { type: "document-controller:binder" };
  /** @type {Map<string, NormalizedCell>} */
  let cache = new Map();

  let applyingRemote = false;
  // True while this binder is applying DocumentController-driven writes into Yjs.
  // Used to suppress observer echo work without suppressing other session-origin
  // mutations (branch checkout/restore, programmatic `session.setCell*`, etc).
  let applyingLocal = false;
  let hasEncryptedCells = false;

  const permissionResolver = createPermissionsResolver(permissions, userId);
  const legacyReadGuard = typeof canReadCell === "function" ? canReadCell : null;
  const legacyEditGuard = typeof canEditCell === "function" ? canEditCell : null;

  /**
   * @param {{ sheetId: string, row: number, col: number }} cell
   */
  const resolveCellPermissions = (cell) => {
    const resolved = permissionResolver ? permissionResolver(cell) : { canRead: true, canEdit: true };
    let canRead = resolved.canRead;
    let canEdit = resolved.canEdit;

    if (legacyReadGuard) canRead = canRead && legacyReadGuard(cell);
    if (legacyEditGuard) canEdit = canEdit && legacyEditGuard(cell);
    if (!canRead) canEdit = false;

    return { canRead, canEdit };
  };

  const readGuard = legacyReadGuard || permissionResolver ? (cell) => resolveCellPermissions(cell).canRead : null;
  const editGuard = legacyEditGuard || permissionResolver ? (cell) => resolveCellPermissions(cell).canEdit : null;
  const maskFn = typeof maskCellValue === "function" ? maskCellValue : defaultMaskCellValue;

  // Ensure local user edits in DocumentController are permission-aware by default.
  // DocumentController.applyExternalDeltas intentionally bypasses canEditCell, but user
  // mutations (setCellValue, setCellFormula, etc) consult this hook.
  const prevDocumentCanEditCell =
    documentController && typeof documentController === "object" && "canEditCell" in documentController
      ? documentController.canEditCell
      : undefined;
  const didPatchDocumentCanEditCell = Boolean(editGuard && prevDocumentCanEditCell !== undefined);
  if (didPatchDocumentCanEditCell) {
    if (typeof prevDocumentCanEditCell === "function") {
      documentController.canEditCell = (cell) => prevDocumentCanEditCell.call(documentController, cell) && editGuard(cell);
    } else {
      documentController.canEditCell = editGuard;
    }
  }

  /**
   * Track raw Yjs keys that correspond to a canonical `${sheetId}:${row}:${col}` key.
   *
   * This lets us apply DocumentController -> Yjs mutations without needing to scan
   * the full cells map even when the doc contains historical key encodings
   * (`${sheetId}:${row},${col}` or `r{row}c{col}`).
   *
   * @type {Map<string, Set<string>>}
  */
  const yjsKeysByCell = new Map();

  // Serialize Yjs -> DocumentController updates (decryption can be async).
  let applyChain = Promise.resolve();
  // Serialize DocumentController -> Yjs writes (encryption can be async).
  let writeChain = Promise.resolve();
  // Serialize DocumentController -> Yjs sheet metadata writes (view state, etc).
  let sheetWriteChain = Promise.resolve();

  /**
   * Normalize foreign nested Y.Maps (e.g. created by a different `yjs` module instance
   * in mixed ESM/CJS environments) into local types before mutating them.
   *
   * This conversion is intentionally performed in an *untracked* transaction (no origin,
   * i.e. not in UndoManager.trackedOrigins) so collaborative undo only captures the
   * user's actual edit.
   *
   * @param {Iterable<string>} rawKeys
   */
  function ensureLocalCellMapsForWrite(rawKeys) {
    /** @type {string[]} */
    const foreign = [];
    for (const rawKey of rawKeys) {
      if (typeof rawKey !== "string") continue;
      const existingCell = getYMapCell(cells.get(rawKey));
      if (!existingCell) continue;
      if (existingCell instanceof Y.Map) continue;
      foreign.push(rawKey);
    }

    if (foreign.length === 0) return;

    const prevApplyingLocal = applyingLocal;
    applyingLocal = true;
    try {
      // Intentionally run in an *untracked* transaction (no origin) so collaborative undo
      // only captures the user's actual edit.
      ydoc.transact(() => {
        for (const rawKey of foreign) {
          const cellData = cells.get(rawKey);
          const cell = getYMapCell(cellData);
          if (!cell || cell instanceof Y.Map) continue;

          const local = new Y.Map();
          cell.forEach((v, k) => {
            local.set(k, v);
          });
          cells.set(rawKey, local);
        }
      });
    } finally {
      applyingLocal = prevApplyingLocal;
    }
  }

  /**
   * @param {Set<string>} changedCanonicalKeys
   */
  function enqueueApply(changedCanonicalKeys) {
    if (!changedCanonicalKeys || changedCanonicalKeys.size === 0) return;
    const snapshot = new Set(changedCanonicalKeys);
    applyChain = applyChain.then(() => applyYjsChangesToDocumentController(snapshot)).catch((err) => {
      // Avoid breaking the chain.
      console.error(err);
    });
  }

  /**
   * @param {Set<string>} changedSheetIds
   */
  function enqueueSheetViewApply(changedSheetIds) {
    if (!changedSheetIds || changedSheetIds.size === 0) return;
    const snapshot = new Set(changedSheetIds);
    applyChain = applyChain.then(() => applyYjsSheetViewChangesToDocumentController(snapshot)).catch((err) => {
      console.error(err);
    });
  }

  /**
   * @param {any[]} deltas
   */
  function enqueueWrite(deltas) {
    const snapshot = Array.from(deltas ?? []);
    writeChain = writeChain.then(() => applyDocumentDeltas(snapshot)).catch((err) => {
      console.error(err);
    });
  }

  /**
   * @param {any[]} deltas
   */
  function enqueueSheetViewWrite(deltas) {
    const snapshot = Array.from(deltas ?? []);
    sheetWriteChain = sheetWriteChain.then(() => applyDocumentSheetViewDeltas(snapshot)).catch((err) => {
      console.error(err);
    });
  }

  /**
   * @param {any[]} deltas
   */
  function enqueueSheetFormatWrite(deltas) {
    const snapshot = Array.from(deltas ?? []);
    sheetWriteChain = sheetWriteChain.then(() => applyDocumentFormatDeltas(snapshot)).catch((err) => {
      console.error(err);
    });
  }

  /**
   * @param {any[]} deltas
   */
  function enqueueSheetRangeRunWrite(deltas) {
    const snapshot = Array.from(deltas ?? []);
    sheetWriteChain = sheetWriteChain.then(() => applyDocumentRangeRunDeltas(snapshot)).catch((err) => {
      console.error(err);
    });
  }

  /**
   * @param {string} canonicalKey
   * @param {string} rawKey
   */
  function trackRawKey(canonicalKey, rawKey) {
    let set = yjsKeysByCell.get(canonicalKey);
    if (!set) {
      set = new Set();
      yjsKeysByCell.set(canonicalKey, set);
    }
    set.add(rawKey);
  }

  /**
   * @param {string} canonicalKey
   * @param {string} rawKey
   */
  function untrackRawKey(canonicalKey, rawKey) {
    const set = yjsKeysByCell.get(canonicalKey);
    if (!set) return;
    set.delete(rawKey);
    if (set.size === 0) yjsKeysByCell.delete(canonicalKey);
  }

  /**
   * @param {string} rawKey
   * @returns {string | null}
   */
  function canonicalKeyFromRawKey(rawKey) {
    return normalizeCellKey(rawKey, { defaultSheetId });
  }

  /**
   * @param {string} canonicalKey
   * @returns {Promise<ParsedYjsCell | null>}
   */
  async function readCanonicalCellFromYjs(canonicalKey) {
    const cellRef = parseCellKey(canonicalKey, { defaultSheetId });
    if (!cellRef) return null;

    const rawKeys = yjsKeysByCell.get(canonicalKey);
    let candidates;
    if (rawKeys && rawKeys.size > 0) {
      if (rawKeys.has(canonicalKey)) {
        candidates = [canonicalKey, ...Array.from(rawKeys).filter((k) => k !== canonicalKey)];
      } else {
        candidates = rawKeys;
      }
    } else {
      candidates = [canonicalKey];
    }

    /** @type {Array<{ rawKey: string, cell: any }>} */
    const candidateCells = [];
    for (const rawKey of candidates) {
      if (typeof rawKey !== "string") continue;
      const cellData = cells.get(rawKey);
      const cell = getYMapCell(cellData);
      if (!cell) continue;
      candidateCells.push({ rawKey, cell });
    }

    if (candidateCells.length === 0) return null;

    // If any representation of this cell is encrypted, treat the cell as encrypted
    // and never fall back to plaintext duplicates (which could leak protected content
    // when legacy key encodings are present).
    const encryptedCandidates = candidateCells.filter(({ cell }) => {
      const encRaw = typeof cell.has === "function" ? (cell.has("enc") ? cell.get("enc") : undefined) : cell.get("enc");
      return encRaw !== undefined;
    });

    // Preserve format even if it's stored on a different key encoding.
    let fallbackFormat = undefined;
    let fallbackFormatKey = undefined;
    for (const { cell } of candidateCells) {
      const raw = cell.get("format");
      const rawOrStyle = raw === undefined ? cell.get("style") : raw;
      if (rawOrStyle === undefined) continue;
      fallbackFormat = rawOrStyle ?? null;
      fallbackFormatKey = stableStringify(fallbackFormat);
      break;
    }

    if (encryptedCandidates.length > 0) {
      const encryptFormat = Boolean(encryption?.encryptFormat);
      const chosen =
        encryptedCandidates.find((entry) => entry.rawKey === canonicalKey) ?? encryptedCandidates[0];
      const parsed = await readCellFromYjs(chosen.cell, cellRef, encryption, docIdForEncryption, maskFn);
      // Back-compat: when `encryptFormat` is disabled (legacy behavior), allow plaintext
      // formatting to remain readable even when value/formula are encrypted.
      if (!encryptFormat && parsed.formatKey === undefined && fallbackFormatKey !== undefined) {
        return { ...parsed, format: fallbackFormat, formatKey: fallbackFormatKey };
      }
      return parsed;
    }

    for (const { cell } of candidateCells) {
      const parsed = await readCellFromYjs(cell, cellRef, encryption, docIdForEncryption, maskFn);
      const hasData = parsed.value != null || parsed.formula != null || parsed.formatKey !== undefined;
      if (!hasData) continue;
      return parsed;
    }

    return null;
  }

  /**
   * Apply a set of canonical cell keys from Yjs into DocumentController, batching
   * into a single `applyExternalDeltas` call.
   *
   * @param {Set<string>} changedCanonicalKeys
   */
  async function applyYjsChangesToDocumentController(changedCanonicalKeys) {
    if (!changedCanonicalKeys || changedCanonicalKeys.size === 0) return;

    /** @type {any[]} */
    const deltas = [];
    /** @type {Map<string, NormalizedCell | null>} */
    const nextByKey = new Map();

    for (const canonicalKey of changedCanonicalKeys) {
      const parsed = parseCellKey(canonicalKey, { defaultSheetId });
      if (!parsed) continue;

      const before = documentController.getCell(parsed.sheetId, { row: parsed.row, col: parsed.col });
      const prev = cache.get(canonicalKey) ?? null;

      const curr = await readCanonicalCellFromYjs(canonicalKey);

      const currValue = curr?.formula ? null : (curr?.value ?? null);
      const currFormula = curr?.formula ?? null;
      const canRead = readGuard ? readGuard(parsed) : true;
      const shouldMask = !canRead && (currValue != null || currFormula != null);
      const displayValue = shouldMask ? maskFn(currFormula ?? currValue ?? null, parsed) : currValue;
      const displayFormula = shouldMask ? null : currFormula;
      const shouldMaskFormat = Boolean(
        maskCellFormat && (shouldMask || (curr && curr.maskedForEncryption === true))
      );

      let styleId = before.styleId;
      if (shouldMaskFormat) {
        styleId = 0;
      } else {
        const encryptFormat = Boolean(encryption?.encryptFormat);
        if (encryptFormat && curr?.hasEnc) {
          // When format encryption is enabled, treat per-cell formatting as available only
          // when it was decrypted successfully. Do not fall back to plaintext `format`.
          if (curr.formatKey !== undefined) {
            const format = curr.format ?? null;
            styleId = format == null ? 0 : documentController.styleTable.intern(format);
          } else {
            styleId = 0;
          }
        } else if (curr?.formatKey !== undefined) {
          const format = curr.format ?? null;
          styleId = format == null ? 0 : documentController.styleTable.intern(format);
        } else if (prev?.formatKey !== undefined) {
          // `format` key removed. Treat as explicit clear even though the key is now absent.
          styleId = 0;
        }
      }

      const next =
        curr && (currValue != null || currFormula != null || curr.formatKey !== undefined)
          ? { value: currValue, formula: currFormula, formatKey: curr.formatKey }
          : null;
      nextByKey.set(canonicalKey, next);

      // When (re)binding, `cache` starts empty but the DocumentController may
      // already have local state (e.g. a masked view before keys are granted).
      // Only use the cache short-circuit when we have a cached previous value.
      if (prev && sameNormalizedCell(prev, next)) continue;

      const after = { value: displayValue, formula: displayFormula, styleId };
      if (
        (before.value ?? null) === (after.value ?? null) &&
        (before.formula ?? null) === (after.formula ?? null) &&
        before.styleId === after.styleId
      ) {
        continue;
      }

      deltas.push({
        sheetId: parsed.sheetId,
        row: parsed.row,
        col: parsed.col,
        before,
        after,
      });
    }

    const updateCache = () => {
      for (const [canonicalKey, next] of nextByKey.entries()) {
        if (!next) cache.delete(canonicalKey);
        else cache.set(canonicalKey, next);
      }
    };

    if (deltas.length === 0) {
      updateCache();
      return;
    }

    applyingRemote = true;
    try {
      if (typeof documentController.applyExternalDeltas === "function") {
        documentController.applyExternalDeltas(deltas, { source: "collab" });
      } else {
        // Fallback for older DocumentController versions: apply via user mutations without feedback.
        const prevCanEdit = "canEditCell" in documentController ? documentController.canEditCell : undefined;
        if (prevCanEdit !== undefined) documentController.canEditCell = null;
        try {
          for (const delta of deltas) {
            if (delta.after.formula != null) {
              documentController.setCellFormula(delta.sheetId, { row: delta.row, col: delta.col }, delta.after.formula);
            } else {
              documentController.setCellValue(delta.sheetId, { row: delta.row, col: delta.col }, delta.after.value);
            }
          }
        } finally {
          if (prevCanEdit !== undefined) documentController.canEditCell = prevCanEdit;
        }
      }
    } finally {
      applyingRemote = false;
      updateCache();
    }
  }

  /**
   * Read the raw `SheetViewState` object stored in Yjs, converting nested Yjs values
   * (Y.Map/Y.Array/Y.Text) into plain JSON.
   *
   * Note: BranchService snapshots may include *format defaults* (defaultFormat/rowFormats/colFormats)
   * alongside the usual view fields. DocumentController stores layered formats outside of the view,
   * so those fields are handled separately.
   *
   * @param {any} sheetEntry
   * @returns {any | null}
   */
  function readSheetViewJsonFromYjsSheetEntry(sheetEntry) {
    /**
     * Extract only the view keys this binder actually cares about.
     *
     * Important: do not call `yjsValueToJson(rawView)` on the full object. The view
     * payload can contain large unrelated metadata (e.g. drawings, merged ranges),
     * and `view.drawings[*].id` may be an arbitrarily large Y.Text. Materializing
     * those ids can cause large allocations or DoS the binder.
     *
     * @param {any} rawView
     * @returns {any | null}
     */
    function sheetViewJsonFromRawView(rawView) {
      const viewMap = getYMap(rawView);
      if (viewMap) {
        const frozenRows = viewMap.get("frozenRows");
        const frozenCols = viewMap.get("frozenCols");
        const backgroundImageId = viewMap.get("backgroundImageId") ?? viewMap.get("background_image_id");
        const colWidths = viewMap.get("colWidths");
        const rowHeights = viewMap.get("rowHeights");
        const defaultFormat = viewMap.get("defaultFormat");
        const rowFormats = viewMap.get("rowFormats");
        const colFormats = viewMap.get("colFormats");

        if (
          frozenRows !== undefined ||
          frozenCols !== undefined ||
          backgroundImageId !== undefined ||
          colWidths !== undefined ||
          rowHeights !== undefined ||
          defaultFormat !== undefined ||
          rowFormats !== undefined ||
          colFormats !== undefined
        ) {
          return {
            frozenRows: yjsValueToJson(frozenRows) ?? 0,
            frozenCols: yjsValueToJson(frozenCols) ?? 0,
            backgroundImageId: yjsValueToJson(backgroundImageId),
            colWidths: yjsValueToJson(colWidths),
            rowHeights: yjsValueToJson(rowHeights),
            defaultFormat: yjsValueToJson(defaultFormat),
            rowFormats: yjsValueToJson(rowFormats),
            colFormats: yjsValueToJson(colFormats),
          };
        }
        return {};
      }

      if (isRecord(rawView)) {
        const frozenRows = rawView.frozenRows;
        const frozenCols = rawView.frozenCols;
        const backgroundImageId = rawView.backgroundImageId ?? rawView.background_image_id;
        const colWidths = rawView.colWidths;
        const rowHeights = rawView.rowHeights;
        const defaultFormat = rawView.defaultFormat;
        const rowFormats = rawView.rowFormats;
        const colFormats = rawView.colFormats;

        if (
          frozenRows !== undefined ||
          frozenCols !== undefined ||
          backgroundImageId !== undefined ||
          colWidths !== undefined ||
          rowHeights !== undefined ||
          defaultFormat !== undefined ||
          rowFormats !== undefined ||
          colFormats !== undefined
        ) {
          return {
            frozenRows: yjsValueToJson(frozenRows) ?? 0,
            frozenCols: yjsValueToJson(frozenCols) ?? 0,
            backgroundImageId: yjsValueToJson(backgroundImageId),
            colWidths: yjsValueToJson(colWidths),
            rowHeights: yjsValueToJson(rowHeights),
            defaultFormat: yjsValueToJson(defaultFormat),
            rowFormats: yjsValueToJson(rowFormats),
            colFormats: yjsValueToJson(colFormats),
          };
        }
        return {};
      }

      // Unknown/invalid view type. Treat as empty instead of materializing it:
      // callers only care about a small subset of keys, and converting arbitrary
      // values (e.g. huge Y.Text) can be expensive.
      return {};
    }

    const sheetMap = getYMap(sheetEntry);
    if (sheetMap) {
      // Canonical format (BranchService):
      //   sheetMap.set("view", { frozenRows, frozenCols, colWidths?, rowHeights? })
      const rawView = sheetMap.get("view");
      if (rawView !== undefined) {
        return sheetViewJsonFromRawView(rawView);
      }

      // Back-compat: legacy top-level sheet view fields.
      //
      // Officially only `frozenRows`/`frozenCols` were used outside `view`, but be generous
      // and also accept `colWidths`/`rowHeights` in case older clients stored them similarly.
      const frozenRows = sheetMap.get("frozenRows");
      const frozenCols = sheetMap.get("frozenCols");
      const backgroundImageId = sheetMap.get("backgroundImageId") ?? sheetMap.get("background_image_id");
      const colWidths = sheetMap.get("colWidths");
      const rowHeights = sheetMap.get("rowHeights");
      const defaultFormat = sheetMap.get("defaultFormat");
      const rowFormats = sheetMap.get("rowFormats");
      const colFormats = sheetMap.get("colFormats");
      if (
        frozenRows !== undefined ||
        frozenCols !== undefined ||
        backgroundImageId !== undefined ||
        colWidths !== undefined ||
        rowHeights !== undefined ||
        defaultFormat !== undefined ||
        rowFormats !== undefined ||
        colFormats !== undefined
      ) {
        return {
          frozenRows: yjsValueToJson(frozenRows) ?? 0,
          frozenCols: yjsValueToJson(frozenCols) ?? 0,
          backgroundImageId: yjsValueToJson(backgroundImageId),
          colWidths: yjsValueToJson(colWidths),
          rowHeights: yjsValueToJson(rowHeights),
          defaultFormat: yjsValueToJson(defaultFormat),
          rowFormats: yjsValueToJson(rowFormats),
          colFormats: yjsValueToJson(colFormats),
        };
      }

      return null;
    }

    // Some workflows/tests (and some historical docs) may store sheet entries as
    // plain objects inside the Y.Array rather than Y.Maps. Treat those as
    // read-only metadata and hydrate view state from them.
    if (isRecord(sheetEntry)) {
      const rawView = sheetEntry.view;
      if (rawView !== undefined) {
        return sheetViewJsonFromRawView(rawView);
      }

      const frozenRows = sheetEntry.frozenRows;
      const frozenCols = sheetEntry.frozenCols;
      const backgroundImageId = sheetEntry.backgroundImageId ?? sheetEntry.background_image_id;
      const colWidths = sheetEntry.colWidths;
      const rowHeights = sheetEntry.rowHeights;
      const defaultFormat = sheetEntry.defaultFormat;
      const rowFormats = sheetEntry.rowFormats;
      const colFormats = sheetEntry.colFormats;
      if (
        frozenRows !== undefined ||
        frozenCols !== undefined ||
        backgroundImageId !== undefined ||
        colWidths !== undefined ||
        rowHeights !== undefined ||
        defaultFormat !== undefined ||
        rowFormats !== undefined ||
        colFormats !== undefined
      ) {
        return { frozenRows, frozenCols, backgroundImageId, colWidths, rowHeights, defaultFormat, rowFormats, colFormats };
      }
    }

    return null;
  }

  /**
   * @param {any} sheetEntry
   * @returns {{ frozenRows: number, frozenCols: number, colWidths?: Record<string, number>, rowHeights?: Record<string, number> }}
   */
  function readSheetViewFromYjsSheetEntry(sheetEntry) {
    const json = readSheetViewJsonFromYjsSheetEntry(sheetEntry);
    return json ? normalizeSheetViewState(json) : emptySheetViewState();
  }

  /**
   * Extract the layered formatting defaults (sheet/row/col) encoded inside a BranchService-style
   * `SheetViewState` object.
   *
   * @param {any} sheetEntry
   * @returns {{ defaultFormat: any, rowFormats: any, colFormats: any }}
   */
  function readSheetFormatDefaultsFromYjsSheetEntry(sheetEntry) {
    const json = readSheetViewJsonFromYjsSheetEntry(sheetEntry);
    if (!isRecord(json)) return { defaultFormat: null, rowFormats: null, colFormats: null };
    return {
      defaultFormat: json.defaultFormat ?? null,
      rowFormats: json.rowFormats ?? null,
      colFormats: json.colFormats ?? null,
    };
  }

  /**
   * Read a single metadata field from a Yjs sheet entry (supports both Y.Maps and
   * plain-object sheet entries).
   *
   * @param {any} sheetEntry
   * @param {string} key
   * @returns {any}
   */
  function readSheetEntryField(sheetEntry, key) {
    const sheetMap = getYMap(sheetEntry);
    if (sheetMap) return sheetMap.get(key);
    if (isRecord(sheetEntry)) return sheetEntry[key];
    return undefined;
  }

  /**
   * @param {any} raw
   * @returns {Record<string, any> | null}
   */
  function normalizeFormatObject(raw) {
    const json = yjsValueToJson(raw);
    if (!isRecord(json)) return null;
    return json;
  }

  /**
   * @param {any} raw
   * @returns {number}
   */
  function styleIdFromYjsFormat(raw) {
    const format = normalizeFormatObject(raw);
    if (format == null) return 0;
    return documentController.styleTable.intern(format);
  }

  /**
   * Normalize a sparse row/col format map into a `Map<index, styleId>`.
   *
   * Supported encodings:
   * - Y.Map / object: `{ "12": { ...format... } }`
   * - arrays: `[{ row: 12, format: {...} }, ...]` / `[{ col: 3, format: {...} }, ...]`
   * - tuple arrays: `[[12, {...format...}], ...]`
   *
   * @param {any} raw
   * @param {"row" | "col"} axis
   * @returns {Map<number, number>}
   */
  function readSparseStyleIds(raw, axis) {
    /** @type {Map<number, number>} */
    const out = new Map();
    const json = yjsValueToJson(raw);
    if (json == null) return out;

    if (Array.isArray(json)) {
      for (const entry of json) {
        let index;
        let format;
        if (Array.isArray(entry)) {
          index = entry[0];
          format = entry[1];
        } else if (isRecord(entry)) {
          index = entry[axis] ?? entry.index;
          format = entry.format ?? entry.style ?? entry.value;
        } else {
          continue;
        }

        const idx = Number(index);
        if (!Number.isInteger(idx) || idx < 0) continue;
        const styleId = styleIdFromYjsFormat(format);
        if (styleId !== 0) out.set(idx, styleId);
      }
      return out;
    }

    if (isRecord(json)) {
      for (const [key, value] of Object.entries(json)) {
        const idx = Number(key);
        if (!Number.isInteger(idx) || idx < 0) continue;
        const styleId = styleIdFromYjsFormat(value);
        if (styleId !== 0) out.set(idx, styleId);
      }
    }

    return out;
  }

  /**
   * Compare two arrays of range-run FormatRuns (by value).
   *
   * @param {any[]} a
   * @param {any[]} b
   * @returns {boolean}
   */
  function formatRunsEqual(a, b) {
    if (a === b) return true;
    if (!Array.isArray(a) || !Array.isArray(b)) return false;
    if (a.length !== b.length) return false;
    for (let i = 0; i < a.length; i += 1) {
      const ar = a[i];
      const br = b[i];
      if (!ar || !br) return false;
      if (ar.startRow !== br.startRow) return false;
      if (ar.endRowExclusive !== br.endRowExclusive) return false;
      if (ar.styleId !== br.styleId) return false;
    }
    return true;
  }

  /**
   * Normalize sparse `formatRunsByCol` encodings from Yjs into a `Map<col, FormatRun[]>`.
   *
   * Supported encodings:
   * - Y.Map / object: `{ "12": [{ startRow, endRowExclusive, format }, ...] }`
   * - array: `[{ col: 12, runs: [...] }, ...]`
   * - tuple array: `[[12, [...]], ...]`
   *
   * Note: the returned runs are not guaranteed to be fully normalized (non-overlapping),
   * but DocumentController will normalize when applying them.
   *
   * @param {any} raw
   * @returns {Map<number, any[]>}
   */
  function readSparseFormatRunsByCol(raw) {
    /** @type {Map<number, any[]>} */
    const out = new Map();
    const json = yjsValueToJson(raw);
    if (!json) return out;

    /**
     * @param {any} colKey
     * @param {any} rawRuns
     */
    const addColRuns = (colKey, rawRuns) => {
      const col = Number(colKey);
      if (!Number.isInteger(col) || col < 0) return;

      const list = Array.isArray(rawRuns)
        ? rawRuns
        : isRecord(rawRuns)
          ? rawRuns.runs ?? rawRuns.formatRuns ?? rawRuns.segments ?? []
          : [];

      if (!Array.isArray(list) || list.length === 0) return;

      /** @type {any[]} */
      const runs = [];
      for (const entry of list) {
        if (!entry || typeof entry !== "object") continue;
        const startRow = Number(entry.startRow);
        const endRowExclusiveNum = Number(entry.endRowExclusive);
        const endRowNum = Number(entry.endRow);
        const endRowExclusive = Number.isInteger(endRowExclusiveNum)
          ? endRowExclusiveNum
          : Number.isInteger(endRowNum)
            ? endRowNum + 1
            : NaN;
        if (!Number.isInteger(startRow) || startRow < 0) continue;
        if (!Number.isInteger(endRowExclusive) || endRowExclusive <= startRow) continue;

        const styleId = styleIdFromYjsFormat(entry.format ?? entry.style ?? entry.value ?? null);
        if (styleId === 0) continue;

        runs.push({ startRow, endRowExclusive, styleId });
      }

      runs.sort((a, b) => a.startRow - b.startRow);
      if (runs.length > 0) out.set(col, runs);
    };

    if (Array.isArray(json)) {
      for (const entry of json) {
        if (Array.isArray(entry)) {
          addColRuns(entry[0], entry[1]);
          continue;
        }
        if (isRecord(entry)) {
          addColRuns(entry.col ?? entry.index, entry.runs ?? entry.formatRuns ?? entry.segments);
        }
      }
      return out;
    }

    if (isRecord(json)) {
      for (const [key, value] of Object.entries(json)) {
        addColRuns(key, value);
      }
    }

    return out;
  }

  /** 
   * @param {string} sheetId
   * @returns {Array<{ index: number, entry: any, isLocal: boolean }>}
   */
  function findYjsSheetEntriesById(sheetId) {
    if (!sheetId) return [];
    /** @type {Array<{ index: number, entry: any, isLocal: boolean }>} */
    const candidates = [];
    for (let i = 0; i < sheets.length; i += 1) {
      const entry = sheets.get(i);
      const id = coerceString(entry?.get?.("id") ?? entry?.id);
      if (id !== sheetId) continue;
      const client = entry?._item?.id?.client;
      const isLocal = typeof client === "number" && client === ydoc.clientID;
      candidates.push({ index: i, entry, isLocal });
    }

    return candidates;
  }

  /**
   * @param {string} sheetId
   * @returns {{ index: number, entry: any } | null}
   */
  function findYjsSheetEntryById(sheetId) {
    const candidates = findYjsSheetEntriesById(sheetId);

    if (candidates.length === 0) return null;
    if (candidates.length === 1) return candidates[0];

    // If we see duplicates created by the current client alongside entries from other clients,
    // prefer selecting a non-local entry. This mirrors `ensureWorkbookSchema` behavior and
    // helps avoid writing view state to a placeholder sheet that might later be pruned.
    const locals = candidates.filter((c) => c.isLocal);
    const nonLocals = candidates.filter((c) => !c.isLocal);
    const preferred = locals.length > 0 && nonLocals.length > 0 ? nonLocals : candidates;

    // Deterministic choice: pick the last matching entry by index.
    let best = preferred[0];
    for (const candidate of preferred) {
      if (candidate.index > best.index) best = candidate;
    }

    return best;
  }

  /**
   * Apply a set of sheet view updates from Yjs into DocumentController.
   *
   * @param {Set<string>} changedSheetIds
   */
  function applyYjsSheetViewChangesToDocumentController(changedSheetIds) {
    if (!changedSheetIds || changedSheetIds.size === 0) return;

    /** @type {any[]} */
    const sheetViewDeltas = [];
    /** @type {any[]} */
    const formatDeltas = [];
    /** @type {any[]} */
    const rangeRunDeltas = [];

    for (const sheetId of changedSheetIds) {
      if (!sheetId) continue;

      const found = findYjsSheetEntryById(sheetId);
      const sheetEntry = found?.entry;

      const before = documentController.getSheetView(sheetId);
      const normalizedAfter = readSheetViewFromYjsSheetEntry(sheetEntry);
      // Preserve unknown view metadata that the binder does not currently sync (e.g. merged ranges,
      // drawings). The desktop `DocumentController` stores some shared layout state alongside frozen
      // panes + axis overrides; clobbering the entire view object would silently drop that state.
      const after = { ...(before ?? {}) };
      after.frozenRows = normalizedAfter.frozenRows;
      after.frozenCols = normalizedAfter.frozenCols;
      if (normalizedAfter.backgroundImageId) {
        after.backgroundImageId = normalizedAfter.backgroundImageId;
      } else {
        delete after.backgroundImageId;
      }
      if (normalizedAfter.colWidths) {
        after.colWidths = normalizedAfter.colWidths;
      } else {
        delete after.colWidths;
      }
      if (normalizedAfter.rowHeights) {
        after.rowHeights = normalizedAfter.rowHeights;
      } else {
        delete after.rowHeights;
      }

      if (!sheetViewStateEquals(before, after)) {
        sheetViewDeltas.push({ sheetId, before, after });
      }

      // Layered formats (sheet/row/col defaults) are stored sparsely on sheet metadata.
      const sheetModel = documentController?.model?.sheets?.get?.(sheetId) ?? null;
      const beforeDefaultStyleId = Number.isInteger(sheetModel?.defaultStyleId) ? sheetModel.defaultStyleId : 0;

      // Prefer top-level sheet metadata fields when present; fall back to BranchService-style
      // encodings nested in `sheet.view`.
      const viewFormats = readSheetFormatDefaultsFromYjsSheetEntry(sheetEntry);
      const rawDefaultFormat = readSheetEntryField(sheetEntry, "defaultFormat");
      const rawRowFormats = readSheetEntryField(sheetEntry, "rowFormats");
      const rawColFormats = readSheetEntryField(sheetEntry, "colFormats");

      const afterDefaultStyleId = styleIdFromYjsFormat(
        rawDefaultFormat !== undefined ? rawDefaultFormat : viewFormats.defaultFormat,
      );
      if (beforeDefaultStyleId !== afterDefaultStyleId) {
        formatDeltas.push({
          sheetId,
          layer: "sheet",
          beforeStyleId: beforeDefaultStyleId,
          afterStyleId: afterDefaultStyleId,
        });
      }

      const beforeRowStyles = sheetModel?.rowStyleIds instanceof Map ? sheetModel.rowStyleIds : new Map();
      const beforeColStyles = sheetModel?.colStyleIds instanceof Map ? sheetModel.colStyleIds : new Map();
      const afterRowStyles = readSparseStyleIds(rawRowFormats !== undefined ? rawRowFormats : viewFormats.rowFormats, "row");
      const afterColStyles = readSparseStyleIds(rawColFormats !== undefined ? rawColFormats : viewFormats.colFormats, "col");

      const rowKeys = new Set([...beforeRowStyles.keys(), ...afterRowStyles.keys()]);
      for (const row of rowKeys) {
        const beforeStyleId = beforeRowStyles.get(row) ?? 0;
        const afterStyleId = afterRowStyles.get(row) ?? 0;
        if (beforeStyleId === afterStyleId) continue;
        formatDeltas.push({ sheetId, layer: "row", index: row, beforeStyleId, afterStyleId });
      }

      const colKeys = new Set([...beforeColStyles.keys(), ...afterColStyles.keys()]);
      for (const col of colKeys) {
        const beforeStyleId = beforeColStyles.get(col) ?? 0;
        const afterStyleId = afterColStyles.get(col) ?? 0;
        if (beforeStyleId === afterStyleId) continue;
        formatDeltas.push({ sheetId, layer: "col", index: col, beforeStyleId, afterStyleId });
      }

      // Range-run formatting (compressed rectangular formatting).
      const beforeRunsByCol = sheetModel?.formatRunsByCol instanceof Map ? sheetModel.formatRunsByCol : new Map();
      const rawRunsByCol = readSheetEntryField(sheetEntry, "formatRunsByCol");
      const viewRaw = readSheetEntryField(sheetEntry, "view");
      // Avoid `yjsValueToJson(viewRaw)`: the view object can contain large unrelated metadata
      // (e.g. drawings) and we only need `formatRunsByCol` here.
      const viewRunsByCol = (() => {
        if (viewRaw === undefined) return null;
        const viewMap = getYMap(viewRaw);
        if (viewMap) return viewMap.get("formatRunsByCol");
        if (isRecord(viewRaw)) return viewRaw.formatRunsByCol ?? null;
        return null;
      })();
      const afterRunsByCol = readSparseFormatRunsByCol(rawRunsByCol !== undefined ? rawRunsByCol : viewRunsByCol);

      const runCols = new Set([...beforeRunsByCol.keys(), ...afterRunsByCol.keys()]);
      for (const col of runCols) {
        const beforeRuns = beforeRunsByCol.get(col) ?? [];
        const afterRuns = afterRunsByCol.get(col) ?? [];
        if (formatRunsEqual(beforeRuns, afterRuns)) continue;

        let startRow = Infinity;
        let endRowExclusive = -Infinity;
        for (const run of [...beforeRuns, ...afterRuns]) {
          if (!run) continue;
          startRow = Math.min(startRow, run.startRow);
          endRowExclusive = Math.max(endRowExclusive, run.endRowExclusive);
        }
        if (startRow === Infinity) startRow = 0;
        if (endRowExclusive === -Infinity) endRowExclusive = 0;

        rangeRunDeltas.push({
          sheetId,
          col,
          startRow,
          endRowExclusive,
          beforeRuns,
          afterRuns,
        });
      }
    }

    if (sheetViewDeltas.length === 0 && formatDeltas.length === 0 && rangeRunDeltas.length === 0) return;

    applyingRemote = true;
    try {
      if (sheetViewDeltas.length > 0) {
        if (typeof documentController.applyExternalSheetViewDeltas === "function") {
          documentController.applyExternalSheetViewDeltas(sheetViewDeltas, { source: "collab" });
        } else {
          // Best-effort fallback: apply the full view state via `setFrozen`/`setColWidth`/`setRowHeight`.
          // This is not perfectly atomic but preserves reasonable behavior for older controllers.
          for (const delta of sheetViewDeltas) {
            const view = delta.after ?? emptySheetViewState();
            documentController.setFrozen(delta.sheetId, view.frozenRows, view.frozenCols);
            for (const [col, width] of Object.entries(view.colWidths ?? {})) {
              documentController.setColWidth(delta.sheetId, Number(col), width);
            }
            for (const [row, height] of Object.entries(view.rowHeights ?? {})) {
              documentController.setRowHeight(delta.sheetId, Number(row), height);
            }
          }
        }
      }

      if (formatDeltas.length > 0 && typeof documentController.applyExternalFormatDeltas === "function") {
        documentController.applyExternalFormatDeltas(formatDeltas, { source: "collab" });
      }

      if (rangeRunDeltas.length > 0 && typeof documentController.applyExternalRangeRunDeltas === "function") {
        documentController.applyExternalRangeRunDeltas(rangeRunDeltas, { source: "collab" });
      }
    } finally {
      applyingRemote = false;
    }
  }

  /**
   * One-time initial hydration of existing Yjs state into DocumentController.
   * Allowed to scan the full `cells` map exactly once at bind time.
   */
  function hydrateFromYjs() {
    /** @type {Set<string>} */
    const changedKeys = new Set();
    cells.forEach((_cellData, rawKey) => {
      const canonicalKey = canonicalKeyFromRawKey(rawKey);
      if (!canonicalKey) return;
      trackRawKey(canonicalKey, rawKey);
      changedKeys.add(canonicalKey);

      const cell = getYMapCell(_cellData);
      if (cell) {
        const encRaw = typeof cell.has === "function" ? (cell.has("enc") ? cell.get("enc") : undefined) : cell.get("enc");
        if (encRaw !== undefined) {
          hasEncryptedCells = true;
        }
      }
    });

    enqueueApply(changedKeys);
  }

  /**
   * Initial hydration of sheet view state from Yjs `sheets` metadata.
   */
  function hydrateSheetViewsFromYjs() {
    /** @type {Set<string>} */
    const changed = new Set();

    for (const entry of sheets.toArray()) {
      const id = coerceString(entry?.get?.("id") ?? entry?.id);
      if (!id) continue;
      changed.add(id);
    }

    // Hydration should not create an undo step (handled by applyExternalSheetViewDeltas).
    applyYjsSheetViewChangesToDocumentController(changed);
  }

  /**
   * @param {any[]} events
   * @param {any} transaction
   */
  const handleCellsDeepChange = (events, transaction) => {
    if (!events || events.length === 0) return;
    // Avoid processing the deep event that immediately follows a local
    // DocumentController -> Yjs write we initiated.
    //
    // Note: we intentionally do NOT blanket-suppress by `transaction.origin`
    // (even if it's "local") because CollabSession and other integrations often
    // reuse the same origin token for programmatic local mutations (e.g. branch
    // checkout/merge/restore) that *must* update the DocumentController.
    if (applyingLocal) return;

    /** @type {Set<string>} */
    const changed = new Set();

    for (const event of events) {
      // Nested map updates: event.path[0] is the root cells map key.
      const path = event?.path;
      if (Array.isArray(path) && path.length > 0) {
        const rawKey = path[0];
        if (typeof rawKey !== "string") continue;

        const changes = event?.changes?.keys;
        if (
          changes &&
          !(
            changes.has("value") ||
            changes.has("formula") ||
            changes.has("format") ||
            changes.has("style") ||
            changes.has("enc")
          )
        ) {
          continue;
        }

        const canonicalKey = canonicalKeyFromRawKey(rawKey);
        if (!canonicalKey) continue;
        trackRawKey(canonicalKey, rawKey);
        if (changes?.has("enc")) hasEncryptedCells = true;
        changed.add(canonicalKey);
        continue;
      }

      // Root map changes: keys added/updated/removed.
      const changes = event?.changes?.keys;
      if (!changes) continue;
      for (const [rawKey, change] of changes.entries()) {
        if (typeof rawKey !== "string") continue;
        const canonicalKey = canonicalKeyFromRawKey(rawKey);
        if (!canonicalKey) continue;

        if (change?.action === "delete") {
          untrackRawKey(canonicalKey, rawKey);
        } else {
          trackRawKey(canonicalKey, rawKey);
          const cellData = cells.get(rawKey);
          const cell = getYMapCell(cellData);
          const encRaw =
            cell && (typeof cell.has === "function" ? (cell.has("enc") ? cell.get("enc") : undefined) : cell.get("enc"));
          if (encRaw !== undefined) hasEncryptedCells = true;
        }

        changed.add(canonicalKey);
      }
    }

    enqueueApply(changed);
  };

  /**
   * @param {any[]} events
   * @param {any} transaction
   */
  const handleSheetsDeepChange = (events, transaction) => {
    if (!events || events.length === 0) return;
    // Avoid processing the deep event that immediately follows a local
    // DocumentController -> Yjs write we initiated.
    //
    // Note: we intentionally do NOT blanket-suppress by `transaction.origin`
    // (even if it's "local") because CollabSession and other integrations often
    // reuse the same origin token for programmatic local mutations (e.g. branch
    // checkout/merge/restore) that *must* update the DocumentController.
    if (applyingLocal) return;

    /** @type {Set<string>} */
    const changed = new Set();
    let needsFullScan = false;

    for (const event of events) {
      const path = event?.path;
      const touchesView = Array.isArray(path) && path.includes("view");
      const touchesFormats =
        Array.isArray(path) &&
        (path.includes("rowFormats") ||
          path.includes("colFormats") ||
          path.includes("defaultFormat") ||
          path.includes("formatRunsByCol"));
      const changes = event?.changes?.keys;

      // Array-level changes (insert/delete/move) don't expose meaningful `changes.keys`. This can
      // happen for the root `sheets` array, or for nested arrays under `sheet.view`.
      //
      // If the path indicates which sheet index was touched (e.g. nested view arrays),
      // we can avoid scanning the entire sheet list.
      // Note: in some Yjs versions `changes.keys` exists but is an empty Map for
      // array-level changes, so treat `size===0` as "no keys" as well.
      if (!changes || changes.size === 0) {
        if ((touchesView || touchesFormats) && Array.isArray(path) && typeof path[0] === "number") {
          const entry = sheets.get(path[0]);
          const id = coerceString(entry?.get?.("id") ?? entry?.id);
          if (id) {
            changed.add(id);
            continue;
          }
        }

        // Root array changes: conservatively rescan sheet ids.
        needsFullScan = true;
        continue;
      }

      // If the event target is a sheet map, read its id directly.
      const targetId = coerceString(event?.target?.get?.("id"));
      if (targetId) {
        changed.add(targetId);
        continue;
      }

      // Otherwise fall back to resolving the sheet index from the deep path.
      if (Array.isArray(path) && typeof path[0] === "number") {
        const entry = sheets.get(path[0]);
        const id = coerceString(entry?.get?.("id") ?? entry?.id);
        if (id) changed.add(id);
        continue;
      }

      needsFullScan = true;
    }

    if (needsFullScan) {
      for (const entry of sheets.toArray()) {
        const id = coerceString(entry?.get?.("id") ?? entry?.id);
        if (id) changed.add(id);
      }
    }

    enqueueSheetViewApply(changed);
  };

  /**
   * Apply DocumentController deltas into Yjs, encrypting cell contents when configured.
   *
   * @param {any[]} deltas
   */
  async function applyDocumentDeltas(deltas) {
    if (!deltas || deltas.length === 0) return;

    /** @type {Array<any>} */
    const prepared = [];

    for (const delta of deltas) {
      const cellRef = { sheetId: delta.sheetId, row: delta.row, col: delta.col };
      const canonicalKey = makeCellKey(cellRef);

      const value = delta.after?.value ?? null;
      const formula = normalizeFormula(delta.after?.formula ?? null);
      const beforeFormula = normalizeFormula(delta.before?.formula ?? null);
      const beforeValue = delta.before?.value ?? null;
      const styleId = Number.isInteger(delta.after?.styleId) ? delta.after.styleId : 0;
      const format = styleId === 0 ? null : documentController.styleTable.get(styleId);
      const formatKey = styleId === 0 ? undefined : stableStringify(format);

      const rawKeys = yjsKeysByCell.get(canonicalKey);
      const targets = rawKeys && rawKeys.size > 0 ? Array.from(rawKeys) : [canonicalKey];

      // Once a cell is encrypted in Yjs, we must keep it encrypted to avoid leaking
      // new plaintext writes into the shared CRDT.
      let existingEnc = false;
      for (const rawKey of targets) {
        const cellData = cells.get(rawKey);
        const cell = getYMapCell(cellData);
        if (!cell) continue;
        const hasEnc = typeof cell.has === "function" ? cell.has("enc") : cell.get("enc") !== undefined;
        if (hasEnc) {
          existingEnc = true;
          break;
        }
      }

      const key = encryption?.keyForCell?.(cellRef) ?? null;
      const shouldEncryptByConfig = encryption
        ? typeof encryption.shouldEncryptCell === "function"
          ? encryption.shouldEncryptCell(cellRef)
          : key != null
        : false;

      const wantsEncryption = existingEnc || shouldEncryptByConfig;

      let encryptedPayload = null;
      if (wantsEncryption) {
        if (!key) {
          throw new Error(`Missing encryption key for cell ${canonicalKey}`);
        }
        /** @type {any} */
        const plaintext = { value: formula != null ? null : value, formula: formula ?? null };
        // Opt-in: encrypt per-cell formatting alongside value/formula.
        if (encryption?.encryptFormat && styleId !== 0) {
          plaintext.format = format;
        }
        encryptedPayload = await encryptCellPlaintext({
          plaintext,
          key,
          context: {
            docId: docIdForEncryption,
            sheetId: cellRef.sheetId,
            row: cellRef.row,
            col: cellRef.col,
          },
        });
      }

      prepared.push({
        canonicalKey,
        targets,
        value,
        formula,
        beforeFormula,
        beforeValue,
        styleId,
        format,
        formatKey,
        encryptedPayload,
      });
    }

    // Normalize any foreign nested cell maps before we mutate them. This keeps
    // Yjs UndoManager (collab undo) reliable in mixed-module environments.
    /** @type {Set<string>} */
    const rawKeysToNormalize = new Set();
    for (const item of prepared) {
      for (const rawKey of item.targets) {
        if (typeof rawKey === "string") rawKeysToNormalize.add(rawKey);
      }
    }
    if (rawKeysToNormalize.size > 0) {
      ensureLocalCellMapsForWrite(rawKeysToNormalize);
    }

    const apply = () => {
      for (const item of prepared) {
        const { canonicalKey, targets, value, formula, beforeFormula, beforeValue, styleId, format, encryptedPayload, formatKey } =
          item;

        const beforeHadFormula = beforeFormula != null;
        const beforeHadValue = beforeValue != null;

        // Represent formula clears with explicit `null` markers so downstream
        // conflict monitors (FormulaConflictMonitor, etc) can deterministically
        // detect delete-vs-overwrite concurrency via causal Item.origin ids.
        //
        // Additionally, always clear formulas via an explicit `null` marker when
        // writing a literal value (even when `formulaConflictsMode` is off) so
        // other collaborators can reason about concurrent formula-vs-value edits.
        // This mirrors `CollabSession.setCellValue`, which always writes
        // `formula=null` rather than deleting the key.
        const shouldWriteFormulaNull =
          formula == null &&
          (value != null ||
            (formulaConflictSemanticsEnabled &&
              (valueConflictSemanticsEnabled || (beforeHadFormula && value == null))));

        // Preserve empty cell maps created by clears when we'd otherwise delete the
        // entire cell entry (which would discard the marker Items).
        const preserveEmptyCellMap =
          shouldWriteFormulaNull && value == null && formula == null && styleId === 0 && (beforeHadFormula || (valueConflictSemanticsEnabled && beforeHadValue));

        if (value == null && formula == null && styleId === 0 && !encryptedPayload && !preserveEmptyCellMap) {
          for (const rawKey of targets) {
            cells.delete(rawKey);
            untrackRawKey(canonicalKey, rawKey);
          }
          cache.delete(canonicalKey);
          continue;
        }
        for (const rawKey of targets) {
          let cellData = cells.get(rawKey);
          let cell = getYMapCell(cellData);
          if (!cell) {
            cell = new YMapCtor();
            cells.set(rawKey, cell);
          }

          const encryptFormat = Boolean(encryption?.encryptFormat);
          if (encryptedPayload) {
            cell.set("enc", encryptedPayload);
            cell.delete("value");
            cell.delete("formula");
            if (encryptFormat) {
              cell.delete("format");
              cell.delete("style");
            }
          } else {
            cell.delete("enc");
            if (formula != null) {
              cell.set("formula", formula);
              cell.set("value", null);
            } else {
              if (shouldWriteFormulaNull) {
                cell.set("formula", null);
              } else {
                cell.delete("formula");
              }
              cell.set("value", value);
            }
          }

          // When `encryptFormat` is enabled, encrypted cells must not leak plaintext `format`.
          if (!encryptedPayload || !encryptFormat) {
            // Always scrub legacy `style` to converge on the canonical `format` key.
            if (styleId === 0) {
              cell.delete("format");
              cell.delete("style");
            } else {
              cell.set("format", format);
              cell.delete("style");
            }
          } else {
            cell.delete("format");
            cell.delete("style");
          }

          cell.set("modified", Date.now());
          if (userId) cell.set("modifiedBy", userId);

          trackRawKey(canonicalKey, rawKey);
        }

        const normalizedValue = formula != null ? null : value;
        const normalizedFormula = formula != null ? formula : null;
        if (normalizedValue != null || normalizedFormula != null || formatKey !== undefined) {
          cache.set(canonicalKey, { value: normalizedValue, formula: normalizedFormula, formatKey });
        } else {
          cache.delete(canonicalKey);
        }
      }
    };

    if (typeof undoService?.transact === "function") {
      applyingLocal = true;
      try {
        undoService.transact(apply);
      } finally {
        applyingLocal = false;
      }
    } else {
      applyingLocal = true;
      try {
        ydoc.transact(apply, binderOrigin);
      } finally {
        applyingLocal = false;
      }
    }
  }

  /**
   * Apply DocumentController sheet view deltas (freeze panes, row/col sizes, etc) into Yjs.
   *
   * Canonical storage format (BranchService compatible):
   *   sheets[i].set("view", { frozenRows, frozenCols, colWidths?, rowHeights? })
   *
   * @param {any[]} deltas
   */
  function applyDocumentSheetViewDeltas(deltas) {
    if (!deltas || deltas.length === 0) return;
    if (!canWriteSharedState()) return;

    /** @type {Array<{ sheetId: string, view: any }>} */
    const prepared = [];

    for (const delta of deltas) {
      const sheetId = delta?.sheetId;
      if (typeof sheetId !== "string" || sheetId === "") continue;
      const before = delta?.before ?? emptySheetViewState();
      const after = delta?.after ?? emptySheetViewState();

      // Some DocumentController implementations store additional shared layout metadata inside
      // sheet view state (e.g. drawings, merged ranges). The binder intentionally does not
      // sync those fields. When a view delta only touches unknown keys, skip the write to
      // avoid clobbering (or redundantly rewriting) the existing Yjs `sheet.view` payload.
      const normalizedBefore = normalizeSheetViewState(before);
      const normalizedAfter = normalizeSheetViewState(after);
      if (sheetViewStateEquals(normalizedBefore, normalizedAfter)) continue;

      prepared.push({ sheetId, view: normalizedAfter });
    }

    if (prepared.length === 0) return;

    /**
     * Apply sparse axis override objects into a Y.Map-backed axis table.
     *
     * @param {any} axisMap
     * @param {Record<string, number> | undefined} overrides
     */
    const applyAxisOverridesToYMap = (axisMap, overrides) => {
      if (!axisMap) return;
      const map = getYMap(axisMap);
      if (!map) return;

      const next = overrides ?? {};
      const nextKeys = new Set(Object.keys(next));

      // Delete keys that are no longer present.
      const keysToDelete = [];
      map.forEach((_value, key) => {
        if (!nextKeys.has(key)) keysToDelete.push(key);
      });
      for (const key of keysToDelete) {
        map.delete(key);
      }

      // Upsert next overrides.
      for (const [key, value] of Object.entries(next)) {
        const prev = map.get(key);
        if (typeof prev !== "number" || Math.abs(prev - value) > 1e-6) {
          map.set(key, value);
        }
      }
    };

    const apply = () => {
      for (const item of prepared) {
        const { sheetId, view } = item;

        // If the doc temporarily contains duplicate sheet entries for the same id,
        // apply the view update to *all* matching entries. This avoids losing view
        // state if schema normalization later prunes one of the duplicates.
        let foundEntries = findYjsSheetEntriesById(sheetId);

        if (foundEntries.length === 0) {
          const sheetMap = new Y.Map();
          sheetMap.set("id", sheetId);
          sheetMap.set("name", sheetId);
          sheets.push([sheetMap]);
          foundEntries = [{ index: sheets.length - 1, entry: sheetMap, isLocal: true }];
        }

        for (const found of foundEntries) {
          let sheetMap = getYMap(found?.entry);
          if (!sheetMap) {
            sheetMap = new Y.Map();
            sheetMap.set("id", sheetId);

            const name = coerceString(found?.entry?.get?.("name") ?? found?.entry?.name) ?? sheetId;
            sheetMap.set("name", name);

            // If an existing entry was stored as a plain object, preserve any unknown
            // metadata keys by copying them into the new Y.Map.
            if (isRecord(found?.entry)) {
              // Preserve any existing view object so we can merge unknown keys (drawings, merged ranges,
              // layered format defaults, etc) when we apply the next view update below.
              //
              // Without this, upgrading a plain-object sheet entry to a Y.Map would drop any view-scoped
              // metadata that this binder does not explicitly manage.
              if (Object.prototype.hasOwnProperty.call(found.entry, "view")) {
                const rawView = found.entry.view;
                if (rawView !== undefined) {
                  sheetMap.set("view", sanitizeSheetViewForPreservation(rawView));
                }
              }

              const keys = Object.keys(found.entry).sort();
              for (const key of keys) {
                if (
                  key === "id" ||
                  key === "name" ||
                  key === "view" ||
                  key === "frozenRows" ||
                  key === "frozenCols" ||
                  key === "colWidths" ||
                  key === "rowHeights"
                ) {
                  continue;
                }
                const value = found.entry[key];
                if (key === "drawings") {
                  const sanitized = sanitizeDrawingsForPreservation(value);
                  if (Array.isArray(sanitized) && sanitized.length === 0) continue;
                  sheetMap.set(key, sanitized);
                  continue;
                }
                try {
                  sheetMap.set(key, structuredClone(value));
                } catch {
                  sheetMap.set(key, value);
                }
              }
            }

            sheets.delete(found.index, 1);
            sheets.insert(found.index, [sheetMap]);
          }

          const existingView = sheetMap.get("view");
          const existingViewMap = getYMap(existingView);
          if (existingViewMap) {
            // When another binder (e.g. the desktop sheet-view binder) stores `sheet.view` as a Y.Map,
            // update keys in-place instead of overwriting the whole value. This avoids:
            // - clobbering unknown keys (drawings, merged ranges, etc)
            // - rewriting large view payloads when only a single field changes
            // - toggling between Y.Map and plain-object representations when multiple binders are attached
            if (existingViewMap.get("frozenRows") !== view.frozenRows) existingViewMap.set("frozenRows", view.frozenRows);
            if (existingViewMap.get("frozenCols") !== view.frozenCols) existingViewMap.set("frozenCols", view.frozenCols);

            const nextBg = normalizeOptionalId(view.backgroundImageId);
            if (nextBg) {
              const prevBg = normalizeOptionalId(existingViewMap.get("backgroundImageId"));
              if (prevBg !== nextBg) existingViewMap.set("backgroundImageId", nextBg);
            } else {
              existingViewMap.delete("backgroundImageId");
            }
            // Converge legacy key to the canonical one (even if the value is unchanged).
            existingViewMap.delete("background_image_id");

            const colWidthsRaw = existingViewMap.get("colWidths");
            const rowHeightsRaw = existingViewMap.get("rowHeights");
            const colWidthsMap = getYMap(colWidthsRaw);
            const rowHeightsMap = getYMap(rowHeightsRaw);

            if (colWidthsMap) {
              applyAxisOverridesToYMap(colWidthsMap, view.colWidths);
            } else if (view.colWidths) {
              existingViewMap.set("colWidths", view.colWidths);
            } else if (colWidthsRaw !== undefined) {
              existingViewMap.delete("colWidths");
            }

            if (rowHeightsMap) {
              applyAxisOverridesToYMap(rowHeightsMap, view.rowHeights);
            } else if (view.rowHeights) {
              existingViewMap.set("rowHeights", view.rowHeights);
            } else if (rowHeightsRaw !== undefined) {
              existingViewMap.delete("rowHeights");
            }

            continue;
          }

          const nextView = { ...view };

          // Preserve any unknown keys stored alongside view state (e.g. BranchService-style
          // layered formatting defaults: defaultFormat/rowFormats/colFormats) so view-only
          // interactions (freeze panes, resizing) do not accidentally wipe them.
          if (existingView !== undefined) {
            const json = yjsValueToJson(existingView);
            const sanitized = sanitizeSheetViewForPreservation(json);
            if (isRecord(sanitized)) {
              for (const [key, value] of Object.entries(sanitized)) {
                if (
                  key === "frozenRows" ||
                  key === "frozenCols" ||
                  key === "backgroundImageId" ||
                  key === "background_image_id" ||
                  key === "colWidths" ||
                  key === "rowHeights"
                ) {
                  continue;
                }
                if (value === undefined) continue;
                nextView[key] = value;
              }
            }
          }

          sheetMap.set("view", nextView);
        }
      }
    };

    if (typeof undoService?.transact === "function") {
      const prevApplyingLocal = applyingLocal;
      applyingLocal = true;
      try {
        undoService.transact(apply);
      } finally {
        applyingLocal = prevApplyingLocal;
      }
    } else {
      const prevApplyingLocal = applyingLocal;
      applyingLocal = true;
      try {
        ydoc.transact(apply, binderOrigin);
      } finally {
        applyingLocal = prevApplyingLocal;
      }
    }
  }

  /**
   * Apply DocumentController layered format deltas (sheet/row/col defaults) into Yjs.
   *
   * Canonical storage format:
   *   sheets[i].set("defaultFormat", styleObject)
   *   sheets[i].set("rowFormats", Y.Map<string,rowIndex> -> styleObject)
   *   sheets[i].set("colFormats", Y.Map<string,colIndex> -> styleObject)
   *
   * Values are stored sparsely:
   * - missing key means "no default" (style id 0)
   * - row/col keys are deleted when clearing
   * - empty rowFormats/colFormats maps are intentionally kept (rather than deleting the key)
   *   to avoid falling back to stale legacy `view.rowFormats`/`view.colFormats` snapshots.
   *
   * @param {any[]} deltas
   */
  function applyDocumentFormatDeltas(deltas) {
    if (!deltas || deltas.length === 0) return;
    if (!canWriteSharedState()) return;

    /** @type {Array<{ sheetId: string, layer: string, index?: number, style: any | null }>} */
    const prepared = [];

    for (const delta of deltas) {
      const sheetId = delta?.sheetId;
      if (typeof sheetId !== "string" || sheetId === "") continue;
      const layer = delta?.layer;
      if (layer !== "sheet" && layer !== "row" && layer !== "col") continue;

      const afterStyleId = Number.isInteger(delta?.afterStyleId) ? delta.afterStyleId : 0;
      const style = afterStyleId === 0 ? null : documentController.styleTable.get(afterStyleId);

      if (layer === "row" || layer === "col") {
        const index = Number(delta?.index);
        if (!Number.isInteger(index) || index < 0) continue;
        prepared.push({ sheetId, layer, index, style });
      } else {
        prepared.push({ sheetId, layer, style });
      }
    }

    if (prepared.length === 0) return;

    /**
     * If we see duplicate sheet entries for the same id (can happen while two clients
     * concurrently initialize an empty workbook), apply format updates to *all* matching
     * entries so schema normalization doesn't drop the newly written formatting.
     *
     * @param {string} sheetId
     * @returns {Y.Map<any>[]}
     */
    const ensureYjsSheetMaps = (sheetId) => {
      let foundEntries = findYjsSheetEntriesById(sheetId);

      if (foundEntries.length === 0) {
        const sheetMap = new Y.Map();
        sheetMap.set("id", sheetId);
        sheetMap.set("name", sheetId);
        sheets.push([sheetMap]);
        foundEntries = [{ index: sheets.length - 1, entry: sheetMap, isLocal: true }];
      }

      /** @type {Y.Map<any>[]} */
      const out = [];

      for (const found of foundEntries) {
        let sheetMap = getYMap(found?.entry);
        if (!sheetMap) {
          sheetMap = new Y.Map();
          sheetMap.set("id", sheetId);

          const name = coerceString(found?.entry?.get?.("name") ?? found?.entry?.name) ?? sheetId;
          sheetMap.set("name", name);

           if (isRecord(found?.entry)) {
             const keys = Object.keys(found.entry).sort();
             for (const key of keys) {
               if (key === "id" || key === "name") continue;
               const value = found.entry[key];
               if (key === "view") {
                 sheetMap.set(key, sanitizeSheetViewForPreservation(value));
                 continue;
               }
               if (key === "drawings") {
                 const sanitized = sanitizeDrawingsForPreservation(value);
                 if (Array.isArray(sanitized) && sanitized.length === 0) continue;
                 sheetMap.set(key, sanitized);
                 continue;
               }
               try {
                 sheetMap.set(key, structuredClone(value));
               } catch {
                 sheetMap.set(key, value);
              }
            }
          }

          sheets.delete(found.index, 1);
          sheets.insert(found.index, [sheetMap]);
        }

        out.push(sheetMap);
      }

      return out;
    };

    /**
     * @param {string} sheetId
     * @param {"row" | "col"} layer
     * @returns {Map<number, number>}
     */
    const getLayerStyleIds = (sheetId, layer) => {
      const sheet = documentController?.model?.sheets?.get?.(sheetId) ?? null;
      const styles = layer === "row" ? sheet?.rowStyleIds : sheet?.colStyleIds;
      return styles instanceof Map ? styles : new Map();
    };

    /** @type {Map<string, Array<{ sheetId: string, layer: string, index?: number, style: any | null }>>} */
    const preparedBySheet = new Map();
    for (const item of prepared) {
      let list = preparedBySheet.get(item.sheetId);
      if (!list) {
        list = [];
        preparedBySheet.set(item.sheetId, list);
      }
      list.push(item);
    }

    const apply = () => {
      for (const [sheetId, items] of preparedBySheet.entries()) {
        const sheetMaps = ensureYjsSheetMaps(sheetId);

        // If the sheet metadata does not yet have top-level format maps (e.g. legacy docs
        // stored them inside `view`), seed them from the DocumentController's current
        // state so future reads don't fall back to stale view-encoded values.
        const seedRowStyles = getLayerStyleIds(sheetId, "row");
        const seedColStyles = getLayerStyleIds(sheetId, "col");

        for (const sheetMap of sheetMaps) {
          const ensureLayerMap = (mapKey, seedStyles) => {
            const existing = sheetMap.get(mapKey);
            const existingMap = getYMap(existing);
            if (existingMap) return existingMap;

            const next = new Y.Map();

            // If a previous client stored this map as a plain object, upgrade it.
            if (isRecord(existing)) {
              const keys = Object.keys(existing).sort();
              for (const k of keys) {
                next.set(k, existing[k]);
              }
            } else if (existing === undefined) {
              // Legacy migration: seed from DocumentController.
              for (const [idx, styleId] of seedStyles.entries()) {
                const style = documentController.styleTable.get(styleId);
                next.set(String(idx), style);
              }
            }

            sheetMap.set(mapKey, next);
            return next;
          };

          for (const item of items) {
            const { layer, index, style } = item;

            if (layer === "sheet") {
              // Always store an explicit value (style object or null) so top-level
              // storage overrides any legacy `view.defaultFormat`.
              sheetMap.set("defaultFormat", style == null ? null : style);
              continue;
            }

            const mapKey = layer === "row" ? "rowFormats" : "colFormats";
            const idxKey = String(index);
            const seedStyles = layer === "row" ? seedRowStyles : seedColStyles;
            const ensured = getYMap(sheetMap.get(mapKey)) ?? ensureLayerMap(mapKey, seedStyles);

            if (style == null) {
              ensured.delete(idxKey);
              // Keep an empty Y.Map rather than deleting the key. This prevents falling
              // back to stale legacy `view.rowFormats`/`view.colFormats` data.
              continue;
            }

            ensured.set(idxKey, style);
          }
        }
      }
    };

    if (typeof undoService?.transact === "function") {
      const prevApplyingLocal = applyingLocal;
      applyingLocal = true;
      try {
        undoService.transact(apply);
      } finally {
        applyingLocal = prevApplyingLocal;
      }
    } else {
      const prevApplyingLocal = applyingLocal;
      applyingLocal = true;
      try {
        ydoc.transact(apply, binderOrigin);
      } finally {
        applyingLocal = prevApplyingLocal;
      }
    }
  }

  /**
   * Apply DocumentController range-run deltas (compressed rectangular formatting) into Yjs.
   *
   * Storage format:
   *   sheets[i].get("formatRunsByCol") is a `Y.Map<string, any>` where each key is a
   *   0-based column index string and each value is an array of runs:
   *     [{ startRow, endRowExclusive, format }, ...]
   *
   * Runs store style objects (not style ids) because style ids are local to each
   * DocumentController instance.
   *
   * @param {any[]} deltas
   */
  function applyDocumentRangeRunDeltas(deltas) {
    if (!deltas || deltas.length === 0) return;
    if (!canWriteSharedState()) return;

    /**
     * @type {Map<string, Array<{ col: number, runs: any[] }>>}
     */
    const preparedBySheet = new Map();

    const cloneJson = (value) => {
      try {
        return structuredClone(value);
      } catch {
        return value;
      }
    };

    for (const delta of deltas) {
      const sheetId = delta?.sheetId;
      if (typeof sheetId !== "string" || sheetId === "") continue;
      const col = Number(delta?.col);
      if (!Number.isInteger(col) || col < 0) continue;

      const afterRuns = Array.isArray(delta?.afterRuns) ? delta.afterRuns : [];
      const runs = afterRuns
        .map((run) => {
          const startRow = Number(run?.startRow);
          const endRowExclusive = Number(run?.endRowExclusive);
          const styleId = Number(run?.styleId);
          if (!Number.isInteger(startRow) || startRow < 0) return null;
          if (!Number.isInteger(endRowExclusive) || endRowExclusive <= startRow) return null;
          if (!Number.isInteger(styleId) || styleId <= 0) return null;
          return {
            startRow,
            endRowExclusive,
            format: cloneJson(documentController.styleTable.get(styleId)),
          };
        })
        .filter(Boolean);

      let list = preparedBySheet.get(sheetId);
      if (!list) {
        list = [];
        preparedBySheet.set(sheetId, list);
      }
      list.push({ col, runs });
    }

    if (preparedBySheet.size === 0) return;

    /**
     * @param {Y.Map<any>} sheetMap
     * @param {string} sheetId
     * @returns {Y.Map<any>}
     */
    const ensureFormatRunsMap = (sheetMap, sheetId) => {
      const existing = sheetMap.get("formatRunsByCol");
      const existingMap = getYMap(existing);
      if (existingMap) return existingMap;

      const next = new Y.Map();

      if (isRecord(existing)) {
        // Upgrade plain-object encodings.
        const keys = Object.keys(existing).sort();
        for (const key of keys) {
          next.set(key, cloneJson(existing[key]));
        }
      } else if (existing === undefined) {
        // Legacy migration: seed from the DocumentController's current range-run state so
        // top-level storage becomes authoritative (overriding any legacy `view.formatRunsByCol`).
        const sheet = documentController?.model?.sheets?.get?.(sheetId) ?? null;
        const runsByCol = sheet?.formatRunsByCol;
        if (runsByCol instanceof Map) {
          for (const [col, runs] of runsByCol.entries()) {
            if (!Array.isArray(runs) || runs.length === 0) continue;
            const serialized = runs
              .map((run) => {
                const startRow = Number(run?.startRow);
                const endRowExclusive = Number(run?.endRowExclusive);
                const styleId = Number(run?.styleId);
                if (!Number.isInteger(startRow) || startRow < 0) return null;
                if (!Number.isInteger(endRowExclusive) || endRowExclusive <= startRow) return null;
                if (!Number.isInteger(styleId) || styleId <= 0) return null;
                return {
                  startRow,
                  endRowExclusive,
                  format: cloneJson(documentController.styleTable.get(styleId)),
                };
              })
              .filter(Boolean);
            if (serialized.length > 0) next.set(String(col), serialized);
          }
        }
      }

      sheetMap.set("formatRunsByCol", next);
      return next;
    };

    const apply = () => {
      for (const [sheetId, items] of preparedBySheet.entries()) {
        // Apply to all matching entries when duplicate sheet ids exist.
        let foundEntries = findYjsSheetEntriesById(sheetId);
        if (foundEntries.length === 0) {
          const sheetMap = new Y.Map();
          sheetMap.set("id", sheetId);
          sheetMap.set("name", sheetId);
          sheets.push([sheetMap]);
          foundEntries = [{ index: sheets.length - 1, entry: sheetMap, isLocal: true }];
        }

        for (const found of foundEntries) {
          let sheetMap = getYMap(found?.entry);
          if (!sheetMap) {
            sheetMap = new Y.Map();
            sheetMap.set("id", sheetId);
            const name = coerceString(found?.entry?.get?.("name") ?? found?.entry?.name) ?? sheetId;
            sheetMap.set("name", name);

            if (isRecord(found?.entry)) {
              const keys = Object.keys(found.entry).sort();
              for (const key of keys) {
                if (key === "id" || key === "name") continue;
                const value = found.entry[key];
                if (key === "view") {
                  sheetMap.set(key, sanitizeSheetViewForPreservation(value));
                  continue;
                }
                if (key === "drawings") {
                  const sanitized = sanitizeDrawingsForPreservation(value);
                  if (Array.isArray(sanitized) && sanitized.length === 0) continue;
                  sheetMap.set(key, sanitized);
                  continue;
                }
                sheetMap.set(key, cloneJson(value));
              }
            }

            sheets.delete(found.index, 1);
            sheets.insert(found.index, [sheetMap]);
          }

          for (const item of items) {
            const colKey = String(item.col);
            const map = getYMap(sheetMap.get("formatRunsByCol")) ?? ensureFormatRunsMap(sheetMap, sheetId);

            if (item.runs.length === 0) {
              map.delete(colKey);
              // Keep an empty Y.Map rather than deleting the key. This prevents falling back
              // to stale legacy `view.formatRunsByCol` data.
              continue;
            }

            map.set(colKey, item.runs);
          }
        }
      }
    };

    if (typeof undoService?.transact === "function") {
      const prevApplyingLocal = applyingLocal;
      applyingLocal = true;
      try {
        undoService.transact(apply);
      } finally {
        applyingLocal = prevApplyingLocal;
      }
    } else {
      const prevApplyingLocal = applyingLocal;
      applyingLocal = true;
      try {
        ydoc.transact(apply, binderOrigin);
      } finally {
        applyingLocal = prevApplyingLocal;
      }
    }
  }

  /**
   * Determine whether a local DocumentController change should be allowed to propagate
   * into Yjs when encryption is enabled (or encrypted cells exist in the doc).
   *
   * This prevents cases where a collaborator without the relevant key edits an
   * encrypted cell locally (or where `shouldEncryptCell` requires encryption but
   * no key is available). In those cases we revert the local edit, similar to the
   * permissions `canEditCell` guard.
   *
   * @param {{ sheetId: string, row: number, col: number }} cellRef
   */
  function canWriteCellWithEncryption(cellRef) {
    const canonicalKey = makeCellKey(cellRef);
    const rawKeys = yjsKeysByCell.get(canonicalKey);
    const targets = rawKeys && rawKeys.size > 0 ? Array.from(rawKeys) : [canonicalKey];

    let existingEnc = false;
    for (const rawKey of targets) {
      const cellData = cells.get(rawKey);
      const cell = getYMapCell(cellData);
      if (!cell) continue;
      const encRaw = typeof cell.has === "function" ? (cell.has("enc") ? cell.get("enc") : undefined) : cell.get("enc");
      if (encRaw !== undefined) {
        existingEnc = true;
        break;
      }
    }

    const key = encryption?.keyForCell?.(cellRef) ?? null;
    const shouldEncryptByConfig = encryption
      ? typeof encryption.shouldEncryptCell === "function"
        ? encryption.shouldEncryptCell(cellRef)
        : key != null
      : false;

    const wantsEncryption = existingEnc || shouldEncryptByConfig;
    if (!wantsEncryption) return true;

    // If the cell is encrypted (or must be encrypted by config) we need a key to
    // avoid writing plaintext into the shared CRDT.
    if (!key) return false;

    return true;
  }

  /**
   * Determine whether a layered format edit should be allowed to propagate into Yjs.
   *
   * For role-based permissions, we conservatively reject sheet/row/col style edits if they
   * intersect *any* restricted range the user cannot edit.
   *
   * @param {any} delta
   * @returns {boolean}
   */
  function canEditFormatDelta(delta) {
    if (!editGuard) return true;
    const sheetId = typeof delta?.sheetId === "string" ? delta.sheetId : null;
    if (!sheetId) return false;

    const layer = delta?.layer;
    if (layer !== "sheet" && layer !== "row" && layer !== "col") return false;

    const indexRaw = delta?.index;
    const index = Number(indexRaw);
    if ((layer === "row" || layer === "col") && (!Number.isInteger(index) || index < 0)) return false;

    const restrictions =
      permissions && typeof permissions === "object" && Array.isArray(permissions.restrictions) ? permissions.restrictions : null;

    // Role-based restrictions: ensure we don't touch any protected ranges.
    if (restrictions && restrictions.length > 0) {
      for (const restriction of restrictions) {
        const range = restriction?.range ?? restriction;
        const restrictionSheetId = coerceString(
          range?.sheetId ?? range?.sheetName ?? restriction?.sheetId ?? restriction?.sheetName,
        );
        if (restrictionSheetId && restrictionSheetId !== sheetId) continue;

        const startRow = Number(range?.startRow);
        const endRow = Number(range?.endRow);
        const startCol = Number(range?.startCol);
        const endCol = Number(range?.endCol);
        if (!Number.isInteger(startRow) || !Number.isInteger(endRow) || startRow < 0 || endRow < startRow) continue;
        if (!Number.isInteger(startCol) || !Number.isInteger(endCol) || startCol < 0 || endCol < startCol) continue;

        let intersects = true;
        if (layer === "row") intersects = index >= startRow && index <= endRow;
        else if (layer === "col") intersects = index >= startCol && index <= endCol;
        if (!intersects) continue;

        const probe = {
          sheetId,
          row: layer === "row" ? index : startRow,
          col: layer === "col" ? index : startCol,
        };
        if (!editGuard(probe)) return false;
      }
      return true;
    }

    // Fallback for custom permission functions: probe a representative cell.
    if (layer === "row") return editGuard({ sheetId, row: index, col: 0 });
    if (layer === "col") return editGuard({ sheetId, row: 0, col: index });
    return editGuard({ sheetId, row: 0, col: 0 });
  }

  /**
   * Determine whether a range-run formatting edit should be allowed to propagate into Yjs.
   *
   * @param {any} delta
   * @returns {boolean}
   */
  function canEditRangeRunDelta(delta) {
    if (!editGuard) return true;
    const sheetId = typeof delta?.sheetId === "string" ? delta.sheetId : null;
    if (!sheetId) return false;

    const col = Number(delta?.col);
    const startRow = Number(delta?.startRow);
    const endRowExclusive = Number(delta?.endRowExclusive);
    if (!Number.isInteger(col) || col < 0) return false;
    if (!Number.isInteger(startRow) || startRow < 0) return false;
    if (!Number.isInteger(endRowExclusive) || endRowExclusive <= startRow) return false;

    const restrictions =
      permissions && typeof permissions === "object" && Array.isArray(permissions.restrictions) ? permissions.restrictions : null;

    // Role-based restrictions: ensure we don't touch any protected ranges.
    if (restrictions && restrictions.length > 0) {
      for (const restriction of restrictions) {
        const range = restriction?.range ?? restriction;
        const restrictionSheetId = coerceString(
          range?.sheetId ?? range?.sheetName ?? restriction?.sheetId ?? restriction?.sheetName,
        );
        if (restrictionSheetId && restrictionSheetId !== sheetId) continue;

        const startRowR = Number(range?.startRow);
        const endRowR = Number(range?.endRow);
        const startColR = Number(range?.startCol);
        const endColR = Number(range?.endCol);
        if (!Number.isInteger(startRowR) || !Number.isInteger(endRowR) || startRowR < 0 || endRowR < startRowR) continue;
        if (!Number.isInteger(startColR) || !Number.isInteger(endColR) || startColR < 0 || endColR < startColR) continue;

        const intersectsCol = col >= startColR && col <= endColR;
        const intersectsRow = startRow <= endRowR && endRowExclusive - 1 >= startRowR;
        if (!intersectsCol || !intersectsRow) continue;

        // Probe any cell in the overlap.
        const probeRow = Math.max(startRow, startRowR);
        if (!editGuard({ sheetId, row: probeRow, col })) return false;
      }
      return true;
    }

    // Fallback for custom permission functions: probe the start cell.
    return editGuard({ sheetId, row: startRow, col });
  }

  /**
   * Extract layered-format deltas from a DocumentController change payload.
   *
   * Prefer the unified `payload.formatDeltas` stream, but accept the split
   * `rowStyleDeltas`/`colStyleDeltas`/`sheetStyleDeltas` variants for compatibility.
   *
   * @param {any} payload
   * @returns {any[]}
   */
  function readFormatDeltasFromDocumentChange(payload) {
    const formatDeltas = Array.isArray(payload?.formatDeltas) ? payload.formatDeltas : [];
    if (formatDeltas.length > 0) return formatDeltas;

    /** @type {any[]} */
    const out = [];

    const sheetStyleDeltas = Array.isArray(payload?.sheetStyleDeltas)
      ? payload.sheetStyleDeltas
      : Array.isArray(payload?.sheetStyleIdDeltas)
        ? payload.sheetStyleIdDeltas
        : [];
    for (const delta of sheetStyleDeltas) {
      if (!delta) continue;
      const sheetId = delta.sheetId;
      if (typeof sheetId !== "string" || sheetId === "") continue;
      out.push({
        sheetId,
        layer: "sheet",
        beforeStyleId: delta.beforeStyleId,
        afterStyleId: delta.afterStyleId,
      });
    }

    const rowStyleDeltas = Array.isArray(payload?.rowStyleDeltas)
      ? payload.rowStyleDeltas
      : Array.isArray(payload?.rowStyleIdDeltas)
        ? payload.rowStyleIdDeltas
        : [];
    for (const delta of rowStyleDeltas) {
      if (!delta) continue;
      const sheetId = delta.sheetId;
      const row = delta.row;
      if (typeof sheetId !== "string" || sheetId === "") continue;
      if (!Number.isInteger(row) || row < 0) continue;
      out.push({
        sheetId,
        layer: "row",
        index: row,
        beforeStyleId: delta.beforeStyleId,
        afterStyleId: delta.afterStyleId,
      });
    }

    const colStyleDeltas = Array.isArray(payload?.colStyleDeltas)
      ? payload.colStyleDeltas
      : Array.isArray(payload?.colStyleIdDeltas)
        ? payload.colStyleIdDeltas
        : [];
    for (const delta of colStyleDeltas) {
      if (!delta) continue;
      const sheetId = delta.sheetId;
      const col = delta.col;
      if (typeof sheetId !== "string" || sheetId === "") continue;
      if (!Number.isInteger(col) || col < 0) continue;
      out.push({
        sheetId,
        layer: "col",
        index: col,
        beforeStyleId: delta.beforeStyleId,
        afterStyleId: delta.afterStyleId,
      });
    }

    return out;
  }

  /**
   * @param {any} payload
   */
  const handleDocumentChange = (payload) => {
    if (applyingRemote) return;

    // Desktop wires multiple binders to the same DocumentController (e.g. a lightweight sheet-view
    // binder for drawings/merged ranges alongside this full binder). Those other binders may apply
    // remote Yjs updates into the DocumentController using `applyExternal*` methods, which emit
    // `change` events tagged with `payload.source === "collab"`.
    //
    // Treat those as remote/external and do not write them back into Yjs (avoids redundant Yjs
    // updates and prevents collaborative undo from incorrectly tracking remote edits as local).
    //
    // Similarly, snapshot restores (`applyState`) should not implicitly overwrite the shared Yjs
    // state through the binder.
    const source = typeof payload?.source === "string" ? payload.source : null;
    if (source === "collab" || source === "applyState") return;

    const allowSharedStateWrites = canWriteSharedState();

    const deltas = Array.isArray(payload?.deltas) ? payload.deltas : [];
    const sheetViewDeltas = Array.isArray(payload?.sheetViewDeltas) ? payload.sheetViewDeltas : [];
    const formatDeltasRaw = readFormatDeltasFromDocumentChange(payload);
    const rangeRunDeltasRaw = Array.isArray(payload?.rangeRunDeltas) ? payload.rangeRunDeltas : [];

    // In read-only roles, do not persist shared-state mutations (sheet view / formatting) into Yjs.
    //
    // Note: Unlike cell edits, some shared-state mutations are *intentionally* allowed to remain
    // local-only (e.g. freeze panes, row/col sizing, and formatting defaults applied to full
    // row/col/sheet selections). UIs are responsible for gating these interactions. The binder
    // simply prevents them from being written into shared Yjs state.
    const formatDeltas = allowSharedStateWrites ? formatDeltasRaw : [];
    const rangeRunDeltas = allowSharedStateWrites ? rangeRunDeltasRaw : [];

    if (allowSharedStateWrites && sheetViewDeltas.length > 0) {
      enqueueSheetViewWrite(sheetViewDeltas);
    }

    const stripAllowedReadOnlySheetViewFields = (view) => {
      if (!isRecord(view)) return {};
      /** @type {Record<string, any>} */
      const out = {};
      for (const [key, value] of Object.entries(view)) {
        // Allow these view-only fields to remain local-only in read-only roles.
        if (key === "frozenRows" || key === "frozenCols" || key === "colWidths" || key === "rowHeights") continue;
        // Treat `undefined` as absent so missing-vs-undefined does not look like a change.
        if (value === undefined) continue;
        out[key] = value;
      }
      return out;
    };
    const isAllowedReadOnlySheetViewDelta = (delta) => {
      const before = stripAllowedReadOnlySheetViewFields(delta?.before);
      const after = stripAllowedReadOnlySheetViewFields(delta?.after);
      return stableStringify(before) === stableStringify(after);
    };

    const normalizeRunList = (runs) => {
      if (!Array.isArray(runs) || runs.length === 0) return [];
      /** @type {Array<{ startRow: number, endRowExclusive: number }>} */
      const out = [];
      for (const run of runs) {
        const startRow = Number(run?.startRow);
        const endRowExclusive = Number(run?.endRowExclusive);
        if (!Number.isInteger(startRow) || !Number.isInteger(endRowExclusive)) continue;
        if (endRowExclusive <= startRow) continue;
        out.push({ startRow, endRowExclusive });
      }
      out.sort((a, b) => a.startRow - b.startRow);
      return out;
    };
    /**
     * Return true when the "after" run coverage is a subset of the "before" coverage (i.e. the
     * delta only patches existing range-run formatting and does not introduce formatting in gaps).
     *
     * This matches `DocumentController`'s `patchExistingFormatRuns` behavior used when applying
     * full-row/full-col/full-sheet formatting defaults: range-runs have higher precedence than
     * row/col/sheet defaults, so the controller patches existing runs but avoids creating new ones.
     */
    const isRangeRunCoverageSubset = (beforeRuns, afterRuns) => {
      const before = normalizeRunList(beforeRuns);
      const after = normalizeRunList(afterRuns);
      if (after.length === 0) return true;
      if (before.length === 0) return false;

      let bi = 0;
      for (const a of after) {
        let cursor = a.startRow;
        const end = a.endRowExclusive;
        while (cursor < end) {
          while (bi < before.length && before[bi].endRowExclusive <= cursor) bi += 1;
          const b = before[bi];
          // Gap or no covering run.
          if (!b || b.startRow > cursor) return false;
          cursor = Math.min(end, b.endRowExclusive);
        }
      }
      return true;
    };

    /** @type {any[]} */
    const deniedInverseSheetViews = [];
    /** @type {any[]} */
    const deniedInverseFormats = [];
    /** @type {any[]} */
    const deniedInverseRangeRuns = [];
    if (!allowSharedStateWrites) {
      for (const delta of sheetViewDeltas) {
        // Only revert sheet-view keys that we do *not* support as local-only view state.
        // Freeze panes + row/col sizing are allowed to diverge locally.
        if (isAllowedReadOnlySheetViewDelta(delta)) continue;
        const sheetId = typeof delta?.sheetId === "string" ? delta.sheetId : null;
        if (!sheetId) continue;
        deniedInverseSheetViews.push({
          sheetId,
          before: delta.after,
          after: delta.before,
        });
      }
      // Layered formatting defaults (sheet/row/col) are allowed to be local-only in read-only roles.
      for (const delta of rangeRunDeltasRaw) {
        const sheetId = typeof delta?.sheetId === "string" ? delta.sheetId : null;
        if (!sheetId) continue;
        const col = Number(delta?.col);
        const startRow = Number(delta?.startRow);
        const endRowExclusive = Number(delta?.endRowExclusive);
        if (!Number.isInteger(col) || col < 0) continue;
        if (!Number.isInteger(startRow) || startRow < 0) continue;
        if (!Number.isInteger(endRowExclusive) || endRowExclusive <= startRow) continue;
        // When applying full-row/full-col/full-sheet formatting defaults, DocumentController may
        // patch existing range-run formatting (so defaults win over higher-precedence range runs)
        // without creating new run segments in gaps. Allow those patches to remain local-only.
        //
        // If a delta would *introduce* new run coverage (e.g. formatting a rectangular range stored
        // as range runs), revert it to avoid enabling arbitrary cell formatting in read-only roles.
        if (isRangeRunCoverageSubset(delta.beforeRuns, delta.afterRuns)) continue;
        deniedInverseRangeRuns.push({
          sheetId,
          col,
          startRow,
          endRowExclusive,
          beforeRuns: delta.afterRuns,
          afterRuns: delta.beforeRuns,
        });
      }
    }

    const needsSharedStateRevert =
      deniedInverseSheetViews.length > 0 || deniedInverseFormats.length > 0 || deniedInverseRangeRuns.length > 0;

    if (deltas.length === 0 && formatDeltas.length === 0 && rangeRunDeltas.length === 0 && !needsSharedStateRevert) return;

    const needsEncryptionGuard = Boolean(encryption || hasEncryptedCells);
    if (!editGuard && !needsEncryptionGuard) {
      if (formatDeltas.length > 0) enqueueSheetFormatWrite(formatDeltas);
      if (rangeRunDeltas.length > 0) enqueueSheetRangeRunWrite(rangeRunDeltas);
      if (deltas.length > 0) enqueueWrite(deltas);

      if (needsSharedStateRevert) {
        applyingRemote = true;
        try {
          if (
            deniedInverseSheetViews.length > 0 &&
            typeof documentController.applyExternalSheetViewDeltas === "function"
          ) {
            documentController.applyExternalSheetViewDeltas(deniedInverseSheetViews, { source: "collab" });
          }
          if (deniedInverseFormats.length > 0 && typeof documentController.applyExternalFormatDeltas === "function") {
            documentController.applyExternalFormatDeltas(deniedInverseFormats, { source: "collab" });
          }
          if (
            deniedInverseRangeRuns.length > 0 &&
            typeof documentController.applyExternalRangeRunDeltas === "function"
          ) {
            documentController.applyExternalRangeRunDeltas(deniedInverseRangeRuns, { source: "collab" });
          }
        } finally {
          applyingRemote = false;
        }
      }
      return;
    }

    /** @type {any[]} */
    const allowed = [];
    /** @type {any[]} */
    const allowedFormatDeltas = [];
    /** @type {any[]} */
    const allowedRangeRunDeltas = [];
    /** @type {any[]} */
    const rejected = [];
    /** @type {any[]} */
    const deniedInverse = [];
    // `deniedInverseRangeRuns` (and sometimes `deniedInverseSheetViews`) can be pre-populated when
    // `canWriteSharedState()` is false. Additional entries may be appended below when per-range
    // permissions deny a write.

    for (const delta of deltas) {
      const cellRef = { sheetId: delta.sheetId, row: delta.row, col: delta.col };
      const allowedByPermissions = editGuard ? editGuard(cellRef) : true;
      const allowedByEncryption = needsEncryptionGuard ? canWriteCellWithEncryption(cellRef) : true;

      if (allowedByPermissions && allowedByEncryption) {
        allowed.push(delta);
      } else {
        rejected.push({
          ...delta,
          rejectionKind: "cell",
          // Prefer the dedicated encryption guard when determining why an edit was rejected.
          //
          // Note: `allowedByPermissions` can be `false` for encryption failures when callers
          // provide a `canEditCell` guard that already incorporates encryption invariants
          // (e.g. `CollabSession.canEditCell`). In those cases we still want the UI to
          // surface "missing encryption key" rather than a generic permission error.
          rejectionReason: !allowedByEncryption ? "encryption" : !allowedByPermissions ? "permission" : "unknown",
        });
        if (!allowedByEncryption) {
          console.warn(
            "bindYjsToDocumentController: refused edit to encrypted cell (missing key or encryption required)",
            makeCellKey(cellRef),
          );
        }
        deniedInverse.push({
          sheetId: delta.sheetId,
          row: delta.row,
          col: delta.col,
          before: delta.after,
          after: delta.before,
        });
      }
    }

    for (const delta of formatDeltas) {
      if (canEditFormatDelta(delta)) {
        allowedFormatDeltas.push(delta);
      } else {
        rejected.push({ ...delta, rejectionKind: "format", rejectionReason: "permission" });
        const inv = {
          sheetId: delta.sheetId,
          layer: delta.layer,
          beforeStyleId: delta.afterStyleId,
          afterStyleId: delta.beforeStyleId,
        };
        if (delta.index != null) inv.index = delta.index;
        deniedInverseFormats.push(inv);
      }
    }

    for (const delta of rangeRunDeltas) {
      if (canEditRangeRunDelta(delta)) {
        allowedRangeRunDeltas.push(delta);
      } else {
        rejected.push({ ...delta, rejectionKind: "rangeRun", rejectionReason: "permission" });
        deniedInverseRangeRuns.push({
          sheetId: delta.sheetId,
          col: delta.col,
          startRow: delta.startRow,
          endRowExclusive: delta.endRowExclusive,
          beforeRuns: delta.afterRuns,
          afterRuns: delta.beforeRuns,
        });
      }
    }

    if (allowed.length > 0) enqueueWrite(allowed);
    if (allowedFormatDeltas.length > 0) enqueueSheetFormatWrite(allowedFormatDeltas);
    if (allowedRangeRunDeltas.length > 0) enqueueSheetRangeRunWrite(allowedRangeRunDeltas);

    if (rejected.length > 0 && typeof onEditRejected === "function") {
      try {
        onEditRejected(rejected);
      } catch (err) {
        console.warn("bindYjsToDocumentController: onEditRejected callback threw", err);
      }
    }

    if (
      deniedInverse.length > 0 ||
      deniedInverseSheetViews.length > 0 ||
      deniedInverseFormats.length > 0 ||
      deniedInverseRangeRuns.length > 0
    ) {
      // Keep local UI state aligned with the shared document when a user attempts to
      // edit a restricted cell.
      applyingRemote = true;
      try {
        if (deniedInverse.length > 0) {
          if (typeof documentController.applyExternalDeltas === "function") {
            documentController.applyExternalDeltas(deniedInverse, { recalc: false, source: "collab" });
          } else {
            const prevCanEdit = "canEditCell" in documentController ? documentController.canEditCell : undefined;
            if (prevCanEdit !== undefined) documentController.canEditCell = null;
            try {
              for (const delta of deniedInverse) {
                if (delta.after.formula != null) {
                  documentController.setCellFormula(delta.sheetId, { row: delta.row, col: delta.col }, delta.after.formula);
                } else {
                  documentController.setCellValue(delta.sheetId, { row: delta.row, col: delta.col }, delta.after.value);
                }
              }
            } finally {
              if (prevCanEdit !== undefined) documentController.canEditCell = prevCanEdit;
            }
          }
        }

        if (
          deniedInverseSheetViews.length > 0 &&
          typeof documentController.applyExternalSheetViewDeltas === "function"
        ) {
          documentController.applyExternalSheetViewDeltas(deniedInverseSheetViews, { source: "collab" });
        }

        if (deniedInverseFormats.length > 0 && typeof documentController.applyExternalFormatDeltas === "function") {
          documentController.applyExternalFormatDeltas(deniedInverseFormats, { source: "collab" });
        }

        if (deniedInverseRangeRuns.length > 0 && typeof documentController.applyExternalRangeRunDeltas === "function") {
          documentController.applyExternalRangeRunDeltas(deniedInverseRangeRuns, { source: "collab" });
        }
      } finally {
        applyingRemote = false;
      }
    }
  };

  const unsubscribe = documentController.on("change", handleDocumentChange);

  cells.observeDeep(handleCellsDeepChange);
  sheets.observeDeep(handleSheetsDeepChange);

  // Initial hydration (and for cases where the provider has already applied some state).
  hydrateFromYjs();
  hydrateSheetViewsFromYjs();

  return {
    /**
     * Re-scan the Yjs document and re-apply state into the DocumentController.
     *
     * This is primarily used by desktop encryption UX: when a user imports an
     * encryption key locally (no Yjs document change), we still need to re-run
     * decryption so previously-masked cells can be revealed without requiring a reload.
     */
    rehydrate() {
      hydrateFromYjs();
      hydrateSheetViewsFromYjs();
    },
    /**
     * Wait for any pending binder work to settle.
     *
     * - Yjs → DocumentController applies are serialized through `applyChain`
     *   (encryption/decryption can be async).
     * - DocumentController → Yjs writes are serialized through `writeChain`/`sheetWriteChain`
     *   (encryption can be async).
     *
     * This helper is best-effort and primarily intended for teardown flows (e.g. flushing
     * local persistence before a hard process exit).
     */
    async whenIdle() {
      const waitForChain = async (getChain) => {
        while (true) {
          const current = getChain();
          await current.catch(() => {});
          if (getChain() === current) return;
        }
      };

      await Promise.all([
        waitForChain(() => applyChain),
        waitForChain(() => writeChain),
        waitForChain(() => sheetWriteChain),
      ]);
    },
    destroy() {
      unsubscribe?.();
      cells.unobserveDeep(handleCellsDeepChange);
      sheets.unobserveDeep(handleSheetsDeepChange);
      if (didPatchDocumentCanEditCell) {
        documentController.canEditCell = prevDocumentCanEditCell;
      }
    },
  };
}

/**
 * @param {any} permissions
 * @param {string | null} defaultUserId
 * @returns {(cell: { sheetId: string, row: number, col: number }) => { canRead: boolean, canEdit: boolean } | null}
 */
function createPermissionsResolver(permissions, defaultUserId) {
  if (!permissions) return null;

  if (typeof permissions === "function") {
    return (cell) => {
      const resolved = permissions(cell);
      const canRead = resolved?.canRead !== false;
      const canEdit = canRead && resolved?.canEdit !== false;
      return { canRead, canEdit };
    };
  }

  if (typeof permissions === "object") {
    const role = permissions.role;
    const restrictions = permissions.restrictions;
    const userId = permissions.userId ?? defaultUserId ?? null;

    if (!role) {
      throw new Error("bindYjsToDocumentController permissions.role is required when using role-based permissions");
    }

    return (cell) => {
      const { canRead, canEdit } = getCellPermissions({ role, restrictions, userId, cell });
      return { canRead, canEdit: canRead && canEdit };
    };
  }

  throw new Error("bindYjsToDocumentController permissions must be a function or { role, restrictions, userId }");
}
