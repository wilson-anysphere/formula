import * as Y from "yjs";

import { makeCellKey, normalizeCellKey, parseCellKey } from "../session/src/cell-key.js";
import {
  decryptCellPlaintext,
  encryptCellPlaintext,
  isEncryptedCellPayload,
} from "../encryption/src/index.node.js";
import { getCellPermissions, maskCellValue as defaultMaskCellValue } from "../permissions/index.js";

const MASKED_CELL_VALUE = "###";

function stableStringify(value) {
  if (value === undefined) return "undefined";
  if (value == null || typeof value !== "object") return JSON.stringify(value);
  if (Array.isArray(value)) return `[${value.map(stableStringify).join(",")}]`;
  const keys = Object.keys(value).sort();
  const entries = keys.map((k) => `${JSON.stringify(k)}:${stableStringify(value[k])}`);
  return `{${entries.join(",")}}`;
}

function getYMapCell(cellData) {
  // y-websocket currently pulls in the CJS build of Yjs, which means callers using
  // ESM `import * as Y from "yjs"` can observe Yjs types that fail `instanceof`
  // checks (same version, different module instance).
  //
  // Use a lightweight duck-type check instead of `instanceof Y.Map` so the
  // binder can interop with y-websocket providers in pnpm workspaces.
  if (!cellData || typeof cellData !== "object") return null;
  // eslint-disable-next-line no-prototype-builtins
  if (cellData.constructor?.name !== "YMap") return null;
  if (typeof cellData.get !== "function") return null;
  if (typeof cellData.set !== "function") return null;
  if (typeof cellData.delete !== "function") return null;
  return cellData;
}

function normalizeFormula(value) {
  if (value == null) return null;
  const trimmed = String(value).trim();
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
 * }} ParsedYjsCell
 */

/**
 * @param {Y.Map<any>} cell
 * @param {{ sheetId: string, row: number, col: number }} cellRef
 * @param {{
 *   keyForCell: (cell: { sheetId: string, row: number, col: number }) => { keyId: string, keyBytes: Uint8Array } | null,
 * } | null} encryption
 * @param {string} docIdForEncryption
 * @param {(value: unknown, cell?: { sheetId: string, row: number, col: number }) => unknown} maskFn
 * @returns {Promise<ParsedYjsCell>}
 */
async function readCellFromYjs(cell, cellRef, encryption, docIdForEncryption, maskFn = defaultMaskCellValue) {
  let value;
  let formula;

  // If `enc` is present, treat the cell as encrypted even if the payload is malformed.
  // This avoids accidentally falling back to plaintext fields.
  const encRaw = typeof cell.has === "function" ? (cell.has("enc") ? cell.get("enc") : undefined) : cell.get("enc");
  if (encRaw !== undefined) {
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
  if (typeof cell.has === "function" ? cell.has("format") : cell.get("format") !== undefined) {
    format = cell.get("format") ?? null;
    formatKey = stableStringify(format);
  }
  if (formula) {
    return { value: null, formula, format, formatKey };
  }
  return { value, formula: null, format, formatKey };
}

function sameNormalizedCell(a, b) {
  if (!a && !b) return true;
  if (!a || !b) return false;
  return a.formula === b.formula && Object.is(a.value, b.value) && a.formatKey === b.formatKey;
}

/**
 * Bind a Yjs spreadsheet document (the `cells` root type) to a desktop `DocumentController`.
 *
 * This binder syncs cell `value`, `formula`, and `format` (cell styles).
 * Other workbook metadata is expected to be handled by future bindings.
 *
 * @param {{
 *   ydoc: import("yjs").Doc,
 *   documentController: import("../../../apps/desktop/src/document/documentController.js").DocumentController,
 *   undoService?: { transact?: (fn: () => void) => void, origin?: any, localOrigins?: Set<any> } | null,
 *   defaultSheetId?: string,
 *   userId?: string | null,
 *   encryption?: {
 *     keyForCell: (cell: { sheetId: string, row: number, col: number }) => { keyId: string, keyBytes: Uint8Array } | null,
 *     shouldEncryptCell?: (cell: { sheetId: string, row: number, col: number }) => boolean,
 *   } | null,
 *   canReadCell?: (cell: { sheetId: string, row: number, col: number }) => boolean,
 *   canEditCell?: (cell: { sheetId: string, row: number, col: number }) => boolean,
 *   permissions?:
 *     | ((cell: { sheetId: string, row: number, col: number }) => { canRead: boolean, canEdit: boolean })
 *     | { role: string, restrictions?: any[], userId?: string | null },
 *   maskCellValue?: (value: unknown, cell?: { sheetId: string, row: number, col: number }) => unknown,
 *   onEditRejected?: (deltas: any[]) => void,
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
    maskCellValue = defaultMaskCellValue,
    onEditRejected = null,
  } = options ?? {};

  if (!ydoc) throw new Error("bindYjsToDocumentController requires { ydoc }");
  if (!documentController) throw new Error("bindYjsToDocumentController requires { documentController }");

  const cells = ydoc.getMap("cells");
  const docIdForEncryption = ydoc.guid ?? "unknown";

  // Stable origin token for local DocumentController -> Yjs transactions when
  // we don't have a dedicated undo service wrapper.
  const binderOrigin = undoService?.origin ?? { type: "document-controller:binder" };

  /**
   * Yjs "origin" values that correspond to local DocumentController-driven writes.
   * We ignore these in the Yjs -> DocumentController observer to avoid echo work.
   *
   * Important: we intentionally do NOT ignore Y.UndoManager (used for undo/redo)
   * origins so local undo/redo still updates the DocumentController.
   * @type {Set<any>}
   */
  const localOrigins = new Set([binderOrigin]);

  if (undoService?.origin) localOrigins.add(undoService.origin);

  // Some undo services expose a `localOrigins` set which includes both the
  // origin token and the UndoManager instance. We only want to treat the origin
  // token as "local" for echo suppression.
  const maybeLocalOrigins = undoService?.localOrigins;
  if (maybeLocalOrigins && typeof maybeLocalOrigins[Symbol.iterator] === "function") {
    for (const origin of maybeLocalOrigins) {
      if (isUndoManager(origin)) continue;
      localOrigins.add(origin);
    }
  }

  /** @type {Map<string, NormalizedCell>} */
  let cache = new Map();

  let applyingRemote = false;
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
   * @param {any[]} deltas
   */
  function enqueueWrite(deltas) {
    const snapshot = Array.from(deltas ?? []);
    writeChain = writeChain.then(() => applyDocumentDeltas(snapshot)).catch((err) => {
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
      if (typeof cell.has === "function" ? cell.has("format") : cell.get("format") !== undefined) {
        fallbackFormat = cell.get("format") ?? null;
        fallbackFormatKey = stableStringify(fallbackFormat);
        break;
      }
    }

    if (encryptedCandidates.length > 0) {
      const chosen =
        encryptedCandidates.find((entry) => entry.rawKey === canonicalKey) ?? encryptedCandidates[0];
      const parsed = await readCellFromYjs(chosen.cell, cellRef, encryption, docIdForEncryption, maskFn);
      if (parsed.formatKey === undefined && fallbackFormatKey !== undefined) {
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

      let styleId = before.styleId;
      if (curr?.formatKey !== undefined) {
        const format = curr.format ?? null;
        styleId = format == null ? 0 : documentController.styleTable.intern(format);
      } else if (prev?.formatKey !== undefined) {
        // `format` key removed. Treat as explicit clear even though the key is now absent.
        styleId = 0;
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
        documentController.applyExternalDeltas(deltas);
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
   * @param {any[]} events
   * @param {any} transaction
   */
  const handleCellsDeepChange = (events, transaction) => {
    if (!events || events.length === 0) return;
    const origin = transaction?.origin ?? null;
    if (localOrigins.has(origin)) return;

    /** @type {Set<string>} */
    const changed = new Set();

    for (const event of events) {
      // Nested map updates: event.path[0] is the root cells map key.
      const path = event?.path;
      if (Array.isArray(path) && path.length > 0) {
        const rawKey = path[0];
        if (typeof rawKey !== "string") continue;

        const changes = event?.changes?.keys;
        if (changes && !(changes.has("value") || changes.has("formula") || changes.has("format") || changes.has("enc"))) {
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
        encryptedPayload = await encryptCellPlaintext({
          plaintext: { value: formula != null ? null : value, formula: formula ?? null },
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
        styleId,
        format,
        formatKey,
        encryptedPayload,
      });
    }

    const apply = () => {
      for (const item of prepared) {
        const { canonicalKey, targets, value, formula, styleId, format, encryptedPayload, formatKey } = item;

        if (value == null && formula == null && styleId === 0 && !encryptedPayload) {
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
            cell = new Y.Map();
            cells.set(rawKey, cell);
          }

          if (encryptedPayload) {
            cell.set("enc", encryptedPayload);
            cell.delete("value");
            cell.delete("formula");
          } else {
            cell.delete("enc");
            if (formula != null) {
              cell.set("formula", formula);
              cell.set("value", null);
            } else {
              cell.delete("formula");
              cell.set("value", value);
            }
          }

          if (styleId === 0) {
            cell.delete("format");
          } else {
            cell.set("format", format);
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
      undoService.transact(apply);
    } else {
      ydoc.transact(apply, binderOrigin);
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
   * @param {any} payload
   */
  const handleDocumentChange = (payload) => {
    if (applyingRemote) return;
    const deltas = Array.isArray(payload?.deltas) ? payload.deltas : [];
    if (deltas.length === 0) return;

    const needsEncryptionGuard = Boolean(encryption || hasEncryptedCells);
    if (!editGuard && !needsEncryptionGuard) {
      enqueueWrite(deltas);
      return;
    }

    /** @type {any[]} */
    const allowed = [];
    /** @type {any[]} */
    const rejected = [];
    /** @type {any[]} */
    const deniedInverse = [];

    for (const delta of deltas) {
      const cellRef = { sheetId: delta.sheetId, row: delta.row, col: delta.col };
      const allowedByPermissions = editGuard ? editGuard(cellRef) : true;
      const allowedByEncryption = needsEncryptionGuard ? canWriteCellWithEncryption(cellRef) : true;

      if (allowedByPermissions && allowedByEncryption) {
        allowed.push(delta);
      } else {
        rejected.push(delta);
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

    if (allowed.length > 0) enqueueWrite(allowed);

    if (rejected.length > 0 && typeof onEditRejected === "function") {
      try {
        onEditRejected(rejected);
      } catch (err) {
        console.warn("bindYjsToDocumentController: onEditRejected callback threw", err);
      }
    }

    if (deniedInverse.length > 0) {
      // Keep local UI state aligned with the shared document when a user attempts to
      // edit a restricted cell.
      applyingRemote = true;
      try {
        if (typeof documentController.applyExternalDeltas === "function") {
          documentController.applyExternalDeltas(deniedInverse, { recalc: false });
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
      } finally {
        applyingRemote = false;
      }
    }
  };

  const unsubscribe = documentController.on("change", handleDocumentChange);

  cells.observeDeep(handleCellsDeepChange);

  // Initial hydration (and for cases where the provider has already applied some state).
  hydrateFromYjs();

  return {
    destroy() {
      unsubscribe?.();
      cells.unobserveDeep(handleCellsDeepChange);
      if (didPatchDocumentCanEditCell) {
        documentController.canEditCell = prevDocumentCanEditCell;
      }
    },
  };
}

function isUndoManager(value) {
  if (!value || typeof value !== "object") return false;
  const maybe = value;
  if (maybe.constructor?.name === "UndoManager") return true;
  return typeof maybe.undo === "function" && typeof maybe.redo === "function" && maybe.trackedOrigins instanceof Set;
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
