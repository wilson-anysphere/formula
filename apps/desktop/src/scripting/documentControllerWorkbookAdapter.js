import { formatCellAddress, formatRangeAddress, parseRangeAddress } from "../../../../packages/scripting/src/a1.js";
import { TypedEventEmitter } from "../../../../packages/scripting/src/events.js";
import { applyStylePatch } from "../formatting/styleTable.js";
import { getStyleNumberFormat } from "../formatting/styleFieldAccess.js";
import { parseImageCellValue } from "../shared/imageCellValue.js";

// Excel limits used by the scripting A1 address helpers and macro recorder.
// - Rows: 1..1048576 (0-based: 0..1048575)
// - Cols: A..XFD (0-based: 0..16383)
const EXCEL_MAX_ROW = 1048575;
const EXCEL_MAX_COL = 16383;

// Scripting APIs (and macro recorder) frequently represent ranges as full 2D JS arrays.
// With Excel-scale sheets this can explode in memory usage if callers request a full
// column/row/sheet. Keep reads/writes bounded to avoid renderer OOMs.
const DEFAULT_SCRIPT_RANGE_CELL_LIMIT = 200_000;

function assertScriptRangeWithinLimit(rows, cols, address, action) {
  const cellCount = rows * cols;
  if (cellCount > DEFAULT_SCRIPT_RANGE_CELL_LIMIT) {
    throw new Error(
      `${action} skipped for range ${address} (rows=${rows}, cols=${cols}, area=${cellCount}). ` +
        `Limit is ${DEFAULT_SCRIPT_RANGE_CELL_LIMIT} cells.`,
    );
  }
}
/**
 * @typedef {import("../sheet/sheetNameResolver.js").SheetNameResolver} SheetNameResolver
 */

function valueEquals(a, b) {
  if (a === b) return true;
  if (a == null || b == null) return false;
  if (typeof a === "object" && typeof b === "object") {
    try {
      return JSON.stringify(a) === JSON.stringify(b);
    } catch {
      return false;
    }
  }
  return false;
}

function isPlainObject(value) {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function stableStringify(value) {
  if (value === undefined) return "undefined";
  if (value == null || typeof value !== "object") return JSON.stringify(value);
  if (Array.isArray(value)) return `[${value.map(stableStringify).join(",")}]`;
  const keys = Object.keys(value).sort();
  const entries = keys.map((k) => `${JSON.stringify(k)}:${stableStringify(value[k])}`);
  return `{${entries.join(",")}}`;
}

function hasOwn(obj, key) {
  return Object.prototype.hasOwnProperty.call(obj, key);
}

function normalizeFormulaText(formula) {
  if (typeof formula !== "string") return null;
  const trimmed = formula.trim();
  const strippedLeading = trimmed.startsWith("=") ? trimmed.slice(1) : trimmed;
  const stripped = strippedLeading.trim();
  if (stripped === "") return null;
  return `=${stripped}`;
}

function scriptCellValueFromStateValue(value) {
  if (value == null) return null;

  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    return value;
  }

  // DocumentController stores rich text as `{ text, runs }`. Scripting does not currently
  // expose rich text formatting, so surface the plain text content.
  if (typeof value === "object" && typeof value.text === "string") {
    return value.text;
  }

  const image = parseImageCellValue(value);
  if (image) return image.altText ?? "[Image]";

  // Best-effort fallback for other object-like values (e.g. Date objects from Parquet import).
  if (value instanceof Date) {
    try {
      return value.toISOString();
    } catch {
      return String(value);
    }
  }

  try {
    return JSON.stringify(value);
  } catch {
    return String(value);
  }
}

function cellInputFromState(state) {
  if (state.formula != null) return normalizeFormulaText(state.formula);
  const value = scriptCellValueFromStateValue(state.value ?? null);
  // This scripting surface treats strings that start with "=" as formulas on write. To allow
  // round-tripping literal strings that start with "=", we re-add the leading apostrophe when
  // reading from the DocumentController (which strips it during input normalization).
  if (typeof value === "string" && value.trimStart().startsWith("=")) {
    return `'${value}`;
  }
  return value;
}

function isFormulaString(input) {
  if (typeof input !== "string") return false;
  const trimmed = input.trimStart();
  return trimmed.startsWith("=");
}

function denseRectForDeltas(deltas) {
  if (!deltas || deltas.length === 0) return null;

  let minRow = Number.POSITIVE_INFINITY;
  let maxRow = Number.NEGATIVE_INFINITY;
  let minCol = Number.POSITIVE_INFINITY;
  let maxCol = Number.NEGATIVE_INFINITY;
  /** @type {Map<string, any>} */
  const byCell = new Map();

  for (const delta of deltas) {
    minRow = Math.min(minRow, delta.row);
    maxRow = Math.max(maxRow, delta.row);
    minCol = Math.min(minCol, delta.col);
    maxCol = Math.max(maxCol, delta.col);
    const key = `${delta.row},${delta.col}`;
    if (byCell.has(key)) return null;
    byCell.set(key, delta);
  }

  const rows = maxRow - minRow + 1;
  const cols = maxCol - minCol + 1;
  if (rows * cols !== deltas.length) return null;

  return { minRow, maxRow, minCol, maxCol, rows, cols, byCell };
}

/**
 * Diff two range-run formatting lists and return the row segments where the effective style id
 * differs between the before/after states.
 *
 * Runs are:
 * - sorted by `startRow`
 * - non-overlapping
 * - half-open `[startRow, endRowExclusive)`
 * - sparse (styleId=0 is represented by absence of a run)
 *
 * The returned segments are within `[startRow, endRowExclusive)` and are also half-open.
 *
 * @param {any[]} beforeRuns
 * @param {any[]} afterRuns
 * @param {number} startRow
 * @param {number} endRowExclusive
 * @returns {Array<{ startRow: number, endRowExclusive: number, beforeStyleId: number, afterStyleId: number }>}
 */
function diffRangeRunSegments(beforeRuns, afterRuns, startRow, endRowExclusive) {
  /** @type {Array<{ startRow: number, endRowExclusive: number, beforeStyleId: number, afterStyleId: number }>} */
  const out = [];

  if (!Number.isInteger(startRow) || !Number.isInteger(endRowExclusive)) return out;
  if (endRowExclusive <= startRow) return out;

  // Clamp to Excel bounds (DocumentController uses an exclusive max row of EXCEL_MAX_ROW + 1).
  const maxRowExclusive = EXCEL_MAX_ROW + 1;
  let pos = Math.max(0, Math.min(maxRowExclusive, startRow));
  const end = Math.max(pos, Math.min(maxRowExclusive, endRowExclusive));
  if (end <= pos) return out;

  const before = Array.isArray(beforeRuns) ? beforeRuns : [];
  const after = Array.isArray(afterRuns) ? afterRuns : [];

  let i = 0;
  let j = 0;
  while (i < before.length && Number(before[i]?.endRowExclusive ?? 0) <= pos) i += 1;
  while (j < after.length && Number(after[j]?.endRowExclusive ?? 0) <= pos) j += 1;

  while (pos < end) {
    let beforeStyleId = 0;
    let beforeEnd = end;
    if (i < before.length) {
      const run = before[i];
      const runStart = Number(run?.startRow);
      const runEnd = Number(run?.endRowExclusive);
      if (Number.isFinite(runStart) && Number.isFinite(runEnd)) {
        if (runStart <= pos && pos < runEnd) {
          beforeStyleId = Number(run?.styleId ?? 0);
          beforeEnd = Math.min(runEnd, end);
        } else if (runStart > pos) {
          beforeEnd = Math.min(runStart, end);
        }
      }
    }

    let afterStyleId = 0;
    let afterEnd = end;
    if (j < after.length) {
      const run = after[j];
      const runStart = Number(run?.startRow);
      const runEnd = Number(run?.endRowExclusive);
      if (Number.isFinite(runStart) && Number.isFinite(runEnd)) {
        if (runStart <= pos && pos < runEnd) {
          afterStyleId = Number(run?.styleId ?? 0);
          afterEnd = Math.min(runEnd, end);
        } else if (runStart > pos) {
          afterEnd = Math.min(runStart, end);
        }
      }
    }

    const nextPos = Math.min(beforeEnd, afterEnd, end);
    if (nextPos <= pos) break;

    if (beforeStyleId !== afterStyleId) {
      const last = out[out.length - 1];
      if (
        last &&
        last.endRowExclusive === pos &&
        last.beforeStyleId === beforeStyleId &&
        last.afterStyleId === afterStyleId
      ) {
        last.endRowExclusive = nextPos;
      } else {
        out.push({ startRow: pos, endRowExclusive: nextPos, beforeStyleId, afterStyleId });
      }
    }

    pos = nextPos;
    while (i < before.length && Number(before[i]?.endRowExclusive ?? 0) <= pos) i += 1;
    while (j < after.length && Number(after[j]?.endRowExclusive ?? 0) <= pos) j += 1;
  }

  return out;
}

/**
 * Convert a DocumentController style object into the (currently minimal) scripting `CellFormat`
 * shape defined in `packages/scripting/formula.d.ts`.
 *
 * @param {any} style
 */
function scriptFormatFromDocStyle(style) {
  const out = {};
  if (!isPlainObject(style)) return out;

  // DocumentController native schema: { font: { bold, italic }, numberFormat, fill: { fgColor } }
  const font = style.font;
  if (isPlainObject(font)) {
    if (hasOwn(font, "bold") && typeof font.bold === "boolean") out.bold = font.bold;
    if (hasOwn(font, "italic") && typeof font.italic === "boolean") out.italic = font.italic;
  }

  // Back-compat: earlier scripting adapters stored these at the top-level.
  if (!hasOwn(out, "bold") && hasOwn(style, "bold") && typeof style.bold === "boolean") out.bold = style.bold;
  if (!hasOwn(out, "italic") && hasOwn(style, "italic") && typeof style.italic === "boolean") out.italic = style.italic;

  const numberFormat = getStyleNumberFormat(style);
  if (numberFormat != null) out.numberFormat = numberFormat;

  const fill = style.fill;
  if (isPlainObject(fill)) {
    const fgColor = fill.fgColor ?? fill.background;
    if (typeof fgColor === "string") out.backgroundColor = fgColor;
  }

  if (!hasOwn(out, "backgroundColor") && hasOwn(style, "backgroundColor") && typeof style.backgroundColor === "string") {
    out.backgroundColor = style.backgroundColor;
  }

  return out;
}

/**
 * Map a scripting `CellFormat` patch (flat keys like `{ bold: true }`) into the DocumentController
 * style patch schema.
 *
 * This function also preserves any DocumentController-style keys already present on the patch
 * (e.g. `{ border: ... }`), so callers can use advanced formatting without losing expressiveness.
 *
 * @param {any} format
 * @returns {any}
 */
function docStylePatchFromScriptFormat(format) {
  if (format == null) return null;
  if (!isPlainObject(format)) return format;

  /** @type {Record<string, any>} */
  const out = {};

  for (const [key, value] of Object.entries(format)) {
    if (key === "bold" || key === "italic") {
      // Handled below.
      continue;
    }
    if (key === "backgroundColor") {
      // Handled below.
      continue;
    }
    if (key === "numberFormat" || key === "number_format") {
      // Normalized below.
      continue;
    }
    out[key] = value;
  }

  if (hasOwn(format, "bold") || hasOwn(format, "italic")) {
    const fontPatch = isPlainObject(out.font) ? { ...out.font } : {};
    if (hasOwn(format, "bold")) fontPatch.bold = format.bold;
    if (hasOwn(format, "italic")) fontPatch.italic = format.italic;
    out.font = fontPatch;
  }

  if (hasOwn(format, "backgroundColor")) {
    const color = format.backgroundColor;
    if (color == null) {
      out.fill = null;
    } else {
      const fillPatch = isPlainObject(out.fill) ? { ...out.fill } : {};
      fillPatch.pattern = "solid";
      fillPatch.fgColor = color;
      out.fill = fillPatch;
    }
  }

  // Number formats: treat empty/"General" as clearing via `numberFormat: null` (Excel semantics).
  // Prefer the camelCase key when both are present.
  if (hasOwn(format, "numberFormat") || hasOwn(format, "number_format")) {
    const raw = hasOwn(format, "numberFormat") ? format.numberFormat : format.number_format;
    if (raw == null) {
      out.numberFormat = null;
    } else if (typeof raw === "string") {
      const trimmed = raw.trim();
      out.numberFormat = !trimmed || trimmed.toLowerCase() === "general" ? null : raw;
    } else if (raw !== undefined) {
      // Be permissive: scripting callers may provide advanced patches.
      out.numberFormat = raw;
    }
    // Keep the canonical key when writing into DocumentController's style table.
    delete out.number_format;
  }

  return out;
}

function diffStylePatch(beforeStyle, afterStyle) {
  const before = isPlainObject(beforeStyle) ? beforeStyle : {};
  const after = isPlainObject(afterStyle) ? afterStyle : {};

  const beforeKeys = Object.keys(before);
  const afterKeys = Object.keys(after);
  if (afterKeys.length === 0 && beforeKeys.length > 0) {
    // Clearing formatting is represented by `null` in the DocumentController API.
    return null;
  }

  /** @type {Record<string, any>} */
  const patch = {};
  for (const key of afterKeys) {
    const a = after[key];
    const b = before[key];
    if (stableStringify(a) === stableStringify(b)) continue;

    if (isPlainObject(a) && isPlainObject(b)) {
      const nested = diffStylePatch(b, a);
      // A nested "clear" isn't representable with the current format patch semantics, so
      // fall back to emitting the full nested object.
      if (nested != null && Object.keys(nested).length > 0) {
        patch[key] = nested;
      } else if (nested == null) {
        patch[key] = a;
      } else {
        patch[key] = a;
      }
      continue;
    }

    patch[key] = a;
  }

  return patch;
}

/**
 * Convert a DocumentController style patch into a scripting format patch where possible.
 *
 * - Promotes `font.bold`/`font.italic` into top-level `bold`/`italic`.
 * - Promotes `fill.fgColor` into top-level `backgroundColor`.
 *
 * Any non-representable keys are kept (so macro recording doesn't lose information).
 *
 * @param {any} patch
 */
function scriptFormatPatchFromDocStylePatch(patch) {
  if (patch == null) return null;
  if (!isPlainObject(patch)) return patch;

  /** @type {Record<string, any>} */
  const out = {};

  // Copy non-formatting keys directly.
  for (const [key, value] of Object.entries(patch)) {
    if (key === "font" || key === "fill") continue;
    out[key] = value;
  }

  if (isPlainObject(patch.font)) {
    const font = { ...patch.font };
    if (hasOwn(font, "bold")) {
      out.bold = font.bold;
      delete font.bold;
    }
    if (hasOwn(font, "italic")) {
      out.italic = font.italic;
      delete font.italic;
    }
    if (Object.keys(font).length > 0) {
      out.font = font;
    }
  }

  if (patch.fill === null) {
    out.backgroundColor = null;
  } else if (isPlainObject(patch.fill)) {
    const fill = { ...patch.fill };
    const color = fill.fgColor ?? fill.background;
    if (color !== undefined) {
      out.backgroundColor = color;
      delete fill.fgColor;
      delete fill.background;
      delete fill.pattern;
    }
    if (Object.keys(fill).length > 0) {
      out.fill = fill;
    }
  }

  return out;
}

/**
 * Resolve a DocumentController style reference into a concrete style object.
 *
 * Layered formatting APIs may expose style ids directly (`number`) or via objects
 * containing a `styleId` property, while older paths only expose per-cell style
 * ids on the cell state.
 *
 * @param {any} styleTable
 * @param {any} ref
 * @returns {any}
 */
function resolveDocStyle(styleTable, ref) {
  if (!styleTable) return isPlainObject(ref) ? ref : {};
  if (ref == null) return {};
  if (typeof ref === "number") return styleTable.get(ref);
  if (isPlainObject(ref) && typeof ref.styleId === "number") return styleTable.get(ref.styleId);
  // Assume it's already a style object.
  return ref;
}

/**
 * @param {any} delta
 * @param {"before" | "after"} which
 */
function deltaStyleRef(delta, which) {
  if (!delta) return null;
  const suffix = which === "before" ? "before" : "after";
  // Common shapes: `{ beforeStyleId, afterStyleId }`, `{ before, after }`.
  if (hasOwn(delta, `${suffix}StyleId`)) return delta[`${suffix}StyleId`];
  if (hasOwn(delta, suffix)) return delta[suffix];
  // Fallbacks for alternative naming schemes.
  if (which === "before") {
    if (hasOwn(delta, "oldStyleId")) return delta.oldStyleId;
    if (hasOwn(delta, "prevStyleId")) return delta.prevStyleId;
  } else {
    if (hasOwn(delta, "newStyleId")) return delta.newStyleId;
    if (hasOwn(delta, "nextStyleId")) return delta.nextStyleId;
  }
  return null;
}

/**
 * Normalize an arbitrary row/col style override table into a sparse index map.
 *
 * Supports the formats we currently see across the codebase:
 * - `Record<string, styleRef>`
 * - `Map<number, styleRef>`
 * - Dense arrays where `raw[idx]` is the styleRef for that index
 * - Entry arrays: `[[idx, styleRef], ...]` or `[{ index, format/style/styleId }, ...]`
 *
 * We intentionally store only non-empty entries (0 / null / {} are treated as "no override").
 *
 * @param {any} raw
 * @returns {Map<number, any>}
 */
function styleIndexMapFromRaw(raw) {
  /** @type {Map<number, any>} */
  const out = new Map();

  /**
   * @param {any} idxRaw
   * @param {any} valueRaw
   */
  const add = (idxRaw, valueRaw) => {
    const idx = Number(idxRaw);
    if (!Number.isInteger(idx) || idx < 0) return;

    if (valueRaw == null) return;
    if (typeof valueRaw === "number") {
      const id = Number(valueRaw);
      if (!Number.isInteger(id) || id <= 0) return;
      out.set(idx, id);
      return;
    }
    if (isPlainObject(valueRaw)) {
      if (Object.keys(valueRaw).length === 0) return;
      out.set(idx, valueRaw);
      return;
    }

    // Unknown styleRef shape; keep as-is so downstream patching has a chance.
    out.set(idx, valueRaw);
  };

  if (!raw) return out;

  if (raw instanceof Map) {
    for (const [idx, value] of raw.entries()) add(idx, value);
    return out;
  }

  if (Array.isArray(raw)) {
    const looksLikeEntryList = raw.some(
      (entry) => Array.isArray(entry) || (isPlainObject(entry) && (hasOwn(entry, "index") || hasOwn(entry, "row") || hasOwn(entry, "col"))),
    );

    if (looksLikeEntryList) {
      for (const entry of raw) {
        if (Array.isArray(entry)) {
          add(entry[0], entry[1]);
          continue;
        }
        if (isPlainObject(entry)) {
          add(entry.index ?? entry.row ?? entry.col, entry.format ?? entry.style ?? entry.styleId);
        }
      }
    } else {
      // Dense array: values are indexed by position.
      for (let idx = 0; idx < raw.length; idx++) add(idx, raw[idx]);
    }

    return out;
  }

  if (isPlainObject(raw)) {
    for (const [idx, value] of Object.entries(raw)) add(idx, value);
  }

  return out;
}

/**
 * Exposes a `DocumentController` instance through the `@formula/scripting` Workbook/Sheet/Range
 * surface area.
 *
 * This adapter focuses on:
 * - The synchronous RPC surface used by `ScriptRuntime` (get/set values, formats, selection)
 * - The events consumed by `MacroRecorder` (`cellChanged`, `formatChanged`, `selectionChanged`)
 */
export class DocumentControllerWorkbookAdapter {
  /**
   * @param {import("../document/documentController.js").DocumentController} documentController
   * @param {{
   *   activeSheetName?: string,
   *   sheetNameResolver?: SheetNameResolver | null,
   *   getActiveSheetName?: () => string,
   *   getSelection?: () => { sheetName: string, address: string },
   *   setSelection?: (sheetName: string, address: string) => void,
   *   onDidMutate?: () => void,
   * }} [options]
   */
  constructor(documentController, options = {}) {
    this.documentController = documentController;
    this.events = new TypedEventEmitter();
    /** @type {Map<string, DocumentControllerSheetAdapter>} */
    this.sheets = new Map();
    this.activeSheetName = options.activeSheetName ?? "Sheet1";
    /** @type {{ sheetName: string, address: string } | null} */
    this.selection = null;

    this.sheetNameResolver = options.sheetNameResolver ?? null;

    this.getActiveSheetNameImpl = typeof options.getActiveSheetName === "function" ? options.getActiveSheetName : null;
    this.getSelectionImpl = typeof options.getSelection === "function" ? options.getSelection : null;
    this.setSelectionImpl = typeof options.setSelection === "function" ? options.setSelection : null;
    this.onDidMutate = typeof options.onDidMutate === "function" ? options.onDidMutate : null;

    this.unsubscribes = [this.documentController.on("change", (payload) => this.#handleDocumentChange(payload))];
  }

  /** @type {import("../document/documentController.js").DocumentController} */
  documentController;

  /** @type {TypedEventEmitter} */
  events;

  /** @type {Map<string, DocumentControllerSheetAdapter>} */
  sheets;

  /** @type {string} */
  activeSheetName;

  /** @type {{ sheetName: string, address: string } | null} */
  selection;

  /** @type {SheetNameResolver | null} */
  sheetNameResolver;

  /** @type {(() => string) | null} */
  getActiveSheetNameImpl;

  /** @type {(() => { sheetName: string, address: string }) | null} */
  getSelectionImpl;

  /** @type {((sheetName: string, address: string) => void) | null} */
  setSelectionImpl;

  /** @type {(() => void) | null} */
  onDidMutate;

  /** @type {Array<() => void>} */
  unsubscribes;

  dispose() {
    for (const unsub of this.unsubscribes) unsub();
    this.unsubscribes = [];
  }

  _notifyMutate() {
    this.onDidMutate?.();
  }

  getSheet(name) {
    const sheetName = String(name);
    const sheetId = this.#resolveSheetId(sheetName);
    if (!sheetId) {
      throw new Error(`Unknown sheet: ${sheetName}`);
    }

    let sheet = this.sheets.get(sheetId);
    if (!sheet) {
      sheet = new DocumentControllerSheetAdapter(this, sheetId);
      this.sheets.set(sheetId, sheet);
    }
    return sheet;
  }

  /**
   * Return all known sheets.
   *
   * DocumentController materializes sheets lazily (on first access). For scripting we still want
   * a stable surface that always includes the active sheet, so we union the controller's
   * materialized sheet ids with the adapter's active sheet name.
   */
  getSheets() {
    const raw = this.documentController.getSheetIds?.();
    const ids = Array.isArray(raw) ? raw : [];
    const active = this.getActiveSheet().name;
    const out = ids.slice();
    if (!out.includes(active)) out.push(active);
    return out.map((name) => this.getSheet(name));
  }

  getActiveSheet() {
    const sheetName = this.getActiveSheetNameImpl ? this.getActiveSheetNameImpl() : this.activeSheetName;
    return this.getSheet(sheetName);
  }

  getActiveSheetName() {
    return this.getActiveSheet().name;
  }

  setActiveSheet(name) {
    this.activeSheetName = String(name);
  }

  getSelection() {
    if (this.getSelectionImpl) {
      const sel = this.getSelectionImpl();
      const rawSheetName = String(sel.sheetName);
      const sheetId = this.#resolveSheetId(rawSheetName);
      const displayName = sheetId ? this.#sheetDisplayName(sheetId) : rawSheetName;
      return { sheetName: displayName, address: String(sel.address) };
    }

    if (!this.selection) {
      this.selection = { sheetName: this.getActiveSheet().name, address: "A1" };
    }
    return this.selection;
  }

  setSelection(sheetName, address) {
    const normalizedSheetName = String(sheetName);
    const normalizedAddress = String(address);
    // Validate address early to keep selection consistent.
    parseRangeAddress(normalizedAddress);

    const sheetId = this.#resolveSheetId(normalizedSheetName);
    if (!sheetId) {
      throw new Error(`Unknown sheet: ${normalizedSheetName}`);
    }

    const displayName = this.#sheetDisplayName(sheetId);
    this.selection = { sheetName: displayName, address: normalizedAddress };
    this.events.emit("selectionChanged", this.selection);
    // `setSelectionImpl` expects the DocumentController sheet id (used by the UI layer).
    this.setSelectionImpl?.(sheetId, normalizedAddress);
  }

  /**
   * @param {{ deltas: Array<any> }} payload
   */
  #handleDocumentChange(payload) {
    const deltas = Array.isArray(payload?.deltas) ? payload.deltas : [];
    const styleTable = this.documentController.styleTable;
    /** @type {Map<string, { value: any[], format: any[] }>} */
    const bySheet = new Map();

    for (const delta of deltas) {
      if (!delta) continue;
      const sheetName = this.#sheetDisplayName(delta.sheetId);
      let bucket = bySheet.get(sheetName);
      if (!bucket) {
        bucket = { value: [], format: [] };
        bySheet.set(sheetName, bucket);
      }

      const valueChanged =
        !valueEquals(delta.before?.value ?? null, delta.after?.value ?? null) ||
        (delta.before?.formula ?? null) !== (delta.after?.formula ?? null);
      const formatChanged = (delta.before?.styleId ?? 0) !== (delta.after?.styleId ?? 0);

      if (valueChanged) bucket.value.push(delta);
      if (formatChanged) bucket.format.push(delta);
    }

    for (const [sheetName, bucket] of bySheet.entries()) {
      if (bucket.value.length > 0) {
        const rect = denseRectForDeltas(bucket.value);
        if (rect && rect.rows * rect.cols > 1) {
          const values = [];
          for (let r = 0; r < rect.rows; r++) {
            const row = [];
            for (let c = 0; c < rect.cols; c++) {
              const delta = rect.byCell.get(`${rect.minRow + r},${rect.minCol + c}`);
              if (!delta) {
                // Should be unreachable given the dense check, but fall back to per-cell events.
                row.length = 0;
                break;
              }
              row.push(cellInputFromState(delta.after ?? {}));
            }
            if (row.length === 0) break;
            values.push(row);
          }

          if (values.length === rect.rows) {
            this.events.emit("cellChanged", {
              sheetName,
              address: formatRangeAddress({
                startRow: rect.minRow,
                startCol: rect.minCol,
                endRow: rect.maxRow,
                endCol: rect.maxCol,
              }),
              values,
            });
          } else {
            for (const delta of bucket.value) {
              const address = formatCellAddress({ row: delta.row, col: delta.col });
              this.events.emit("cellChanged", {
                sheetName,
                address,
                values: [[cellInputFromState(delta.after ?? {})]],
              });
            }
          }
        } else {
          for (const delta of bucket.value) {
            const address = formatCellAddress({ row: delta.row, col: delta.col });
            this.events.emit("cellChanged", {
              sheetName,
              address,
              values: [[cellInputFromState(delta.after ?? {})]],
            });
          }
        }
      }

      if (bucket.format.length > 0) {
        /** @type {Map<string, { patch: any, deltas: any[] }>} */
        const groups = new Map();
        for (const delta of bucket.format) {
          const beforeStyle = styleTable.get(delta.before?.styleId ?? 0);
          const afterStyle = styleTable.get(delta.after?.styleId ?? 0);
          const docPatch = diffStylePatch(beforeStyle, afterStyle);
          const patch = scriptFormatPatchFromDocStylePatch(docPatch);
          if (patch !== null && isPlainObject(patch) && Object.keys(patch).length === 0) continue;

          const key = patch === null ? "null" : stableStringify(patch);
          const entry = groups.get(key) ?? { patch, deltas: [] };
          entry.deltas.push(delta);
          groups.set(key, entry);
        }

        for (const entry of groups.values()) {
          const rect = denseRectForDeltas(entry.deltas);
          if (rect && rect.rows * rect.cols > 1) {
            this.events.emit("formatChanged", {
              sheetName,
              address: formatRangeAddress({
                startRow: rect.minRow,
                startCol: rect.minCol,
                endRow: rect.maxRow,
                endCol: rect.maxCol,
              }),
              format: entry.patch,
            });
            continue;
          }

          for (const delta of entry.deltas) {
            const address = formatCellAddress({ row: delta.row, col: delta.col });
            this.events.emit("formatChanged", {
              sheetName,
              address,
              format: entry.patch,
            });
          }
        }
      }
    }

    // Range-run formatting changes (sheet.formatRunsByCol). Large rectangular format changes are
    // stored here instead of emitting per-cell deltas, so synthesize formatChanged events for
    // the macro recorder.
    const rangeRunDeltas = Array.isArray(payload?.rangeRunDeltas) ? payload.rangeRunDeltas : [];
    if (rangeRunDeltas.length > 0) {
      /** @type {Map<string, Map<string, { patch: any, startRow: number, endRowExclusive: number, cols: Set<number> }>>} */
      const rangeRunGroupsBySheet = new Map();
      /** @type {Map<string, { patch: any, key: string, skip: boolean }>} */
      const patchCache = new Map();

      for (const delta of rangeRunDeltas) {
        const sheetName = typeof delta?.sheetId === "string" ? delta.sheetId : typeof delta?.sheetName === "string" ? delta.sheetName : null;
        if (!sheetName) continue;

        const col = Number(delta.col ?? delta.column ?? delta.colIndex ?? delta.columnIndex);
        if (!Number.isInteger(col) || col < 0 || col > EXCEL_MAX_COL) continue;

        const startRow = Number(delta.startRow);
        const endRowExclusive = Number(delta.endRowExclusive);
        if (!Number.isInteger(startRow) || startRow < 0) continue;
        if (!Number.isInteger(endRowExclusive) || endRowExclusive <= startRow) continue;

        const segments = diffRangeRunSegments(delta.beforeRuns, delta.afterRuns, startRow, endRowExclusive);
        for (const segment of segments) {
          const beforeStyleId = segment.beforeStyleId ?? 0;
          const afterStyleId = segment.afterStyleId ?? 0;
          const cacheKey = `${beforeStyleId},${afterStyleId}`;
          let cached = patchCache.get(cacheKey);
          if (!cached) {
            const beforeStyle = styleTable?.get(beforeStyleId) ?? {};
            const afterStyle = styleTable?.get(afterStyleId) ?? {};
            const docPatch = diffStylePatch(beforeStyle, afterStyle);
            const patch = scriptFormatPatchFromDocStylePatch(docPatch);
            if (patch !== null && isPlainObject(patch) && Object.keys(patch).length === 0) {
              cached = { patch: null, key: "", skip: true };
            } else {
              cached = { patch, key: patch === null ? "null" : stableStringify(patch), skip: false };
            }
            patchCache.set(cacheKey, cached);
          }
          if (cached.skip) continue;

          const groupKey = `${segment.startRow}:${segment.endRowExclusive}:${cached.key}`;
          let groups = rangeRunGroupsBySheet.get(sheetName);
          if (!groups) {
            groups = new Map();
            rangeRunGroupsBySheet.set(sheetName, groups);
          }
          const entry = groups.get(groupKey) ?? {
            patch: cached.patch,
            startRow: segment.startRow,
            endRowExclusive: segment.endRowExclusive,
            cols: new Set(),
          };
          entry.cols.add(col);
          groups.set(groupKey, entry);
        }
      }

      for (const [sheetName, groups] of rangeRunGroupsBySheet.entries()) {
        for (const entry of groups.values()) {
          const cols = [...entry.cols].sort((a, b) => a - b);
          let i = 0;
          while (i < cols.length) {
            let j = i + 1;
            while (j < cols.length && cols[j] === cols[j - 1] + 1) j += 1;
            const startCol = cols[i];
            const endCol = cols[j - 1];
            this.events.emit("formatChanged", {
              sheetName,
              address: formatRangeAddress({
                startRow: entry.startRow,
                startCol,
                endRow: entry.endRowExclusive - 1,
                endCol,
              }),
              format: entry.patch,
            });
            i = j;
          }
        }
      }
    }

    // Layered formatting deltas (row/col/sheet) may not produce any per-cell styleId changes,
    // so we must also listen for those explicit delta streams (when present) and synthesize
    // macro recorder format events.
    /** @type {any[]} */
    const rawColStyleDeltas = [];
    /** @type {any[]} */
    const rawRowStyleDeltas = [];
    /** @type {any[]} */
    const rawSheetStyleDeltas = [];

    const addAll = (target, value) => {
      if (Array.isArray(value)) {
        target.push(...value);
      } else if (value && typeof value === "object") {
        target.push(value);
      }
    };

    // Common payload keys (preferred).
    addAll(rawColStyleDeltas, payload?.colStyleIdDeltas);
    addAll(rawColStyleDeltas, payload?.colStyleDeltas);
    addAll(rawColStyleDeltas, payload?.columnStyleIdDeltas);
    addAll(rawColStyleDeltas, payload?.columnStyleDeltas);
    addAll(rawRowStyleDeltas, payload?.rowStyleIdDeltas);
    addAll(rawRowStyleDeltas, payload?.rowStyleDeltas);
    addAll(rawSheetStyleDeltas, payload?.sheetStyleIdDeltas);
    addAll(rawSheetStyleDeltas, payload?.sheetDefaultStyleIdDeltas);
    addAll(rawSheetStyleDeltas, payload?.sheetStyleDeltas);
    addAll(rawSheetStyleDeltas, payload?.sheetDefaultStyleDeltas);

    // Some integrations may bundle all format deltas into a `formatDeltas` object/array.
    const formatDeltas = payload?.formatDeltas;
    if (Array.isArray(formatDeltas)) {
      for (const delta of formatDeltas) {
        if (!delta) continue;

        // DocumentController layered formatting emits `FormatDelta` entries:
        //   { sheetId, layer: "sheet" | "row" | "col", index?, beforeStyleId, afterStyleId }
        // Normalize those into the row/col/sheet delta streams expected by the macro recorder.
        const layer =
          typeof delta.layer === "string"
            ? delta.layer
            : typeof delta.scope === "string"
              ? delta.scope
              : typeof delta.kind === "string"
                ? delta.kind
                : null;
        if (layer === "col") {
          const col = delta.col ?? delta.column ?? delta.colIndex ?? delta.columnIndex ?? delta.index;
          rawColStyleDeltas.push({ ...delta, col });
          continue;
        }
        if (layer === "row") {
          const row = delta.row ?? delta.rowIndex ?? delta.index;
          rawRowStyleDeltas.push({ ...delta, row });
          continue;
        }
        if (layer === "sheet") {
          rawSheetStyleDeltas.push(delta);
          continue;
        }
        const hasRow = delta.row != null;
        const hasCol = delta.col != null;
        if (hasCol && !hasRow) rawColStyleDeltas.push(delta);
        else if (hasRow && !hasCol) rawRowStyleDeltas.push(delta);
        else if (!hasRow && !hasCol) rawSheetStyleDeltas.push(delta);
      }
    } else if (formatDeltas && typeof formatDeltas === "object") {
      addAll(rawColStyleDeltas, formatDeltas.colStyleIdDeltas);
      addAll(rawColStyleDeltas, formatDeltas.colStyleDeltas);
      addAll(rawColStyleDeltas, formatDeltas.columnStyleIdDeltas);
      addAll(rawRowStyleDeltas, formatDeltas.rowStyleIdDeltas);
      addAll(rawRowStyleDeltas, formatDeltas.rowStyleDeltas);
      addAll(rawSheetStyleDeltas, formatDeltas.sheetStyleIdDeltas);
      addAll(rawSheetStyleDeltas, formatDeltas.sheetDefaultStyleIdDeltas);
      addAll(rawSheetStyleDeltas, formatDeltas.sheetStyleDeltas);
    }

    // DocumentController changes may also describe layered formatting changes as part of
    // `sheetViewDeltas` (since row/col defaults are sheet-scoped metadata, not per-cell state).
    // Derive per-row/col/sheet style deltas from those view deltas so macro recording still
    // sees formatting edits even when there are no cell deltas.
    const sheetViewDeltas = Array.isArray(payload?.sheetViewDeltas) ? payload.sheetViewDeltas : [];
    for (const delta of sheetViewDeltas) {
      const sheetName = typeof delta?.sheetId === "string" ? delta.sheetId : typeof delta?.sheetName === "string" ? delta.sheetName : null;
      if (!sheetName) continue;

      const beforeView = delta.before ?? {};
      const afterView = delta.after ?? {};

      const beforeDefault = beforeView.defaultFormat ?? beforeView.defaultStyleId ?? 0;
      const afterDefault = afterView.defaultFormat ?? afterView.defaultStyleId ?? 0;
      if (stableStringify(beforeDefault) !== stableStringify(afterDefault)) {
        rawSheetStyleDeltas.push({ sheetId: sheetName, before: beforeDefault, after: afterDefault });
      }

      const beforeCols = styleIndexMapFromRaw(beforeView.colFormats ?? beforeView.colStyleIds);
      const afterCols = styleIndexMapFromRaw(afterView.colFormats ?? afterView.colStyleIds);
      const colKeys = new Set([...beforeCols.keys(), ...afterCols.keys()]);
      for (const col of colKeys) {
        const beforeRef = beforeCols.get(col) ?? 0;
        const afterRef = afterCols.get(col) ?? 0;
        if (stableStringify(beforeRef) === stableStringify(afterRef)) continue;
        rawColStyleDeltas.push({ sheetId: sheetName, col, before: beforeRef, after: afterRef });
      }

      const beforeRows = styleIndexMapFromRaw(beforeView.rowFormats ?? beforeView.rowStyleIds);
      const afterRows = styleIndexMapFromRaw(afterView.rowFormats ?? afterView.rowStyleIds);
      const rowKeys = new Set([...beforeRows.keys(), ...afterRows.keys()]);
      for (const row of rowKeys) {
        const beforeRef = beforeRows.get(row) ?? 0;
        const afterRef = afterRows.get(row) ?? 0;
        if (stableStringify(beforeRef) === stableStringify(afterRef)) continue;
        rawRowStyleDeltas.push({ sheetId: sheetName, row, before: beforeRef, after: afterRef });
      }
    }

    /** @type {Map<string, Map<string, { patch: any, cols: Set<number> }>>} */
    const colGroupsBySheet = new Map();
    for (const delta of rawColStyleDeltas) {
      const sheetName = typeof delta?.sheetId === "string" ? delta.sheetId : typeof delta?.sheetName === "string" ? delta.sheetName : null;
      if (!sheetName) continue;
      const col = Number(delta.col ?? delta.column ?? delta.colIndex ?? delta.columnIndex);
      if (!Number.isInteger(col) || col < 0 || col > EXCEL_MAX_COL) continue;

      const beforeStyle = resolveDocStyle(styleTable, deltaStyleRef(delta, "before"));
      const afterStyle = resolveDocStyle(styleTable, deltaStyleRef(delta, "after"));
      const docPatch = diffStylePatch(beforeStyle, afterStyle);
      const patch = scriptFormatPatchFromDocStylePatch(docPatch);
      if (patch !== null && isPlainObject(patch) && Object.keys(patch).length === 0) continue;

      const key = patch === null ? "null" : stableStringify(patch);
      let groups = colGroupsBySheet.get(sheetName);
      if (!groups) {
        groups = new Map();
        colGroupsBySheet.set(sheetName, groups);
      }
      const entry = groups.get(key) ?? { patch, cols: new Set() };
      entry.cols.add(col);
      groups.set(key, entry);
    }

    for (const [sheetName, groups] of colGroupsBySheet.entries()) {
      for (const entry of groups.values()) {
        const cols = [...entry.cols].sort((a, b) => a - b);
        let i = 0;
        while (i < cols.length) {
          let j = i + 1;
          while (j < cols.length && cols[j] === cols[j - 1] + 1) j += 1;
          const startCol = cols[i];
          const endCol = cols[j - 1];
          this.events.emit("formatChanged", {
            sheetName,
            address: formatRangeAddress({ startRow: 0, startCol, endRow: EXCEL_MAX_ROW, endCol }),
            format: entry.patch,
          });
          i = j;
        }
      }
    }

    /** @type {Map<string, Map<string, { patch: any, rows: Set<number> }>>} */
    const rowGroupsBySheet = new Map();
    for (const delta of rawRowStyleDeltas) {
      const sheetName = typeof delta?.sheetId === "string" ? delta.sheetId : typeof delta?.sheetName === "string" ? delta.sheetName : null;
      if (!sheetName) continue;
      const row = Number(delta.row ?? delta.rowIndex);
      if (!Number.isInteger(row) || row < 0 || row > EXCEL_MAX_ROW) continue;

      const beforeStyle = resolveDocStyle(styleTable, deltaStyleRef(delta, "before"));
      const afterStyle = resolveDocStyle(styleTable, deltaStyleRef(delta, "after"));
      const docPatch = diffStylePatch(beforeStyle, afterStyle);
      const patch = scriptFormatPatchFromDocStylePatch(docPatch);
      if (patch !== null && isPlainObject(patch) && Object.keys(patch).length === 0) continue;

      const key = patch === null ? "null" : stableStringify(patch);
      let groups = rowGroupsBySheet.get(sheetName);
      if (!groups) {
        groups = new Map();
        rowGroupsBySheet.set(sheetName, groups);
      }
      const entry = groups.get(key) ?? { patch, rows: new Set() };
      entry.rows.add(row);
      groups.set(key, entry);
    }

    for (const [sheetName, groups] of rowGroupsBySheet.entries()) {
      for (const entry of groups.values()) {
        const rows = [...entry.rows].sort((a, b) => a - b);
        let i = 0;
        while (i < rows.length) {
          let j = i + 1;
          while (j < rows.length && rows[j] === rows[j - 1] + 1) j += 1;
          const startRow = rows[i];
          const endRow = rows[j - 1];
          this.events.emit("formatChanged", {
            sheetName,
            address: formatRangeAddress({ startRow, startCol: 0, endRow, endCol: EXCEL_MAX_COL }),
            format: entry.patch,
          });
          i = j;
        }
      }
    }

    /** @type {Map<string, Map<string, any>>} */
    const sheetPatchBySheet = new Map();
    for (const delta of rawSheetStyleDeltas) {
      const sheetName = typeof delta?.sheetId === "string" ? delta.sheetId : typeof delta?.sheetName === "string" ? delta.sheetName : null;
      if (!sheetName) continue;

      const beforeStyle = resolveDocStyle(styleTable, deltaStyleRef(delta, "before"));
      const afterStyle = resolveDocStyle(styleTable, deltaStyleRef(delta, "after"));
      const docPatch = diffStylePatch(beforeStyle, afterStyle);
      const patch = scriptFormatPatchFromDocStylePatch(docPatch);
      if (patch !== null && isPlainObject(patch) && Object.keys(patch).length === 0) continue;

      const key = patch === null ? "null" : stableStringify(patch);
      let patches = sheetPatchBySheet.get(sheetName);
      if (!patches) {
        patches = new Map();
        sheetPatchBySheet.set(sheetName, patches);
      }
      patches.set(key, patch);
    }

    for (const [sheetName, patches] of sheetPatchBySheet.entries()) {
      for (const patch of patches.values()) {
        this.events.emit("formatChanged", {
          sheetName,
          address: formatRangeAddress({ startRow: 0, startCol: 0, endRow: EXCEL_MAX_ROW, endCol: EXCEL_MAX_COL }),
          format: patch,
        });
      }
    }
  }

  /**
   * Resolve a scripting-facing sheet identifier (usually a display name) to the
   * underlying DocumentController sheet id.
   *
   * Returning `null` signals "unknown" and MUST NOT cause sheet creation.
   *
   * @param {string} sheetNameOrId
   * @returns {string | null}
   */
  #resolveSheetId(sheetNameOrId) {
    const raw = String(sheetNameOrId ?? "").trim();
    if (!raw) return null;

    // Prefer the UI resolver when available (handles renames).
    if (this.sheetNameResolver) {
      const resolved = this.sheetNameResolver.getSheetIdByName(raw);
      if (resolved) {
        // Avoid resurrecting deleted sheets (or creating phantom sheets) when the resolver is stale.
        // Only accept ids that are known to the underlying DocumentController.
        const meta = this.documentController.getSheetMeta?.(resolved) ?? null;
        if (meta) return resolved;
      }

      // If the caller provided a sheet id, translate id->name->id to canonicalize.
      const resolvedName = this.sheetNameResolver.getSheetNameById(raw);
      if (resolvedName) {
        const roundTrip = this.sheetNameResolver.getSheetIdByName(resolvedName);
        if (roundTrip) {
          const meta = this.documentController.getSheetMeta?.(roundTrip) ?? null;
          if (meta) return roundTrip;
        }
      }
    }

    // Fallback: avoid phantom sheet creation by only allowing known ids.
    const ids = typeof this.documentController.getSheetIds === "function" ? this.documentController.getSheetIds() : [];
    const candidates = ids.length > 0 ? ids : ["Sheet1"];
    return candidates.find((id) => String(id).toLowerCase() === raw.toLowerCase()) ?? null;
  }

  /**
   * @param {string} sheetId
   * @returns {string}
   */
  #sheetDisplayName(sheetId) {
    if (!this.sheetNameResolver) return sheetId;
    return this.sheetNameResolver.getSheetNameById(sheetId) ?? sheetId;
  }
}

class DocumentControllerSheetAdapter {
  /**
   * @param {DocumentControllerWorkbookAdapter} workbook
   * @param {string} sheetId
   */
  constructor(workbook, sheetId) {
    this.workbook = workbook;
    this.sheetId = sheetId;
  }

  /** @type {DocumentControllerWorkbookAdapter} */
  workbook;

  /** @type {string} */
  sheetId;

  get name() {
    return this.workbook.sheetNameResolver?.getSheetNameById(this.sheetId) ?? this.sheetId;
  }

  getRange(address) {
    return new DocumentControllerRangeAdapter(this, parseRangeAddress(String(address)));
  }

  /**
   * Return a single-cell range addressed by 0-based row/col coordinates.
   *
   * Note: scripts generally use A1 addresses; this helper is mainly for API parity with the
   * in-memory workbook implementation.
   */
  getCell(row, col) {
    const r = Number(row);
    const c = Number(col);
    if (!Number.isInteger(r) || r < 0) throw new Error(`getCell expected non-negative integer row, got ${row}`);
    if (!Number.isInteger(c) || c < 0) throw new Error(`getCell expected non-negative integer col, got ${col}`);
    return new DocumentControllerRangeAdapter(this, { startRow: r, startCol: c, endRow: r, endCol: c });
  }

  getUsedRange() {
    const bounds = this.workbook.documentController.getUsedRange(this.sheetId, { includeFormat: true });
    if (!bounds) {
      // Match the in-memory workbook behavior: empty sheets return A1.
      return this.getRange("A1");
    }
    return new DocumentControllerRangeAdapter(this, bounds);
  }

  setCellValue(address, value) {
    return this.getRange(address).setValue(value);
  }

  setRangeValues(address, values) {
    return this.getRange(address).setValues(values);
  }
}

class DocumentControllerRangeAdapter {
  /**
   * @param {DocumentControllerSheetAdapter} sheet
   * @param {{ startRow: number, startCol: number, endRow: number, endCol: number }} coords
   */
  constructor(sheet, coords) {
    this.sheet = sheet;
    this.coords = coords;
  }

  /** @type {DocumentControllerSheetAdapter} */
  sheet;

  /** @type {{ startRow: number, startCol: number, endRow: number, endCol: number }} */
  coords;

  get address() {
    const start = formatCellAddress({ row: this.coords.startRow, col: this.coords.startCol });
    const end = formatCellAddress({ row: this.coords.endRow, col: this.coords.endCol });
    return start === end ? start : `${start}:${end}`;
  }

  getValues() {
    this.#assertSheetExists();
    const rows = this.coords.endRow - this.coords.startRow + 1;
    const cols = this.coords.endCol - this.coords.startCol + 1;
    assertScriptRangeWithinLimit(rows, cols, this.address, "getValues");
    const out = [];
    for (let r = 0; r < rows; r++) {
      const row = [];
      for (let c = 0; c < cols; c++) {
        const cell = this.sheet.workbook.documentController.getCell(this.sheet.sheetId, {
          row: this.coords.startRow + r,
          col: this.coords.startCol + c,
        });
        row.push(cellInputFromState(cell));
      }
      out.push(row);
    }
    return out;
  }

  getFormulas() {
    this.#assertSheetExists();
    const rows = this.coords.endRow - this.coords.startRow + 1;
    const cols = this.coords.endCol - this.coords.startCol + 1;
    assertScriptRangeWithinLimit(rows, cols, this.address, "getFormulas");
    const out = [];
    for (let r = 0; r < rows; r++) {
      const row = [];
      for (let c = 0; c < cols; c++) {
        const cell = this.sheet.workbook.documentController.getCell(this.sheet.sheetId, {
          row: this.coords.startRow + r,
          col: this.coords.startCol + c,
        });
        row.push(cell.formula != null ? normalizeFormulaText(cell.formula) : null);
      }
      out.push(row);
    }
    return out;
  }

  setValues(values) {
    this.#assertSheetExists();
    const rows = this.coords.endRow - this.coords.startRow + 1;
    const cols = this.coords.endCol - this.coords.startCol + 1;
    assertScriptRangeWithinLimit(rows, cols, this.address, "setValues");
    if (!Array.isArray(values) || values.length !== rows || values.some((row) => row.length !== cols)) {
      throw new Error(
        `setValues expected ${rows}x${cols} matrix for range ${this.address}, got ${values.length}x${values[0]?.length ?? 0}`,
      );
    }

    this.#assertCanEditRange({ rows, cols }, { action: "edit", kind: "cell" });

    this.sheet.workbook.documentController.setRangeValues(this.sheet.sheetId, this.address, values, {
      label: "Script: set values",
    });
    this.sheet.workbook._notifyMutate();
  }

  setFormulas(formulas) {
    this.#assertSheetExists();
    const rows = this.coords.endRow - this.coords.startRow + 1;
    const cols = this.coords.endCol - this.coords.startCol + 1;
    assertScriptRangeWithinLimit(rows, cols, this.address, "setFormulas");
    if (!Array.isArray(formulas) || formulas.length !== rows || formulas.some((row) => row.length !== cols)) {
      throw new Error(
        `setFormulas expected ${rows}x${cols} matrix for range ${this.address}, got ${formulas.length}x${formulas[0]?.length ?? 0}`,
      );
    }

    this.#assertCanEditRange({ rows, cols }, { action: "edit", kind: "cell" });

    const values = formulas.map((row) =>
      row.map((formula) => {
        if (formula == null) return null;
        return normalizeFormulaText(String(formula));
      }),
    );

    this.sheet.workbook.documentController.setRangeValues(this.sheet.sheetId, this.address, values, {
      label: "Script: set formulas",
    });
    this.sheet.workbook._notifyMutate();
  }

  getValue() {
    const values = this.getValues();
    if (values.length !== 1 || values[0].length !== 1) {
      throw new Error(`getValue is only valid for a single cell, got range ${this.address}`);
    }
    return values[0][0];
  }

  setValue(value) {
    this.#assertSheetExists();
    const range = this.coords;
    if (range.startRow !== range.endRow || range.startCol !== range.endCol) {
      throw new Error(`setValue is only valid for a single cell, got range ${this.address}`);
    }

    const coord = { row: range.startRow, col: range.startCol };
    this.#assertCanEditCell(coord.row, coord.col, { action: "edit", kind: "cell" });

    if (typeof value === "string") {
      if (value.startsWith("'")) {
        this.sheet.workbook.documentController.setCellValue(this.sheet.sheetId, coord, value.slice(1), {
          label: "Script: set value",
        });
        this.sheet.workbook._notifyMutate();
        return;
      }
      if (isFormulaString(value)) {
        this.sheet.workbook.documentController.setCellFormula(this.sheet.sheetId, coord, value, {
          label: "Script: set value",
        });
        this.sheet.workbook._notifyMutate();
        return;
      }
    }

    this.sheet.workbook.documentController.setCellValue(this.sheet.sheetId, coord, value ?? null, { label: "Script: set value" });
    this.sheet.workbook._notifyMutate();
  }

  getFormat() {
    this.#assertSheetExists();
    const doc = this.sheet.workbook.documentController;
    const coord = { row: this.coords.startRow, col: this.coords.startCol };

    // Prefer DocumentController's effective/layered formatting API when available so
    // scripts observe inherited row/col/sheet formats (not just cell-level overrides).
    if (typeof doc.getCellFormat === "function") {
      const effective = doc.getCellFormat(this.sheet.sheetId, coord);
      // Some implementations may return a styleId instead of a style object.
      if (typeof effective === "number") {
        return scriptFormatFromDocStyle(doc.styleTable?.get(effective) ?? {});
      }
      if (isPlainObject(effective) && typeof effective.styleId === "number") {
        return scriptFormatFromDocStyle(doc.styleTable?.get(effective.styleId) ?? {});
      }
      return scriptFormatFromDocStyle(effective);
    }

    const cell = doc.getCell(this.sheet.sheetId, coord);
    const style = doc.styleTable.get(cell.styleId);
    return scriptFormatFromDocStyle(style);
  }

  getFormats() {
    this.#assertSheetExists();
    const doc = this.sheet.workbook.documentController;
    const rows = this.coords.endRow - this.coords.startRow + 1;
    const cols = this.coords.endCol - this.coords.startCol + 1;
    assertScriptRangeWithinLimit(rows, cols, this.address, "getFormats");
    const out = [];

    for (let r = 0; r < rows; r++) {
      const row = [];
      for (let c = 0; c < cols; c++) {
        const coord = { row: this.coords.startRow + r, col: this.coords.startCol + c };

        if (typeof doc.getCellFormat === "function") {
          const effective = doc.getCellFormat(this.sheet.sheetId, coord);
          if (typeof effective === "number") {
            row.push(scriptFormatFromDocStyle(doc.styleTable?.get(effective) ?? {}));
            continue;
          }
          if (isPlainObject(effective) && typeof effective.styleId === "number") {
            row.push(scriptFormatFromDocStyle(doc.styleTable?.get(effective.styleId) ?? {}));
            continue;
          }
          row.push(scriptFormatFromDocStyle(effective));
          continue;
        }

        const cell = doc.getCell(this.sheet.sheetId, coord);
        const style = doc.styleTable.get(cell.styleId);
        row.push(scriptFormatFromDocStyle(style));
      }
      out.push(row);
    }

    return out;
  }

  setFormat(format) {
    this.#assertSheetExists();
    const rows = this.coords.endRow - this.coords.startRow + 1;
    const cols = this.coords.endCol - this.coords.startCol + 1;
    // Most script formatting calls are small; check every cell so a partially protected range fails
    // early instead of silently filtering out changes.
    //
    // For very large ranges, avoid O(area) permission scans and instead check the anchor cell.
    this.#assertCanEditRange({ rows, cols }, { action: "format", kind: "format" });
    const patch = docStylePatchFromScriptFormat(format);
    const applied = this.sheet.workbook.documentController.setRangeFormat(this.sheet.sheetId, this.address, patch, {
      label: "Script: set format",
    });
    if (applied === false) {
      const area = rows * cols;
      throw new Error(
        `setFormat skipped for range ${this.address} (rows=${rows}, cols=${cols}, area=${area}). ` +
          "Select fewer cells/rows and try again.",
      );
    }
    this.sheet.workbook._notifyMutate();
  }

  setFormats(formats) {
    this.#assertSheetExists();
    const rows = this.coords.endRow - this.coords.startRow + 1;
    const cols = this.coords.endCol - this.coords.startCol + 1;
    assertScriptRangeWithinLimit(rows, cols, this.address, "setFormats");
    if (!Array.isArray(formats) || formats.length !== rows || formats.some((row) => !Array.isArray(row) || row.length !== cols)) {
      throw new Error(
        `setFormats expected ${rows}x${cols} matrix for range ${this.address}, got ${formats?.length ?? 0}x${formats?.[0]?.length ?? 0}`,
      );
    }

    this.#assertCanEditRange({ rows, cols }, { action: "format", kind: "format" });

    const doc = this.sheet.workbook.documentController;
    const styleTable = doc.styleTable;

    const values = [];
    for (let r = 0; r < rows; r++) {
      const row = [];
      for (let c = 0; c < cols; c++) {
        const coord = { row: this.coords.startRow + r, col: this.coords.startCol + c };
        const patch = docStylePatchFromScriptFormat(formats[r][c]);
        if (patch == null) {
          row.push({ styleId: 0 });
          continue;
        }

        const cell = doc.getCell(this.sheet.sheetId, coord);
        const baseStyle = styleTable.get(cell.styleId ?? 0);
        const merged = applyStylePatch(baseStyle, patch);
        const styleId = styleTable.intern(merged);
        row.push({ styleId });
      }
      values.push(row);
    }

    doc.setRangeValues(this.sheet.sheetId, this.address, values, { label: "Script: set formats" });
    this.sheet.workbook._notifyMutate();
  }

  #assertSheetExists() {
    const doc = this.sheet.workbook.documentController;
    const sheetId = this.sheet.sheetId;
    const sheets = doc?.model?.sheets;
    const meta = doc?.sheetMeta;
    // If we can't inspect sheet registries, be permissive (supports unit tests with lightweight
    // DocumentController mocks and avoids hard dependency on internal controller shape).
    if (!(sheets instanceof Map) && !(meta instanceof Map)) return;
    // DocumentController materializes sheets lazily; treat the default Sheet1 as existing
    // even when the controller hasn't created any sheets yet.
    if (sheets instanceof Map) {
      if (sheets.size === 0 && String(sheetId).toLowerCase() === "sheet1") return;
      if (sheets.has(sheetId)) return;
    }
    if (meta instanceof Map && meta.has(sheetId)) return;
    throw new Error(`Unknown sheet: ${this.sheet.name}`);
  }

  /**
   * @param {number} row
   * @param {number} col
   * @param {{ action: string, kind: "cell" | "format" }} params
   */
  #assertCanEditCell(row, col, params) {
    this.#assertSheetExists();
    const doc = this.sheet.workbook.documentController;
    if (typeof doc?.canEditCell !== "function") return;
    if (doc.canEditCell({ sheetId: this.sheet.sheetId, row, col })) return;

    const address = formatCellAddress({ row, col });
    if (params.kind === "format") {
      throw new Error(`Read-only: you don't have permission to change formatting (${address})`);
    }
    throw new Error(`Read-only: you don't have permission to edit that cell (${address})`);
  }

  /**
   * @param {{ rows: number, cols: number }} size
   * @param {{ action: string, kind: "cell" | "format" }} params
   */
  #assertCanEditRange(size, params) {
    // Always validate the anchor cell first so read-only users fail quickly even for huge ranges.
    this.#assertCanEditCell(this.coords.startRow, this.coords.startCol, params);

    const rows = Number(size?.rows);
    const cols = Number(size?.cols);
    const cellCount = rows * cols;
    if (!Number.isFinite(cellCount) || cellCount <= 1) return;

    // Avoid O(area) checks for giant selections; the anchor check is sufficient to catch the common
    // "viewer/commenter" (all cells blocked) case without freezing the renderer.
    if (cellCount > DEFAULT_SCRIPT_RANGE_CELL_LIMIT) return;

    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        if (r === 0 && c === 0) continue; // already checked anchor
        this.#assertCanEditCell(this.coords.startRow + r, this.coords.startCol + c, params);
      }
    }
  }
}
