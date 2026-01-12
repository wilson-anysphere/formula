import { formatCellAddress, formatRangeAddress, parseRangeAddress } from "../../../../packages/scripting/src/a1.js";
import { TypedEventEmitter } from "../../../../packages/scripting/src/events.js";

// Excel limits used by the scripting A1 address helpers and macro recorder.
// - Rows: 1..1048576 (0-based: 0..1048575)
// - Cols: A..XFD (0-based: 0..16383)
const EXCEL_MAX_ROW = 1048575;
const EXCEL_MAX_COL = 16383;

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

function cellInputFromState(state) {
  if (state.formula != null) return normalizeFormulaText(state.formula);
  const value = state.value ?? null;
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

  if (hasOwn(style, "numberFormat") && typeof style.numberFormat === "string") {
    out.numberFormat = style.numberFormat;
  }

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
    let sheet = this.sheets.get(sheetName);
    if (!sheet) {
      sheet = new DocumentControllerSheetAdapter(this, sheetName);
      this.sheets.set(sheetName, sheet);
    }
    return sheet;
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
      return { sheetName: String(sel.sheetName), address: String(sel.address) };
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

    this.selection = { sheetName: normalizedSheetName, address: normalizedAddress };
    this.events.emit("selectionChanged", this.selection);
    this.setSelectionImpl?.(normalizedSheetName, normalizedAddress);
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
      const sheetName = delta.sheetId;
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

    /** @type {Map<string, Map<string, { patch: any, cols: number[] }>>} */
    const colGroupsBySheet = new Map();
    for (const delta of rawColStyleDeltas) {
      const sheetName = typeof delta?.sheetId === "string" ? delta.sheetId : typeof delta?.sheetName === "string" ? delta.sheetName : null;
      if (!sheetName) continue;
      const col = Number(delta.col ?? delta.column ?? delta.colIndex ?? delta.columnIndex);
      if (!Number.isInteger(col) || col < 0) continue;

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
      const entry = groups.get(key) ?? { patch, cols: [] };
      entry.cols.push(col);
      groups.set(key, entry);
    }

    for (const [sheetName, groups] of colGroupsBySheet.entries()) {
      for (const entry of groups.values()) {
        entry.cols.sort((a, b) => a - b);
        let i = 0;
        while (i < entry.cols.length) {
          let j = i + 1;
          while (j < entry.cols.length && entry.cols[j] === entry.cols[j - 1] + 1) j += 1;
          const startCol = entry.cols[i];
          const endCol = entry.cols[j - 1];
          this.events.emit("formatChanged", {
            sheetName,
            address: formatRangeAddress({ startRow: 0, startCol, endRow: EXCEL_MAX_ROW, endCol }),
            format: entry.patch,
          });
          i = j;
        }
      }
    }

    /** @type {Map<string, Map<string, { patch: any, rows: number[] }>>} */
    const rowGroupsBySheet = new Map();
    for (const delta of rawRowStyleDeltas) {
      const sheetName = typeof delta?.sheetId === "string" ? delta.sheetId : typeof delta?.sheetName === "string" ? delta.sheetName : null;
      if (!sheetName) continue;
      const row = Number(delta.row ?? delta.rowIndex);
      if (!Number.isInteger(row) || row < 0) continue;

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
      const entry = groups.get(key) ?? { patch, rows: [] };
      entry.rows.push(row);
      groups.set(key, entry);
    }

    for (const [sheetName, groups] of rowGroupsBySheet.entries()) {
      for (const entry of groups.values()) {
        entry.rows.sort((a, b) => a - b);
        let i = 0;
        while (i < entry.rows.length) {
          let j = i + 1;
          while (j < entry.rows.length && entry.rows[j] === entry.rows[j - 1] + 1) j += 1;
          const startRow = entry.rows[i];
          const endRow = entry.rows[j - 1];
          this.events.emit("formatChanged", {
            sheetName,
            address: formatRangeAddress({ startRow, startCol: 0, endRow, endCol: EXCEL_MAX_COL }),
            format: entry.patch,
          });
          i = j;
        }
      }
    }

    for (const delta of rawSheetStyleDeltas) {
      const sheetName = typeof delta?.sheetId === "string" ? delta.sheetId : typeof delta?.sheetName === "string" ? delta.sheetName : null;
      if (!sheetName) continue;

      const beforeStyle = resolveDocStyle(styleTable, deltaStyleRef(delta, "before"));
      const afterStyle = resolveDocStyle(styleTable, deltaStyleRef(delta, "after"));
      const docPatch = diffStylePatch(beforeStyle, afterStyle);
      const patch = scriptFormatPatchFromDocStylePatch(docPatch);
      if (patch !== null && isPlainObject(patch) && Object.keys(patch).length === 0) continue;

      this.events.emit("formatChanged", {
        sheetName,
        address: formatRangeAddress({ startRow: 0, startCol: 0, endRow: EXCEL_MAX_ROW, endCol: EXCEL_MAX_COL }),
        format: patch,
      });
    }
  }
}

class DocumentControllerSheetAdapter {
  /**
   * @param {DocumentControllerWorkbookAdapter} workbook
   * @param {string} name
   */
  constructor(workbook, name) {
    this.workbook = workbook;
    this.name = name;
  }

  /** @type {DocumentControllerWorkbookAdapter} */
  workbook;

  /** @type {string} */
  name;

  getRange(address) {
    return new DocumentControllerRangeAdapter(this, parseRangeAddress(String(address)));
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
    const rows = this.coords.endRow - this.coords.startRow + 1;
    const cols = this.coords.endCol - this.coords.startCol + 1;
    const out = [];
    for (let r = 0; r < rows; r++) {
      const row = [];
      for (let c = 0; c < cols; c++) {
        const cell = this.sheet.workbook.documentController.getCell(this.sheet.name, {
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
    const rows = this.coords.endRow - this.coords.startRow + 1;
    const cols = this.coords.endCol - this.coords.startCol + 1;
    const out = [];
    for (let r = 0; r < rows; r++) {
      const row = [];
      for (let c = 0; c < cols; c++) {
        const cell = this.sheet.workbook.documentController.getCell(this.sheet.name, {
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
    const rows = this.coords.endRow - this.coords.startRow + 1;
    const cols = this.coords.endCol - this.coords.startCol + 1;
    if (!Array.isArray(values) || values.length !== rows || values.some((row) => row.length !== cols)) {
      throw new Error(
        `setValues expected ${rows}x${cols} matrix for range ${this.address}, got ${values.length}x${values[0]?.length ?? 0}`,
      );
    }

    this.sheet.workbook.documentController.setRangeValues(this.sheet.name, this.address, values, {
      label: "Script: set values",
    });
    this.sheet.workbook._notifyMutate();
  }

  setFormulas(formulas) {
    const rows = this.coords.endRow - this.coords.startRow + 1;
    const cols = this.coords.endCol - this.coords.startCol + 1;
    if (!Array.isArray(formulas) || formulas.length !== rows || formulas.some((row) => row.length !== cols)) {
      throw new Error(
        `setFormulas expected ${rows}x${cols} matrix for range ${this.address}, got ${formulas.length}x${formulas[0]?.length ?? 0}`,
      );
    }

    const values = formulas.map((row) =>
      row.map((formula) => {
        if (formula == null) return null;
        return normalizeFormulaText(String(formula));
      }),
    );

    this.sheet.workbook.documentController.setRangeValues(this.sheet.name, this.address, values, {
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
    const range = this.coords;
    if (range.startRow !== range.endRow || range.startCol !== range.endCol) {
      throw new Error(`setValue is only valid for a single cell, got range ${this.address}`);
    }

    const coord = { row: range.startRow, col: range.startCol };

    if (typeof value === "string") {
      if (value.startsWith("'")) {
        this.sheet.workbook.documentController.setCellValue(this.sheet.name, coord, value.slice(1), {
          label: "Script: set value",
        });
        this.sheet.workbook._notifyMutate();
        return;
      }
      if (isFormulaString(value)) {
        this.sheet.workbook.documentController.setCellFormula(this.sheet.name, coord, value, {
          label: "Script: set value",
        });
        this.sheet.workbook._notifyMutate();
        return;
      }
    }

    this.sheet.workbook.documentController.setCellValue(this.sheet.name, coord, value ?? null, { label: "Script: set value" });
    this.sheet.workbook._notifyMutate();
  }

  getFormat() {
    const doc = this.sheet.workbook.documentController;
    const coord = { row: this.coords.startRow, col: this.coords.startCol };

    // Prefer DocumentController's effective/layered formatting API when available so
    // scripts observe inherited row/col/sheet formats (not just cell-level overrides).
    if (typeof doc.getCellFormat === "function") {
      const effective = doc.getCellFormat(this.sheet.name, coord);
      // Some implementations may return a styleId instead of a style object.
      if (typeof effective === "number") {
        return scriptFormatFromDocStyle(doc.styleTable?.get(effective) ?? {});
      }
      if (isPlainObject(effective) && typeof effective.styleId === "number") {
        return scriptFormatFromDocStyle(doc.styleTable?.get(effective.styleId) ?? {});
      }
      return scriptFormatFromDocStyle(effective);
    }

    const cell = doc.getCell(this.sheet.name, coord);
    const style = doc.styleTable.get(cell.styleId);
    return scriptFormatFromDocStyle(style);
  }

  setFormat(format) {
    const patch = docStylePatchFromScriptFormat(format);
    this.sheet.workbook.documentController.setRangeFormat(this.sheet.name, this.address, patch, { label: "Script: set format" });
    this.sheet.workbook._notifyMutate();
  }
}
