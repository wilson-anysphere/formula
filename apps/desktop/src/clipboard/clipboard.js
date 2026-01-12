import { normalizeRange, parseA1, parseRangeA1 } from "../document/coords.js";
import { t } from "../i18n/index.js";
import { parseHtmlToCellGrid, serializeCellGridToHtml } from "./html.js";
import { parseTsvToCellGrid, serializeCellGridToTsv } from "./tsv.js";
import { enforceClipboardCopy } from "../dlp/enforceClipboardCopy.js";

/**
 * @typedef {import("./types.js").CellGrid} CellGrid
 * @typedef {import("./types.js").ClipboardContent} ClipboardContent
 * @typedef {import("./types.js").ClipboardWritePayload} ClipboardWritePayload
 * @typedef {import("./types.js").PasteOptions} PasteOptions
 *
 * @typedef {import("../document/coords.js").CellCoord} CellCoord
 * @typedef {import("../document/coords.js").CellRange} CellRange
 * @typedef {import("../document/documentController.js").DocumentController} DocumentController
 * @typedef {import("../document/cell.js").CellState} CellState
 */

function normalizeHexToCssColor(hex) {
  const normalized = String(hex).trim().replace(/^#/, "");
  if (!/^[0-9a-fA-F]+$/.test(normalized)) return null;

  // #RRGGBB
  if (normalized.length === 6) return `#${normalized}`;

  // Excel/OOXML commonly stores colors as ARGB (AARRGGBB).
  if (normalized.length === 8) {
    const a = Number.parseInt(normalized.slice(0, 2), 16);
    const r = Number.parseInt(normalized.slice(2, 4), 16);
    const g = Number.parseInt(normalized.slice(4, 6), 16);
    const b = Number.parseInt(normalized.slice(6, 8), 16);

    if (![a, r, g, b].every((n) => Number.isFinite(n))) return null;

    if (a >= 255) {
      return `#${normalized.slice(2)}`;
    }

    const alpha = Math.max(0, Math.min(1, a / 255));
    const rounded = Math.round(alpha * 1000) / 1000;
    return `rgba(${r}, ${g}, ${b}, ${rounded})`;
  }

  return null;
}

function normalizeCssColor(value) {
  if (typeof value !== "string") return null;
  const trimmed = value.trim();
  if (!trimmed) return null;

  // Accept plain CSS colors as-is.
  if (!/^[0-9a-fA-F#]+$/.test(trimmed)) return trimmed;

  return normalizeHexToCssColor(trimmed) ?? trimmed;
}

function stableStringify(value) {
  if (value === undefined) return "undefined";
  if (value == null || typeof value !== "object") return JSON.stringify(value);
  if (Array.isArray(value)) return `[${value.map(stableStringify).join(",")}]`;
  const keys = Object.keys(value).sort();
  const entries = keys.map((k) => `${JSON.stringify(k)}:${stableStringify(value[k])}`);
  return `{${entries.join(",")}}`;
}

/**
 * Convert a DocumentController style table entry to the clipboard's lightweight format.
 *
 * Clipboard HTML serialization expects flat keys like `bold`, `textColor`, `backgroundColor`,
 * and `numberFormat`, whereas the DocumentController stores richer OOXML-ish styles
 * (`font.bold`, `fill.fgColor`, ...).
 *
 * @param {any} style
 * @returns {any | null}
 */
function styleToClipboardFormat(style) {
  if (!style || typeof style !== "object") return null;

  /** @type {any} */
  const out = {};

  const font = style.font;
  if (font && typeof font === "object") {
    if (typeof font.bold === "boolean") out.bold = font.bold;
    if (typeof font.italic === "boolean") out.italic = font.italic;
    if (typeof font.underline === "boolean") out.underline = font.underline;
    if (typeof font.underline === "string") out.underline = font.underline !== "none";

    const color = normalizeCssColor(font.color);
    if (color) out.textColor = color;
  }

  const fill = style.fill;
  if (fill && typeof fill === "object") {
    const raw = fill.fgColor ?? fill.background ?? fill.bgColor;
    const color = normalizeCssColor(raw);
    if (color) out.backgroundColor = color;
  }

  const rawNumberFormat = style.numberFormat ?? style.number_format;
  if (typeof rawNumberFormat === "string" && rawNumberFormat.trim() !== "") {
    out.numberFormat = rawNumberFormat;
  }

  // Back-compat: allow flat clipboard-ish styles to round trip.
  if (out.bold === undefined && typeof style.bold === "boolean") out.bold = style.bold;
  if (out.italic === undefined && typeof style.italic === "boolean") out.italic = style.italic;
  if (out.underline === undefined && typeof style.underline === "boolean") out.underline = style.underline;
  if (out.textColor === undefined) {
    const color = normalizeCssColor(style.textColor ?? style.text_color ?? style.fontColor ?? style.font_color);
    if (color) out.textColor = color;
  }
  if (out.backgroundColor === undefined) {
    const color = normalizeCssColor(
      style.backgroundColor ?? style.background_color ?? style.fillColor ?? style.fill_color
    );
    if (color) out.backgroundColor = color;
  }
  if (out.numberFormat === undefined) {
    const nf = style.numberFormat ?? style.number_format;
    if (typeof nf === "string" && nf.trim() !== "") out.numberFormat = nf;
  }

  return Object.keys(out).length > 0 ? out : null;
}

/**
 * @param {CellGrid} grid
 * @returns {ClipboardWritePayload}
 */
export function serializeCellGridToClipboardPayload(grid) {
  return {
    text: serializeCellGridToTsv(grid),
    html: serializeCellGridToHtml(grid),
  };
}

/**
 * Parse clipboard payloads in priority order: HTML > TSV/plain.
 * @param {ClipboardContent} content
 * @returns {CellGrid | null}
 */
export function parseClipboardContentToCellGrid(content) {
  const html = content.html?.trim();
  if (html) {
    const parsed = parseHtmlToCellGrid(html);
    if (parsed) return parsed;
  }

  const text = content.text;
  if (typeof text === "string") return parseTsvToCellGrid(text);

  return null;
}

/**
 * Extract a rectangular cell grid from the document controller.
 *
 * @param {DocumentController} doc
 * @param {string} sheetId
 * @param {CellRange | string} range
 * @returns {CellGrid}
 */
export function getCellGridFromRange(doc, sheetId, range) {
  const r = typeof range === "string" ? parseRangeA1(range) : normalizeRange(range);

  /** @type {CellGrid} */
  const grid = [];
  /** @type {Map<string, any>} */
  const formatCache = new Map();

  /**
   * Normalize style ids returned from controller helpers / internal state.
   * @param {any} value
   */
  const normalizeStyleId = (value) => {
    const n = Number(value);
    return Number.isInteger(n) && n >= 0 ? n : 0;
  };

  // `getSheetView` exists on the legacy DocumentController, but we only want to
  // consult it for layered style ids when it actually exposes them. Cache the
  // view so we don't clone/allocate it for every cell in the range.
  let cachedSheetView = null;
  let cachedSheetViewLoaded = false;
  let cachedSheetViewHasStyleLayers = false;
  const getSheetViewForStyleLayers = () => {
    if (cachedSheetViewLoaded) return cachedSheetView;
    cachedSheetViewLoaded = true;
    if (doc && typeof doc.getSheetView === "function") {
      cachedSheetView = doc.getSheetView(sheetId);
      cachedSheetViewHasStyleLayers = Boolean(
        cachedSheetView &&
          (cachedSheetView.sheetDefaultStyleId != null ||
            cachedSheetView.defaultStyleId != null ||
            cachedSheetView.sheetStyleId != null ||
            cachedSheetView.rowStyleIds != null ||
            cachedSheetView.rowStyles != null ||
            cachedSheetView.rowStyleIdByRow != null ||
            cachedSheetView.colStyleIds != null ||
            cachedSheetView.colStyles != null ||
            cachedSheetView.colStyleIdByCol != null)
      );
    }
    return cachedSheetView;
  };

  /**
   * Resolve the set of style ids contributing to a cell's effective format.
   *
   * Prefer controller helpers when available; otherwise fall back to best-effort
   * inspection of common internal state shapes.
   *
   * @param {number} row
   * @param {number} col
   * @param {number} cellStyleId
   * @returns {[number, number, number, number]}
   */
  const getCellStyleIdTuple = (row, col, cellStyleId) => {
    const coord = { row, col };

    // Preferred: helper that returns the tuple directly.
    /** @type {any} */
    const helper =
      (doc && typeof doc.getCellFormatStyleIds === "function" && doc.getCellFormatStyleIds) ||
      (doc && typeof doc.getCellFormatIds === "function" && doc.getCellFormatIds) ||
      (doc && typeof doc.getCellFormatStyleIdTuple === "function" && doc.getCellFormatStyleIdTuple) ||
      (doc && typeof doc.getCellStyleIdTuple === "function" && doc.getCellStyleIdTuple);

    if (helper) {
      const out = helper.call(doc, sheetId, coord);
      if (Array.isArray(out) && out.length >= 4) {
        return [
          normalizeStyleId(out[0]),
          normalizeStyleId(out[1]),
          normalizeStyleId(out[2]),
          normalizeStyleId(out[3]),
        ];
      }
      if (out && typeof out === "object") {
        const sheetDefaultStyleId = normalizeStyleId(
          out.sheetDefaultStyleId ?? out.sheetStyleId ?? out.defaultStyleId ?? out.sheetDefault
        );
        const rowStyleId = normalizeStyleId(out.rowStyleId ?? out.rowDefaultStyleId ?? out.rowDefault);
        const colStyleId = normalizeStyleId(out.colStyleId ?? out.colDefaultStyleId ?? out.colDefault);
        const cellId = normalizeStyleId(out.cellStyleId ?? out.styleId ?? cellStyleId);
        return [sheetDefaultStyleId, rowStyleId, colStyleId, cellId];
      }
    }

    // Best-effort fallback: query per-layer style ids if exposed.
    let sheetDefaultStyleId = 0;
    let rowStyleId = 0;
    let colStyleId = 0;

    if (doc && typeof doc.getSheetDefaultStyleId === "function") {
      sheetDefaultStyleId = normalizeStyleId(doc.getSheetDefaultStyleId(sheetId));
    } else if (doc && typeof doc.getSheetView === "function") {
      const view = getSheetViewForStyleLayers();
      if (cachedSheetViewHasStyleLayers) {
        sheetDefaultStyleId = normalizeStyleId(
          view?.sheetDefaultStyleId ?? view?.defaultStyleId ?? view?.sheetStyleId
        );

        const rowStyles = view?.rowStyleIds ?? view?.rowStyles ?? view?.rowStyleIdByRow;
        if (rowStyles) {
          if (typeof rowStyles.get === "function") rowStyleId = normalizeStyleId(rowStyles.get(row));
          else rowStyleId = normalizeStyleId(rowStyles[String(row)] ?? rowStyles[row]);
        }

        const colStyles = view?.colStyleIds ?? view?.colStyles ?? view?.colStyleIdByCol;
        if (colStyles) {
          if (typeof colStyles.get === "function") colStyleId = normalizeStyleId(colStyles.get(col));
          else colStyleId = normalizeStyleId(colStyles[String(col)] ?? colStyles[col]);
        }
      }
    }

    if (doc && typeof doc.getRowStyleId === "function") {
      rowStyleId = normalizeStyleId(doc.getRowStyleId(sheetId, row));
    }
    if (doc && typeof doc.getColStyleId === "function") {
      colStyleId = normalizeStyleId(doc.getColStyleId(sheetId, col));
    }

    return [sheetDefaultStyleId, rowStyleId, colStyleId, normalizeStyleId(cellStyleId)];
  };

  /**
   * Resolve the effective style object for a given cell.
   *
   * @param {number} row
   * @param {number} col
   * @param {number} cellStyleId
   * @returns {any | null}
   */
  const getEffectiveStyle = (row, col, cellStyleId) => {
    const coord = { row, col };
    if (doc && typeof doc.getCellFormat === "function") {
      return doc.getCellFormat(sheetId, coord);
    }
    if (doc && typeof doc.getEffectiveCellStyle === "function") {
      return doc.getEffectiveCellStyle(sheetId, coord);
    }
    if (doc && typeof doc.getCellStyle === "function") {
      return doc.getCellStyle(sheetId, coord);
    }

    // Legacy DocumentController: per-cell styleId only.
    return doc?.styleTable?.get && typeof doc.styleTable.get === "function" ? doc.styleTable.get(cellStyleId) : null;
  };

  const hasStyleIdTupleHelper =
    doc &&
    (typeof doc.getCellFormatStyleIds === "function" ||
      typeof doc.getCellFormatIds === "function" ||
      typeof doc.getCellFormatStyleIdTuple === "function" ||
      typeof doc.getCellStyleIdTuple === "function");
  const hasPerLayerStyleIdMethods =
    doc &&
    (typeof doc.getSheetDefaultStyleId === "function" ||
      typeof doc.getRowStyleId === "function" ||
      typeof doc.getColStyleId === "function");

  // Initialize cachedSheetViewHasStyleLayers (at most one call).
  getSheetViewForStyleLayers();
  const canDeriveStyleIdTuple = Boolean(hasStyleIdTupleHelper || hasPerLayerStyleIdMethods || cachedSheetViewHasStyleLayers);

  for (let row = r.start.row; row <= r.end.row; row++) {
    /** @type {CellState[]} */
    const outRow = [];
    for (let col = r.start.col; col <= r.end.col; col++) {
      const cell = doc.getCell(sheetId, { row, col });

      const cellStyleId = typeof cell?.styleId === "number" ? cell.styleId : 0;

      // If the controller can supply the contributing style ids, cache by the
      // (sheet,row,col,cell) style-id tuple. Otherwise, fall back to caching by
      // the resolved style object to avoid collisions.
      let format = null;
      if (doc && typeof doc.getCellFormat === "function" && !canDeriveStyleIdTuple) {
        const style = getEffectiveStyle(row, col, cellStyleId);
        const cacheKey = stableStringify(style);
        if (formatCache.has(cacheKey)) {
          format = formatCache.get(cacheKey) ?? null;
        } else {
          format = styleToClipboardFormat(style);
          formatCache.set(cacheKey, format);
        }
      } else {
        const styleIdTuple = getCellStyleIdTuple(row, col, cellStyleId);
        const hasAnyStyle = styleIdTuple.some((id) => id !== 0);
        if (hasAnyStyle) {
          const cacheKey = styleIdTuple.join(",");
          if (formatCache.has(cacheKey)) {
            format = formatCache.get(cacheKey) ?? null;
          } else {
            const style = getEffectiveStyle(row, col, cellStyleId);
            format = styleToClipboardFormat(style);
            formatCache.set(cacheKey, format);
          }
        }
      }
      outRow.push({ value: cell.value, formula: cell.formula, format });
    }
    grid.push(outRow);
  }

  return grid;
}

/**
 * Copy a document range into an Excel-compatible clipboard payload (TSV + HTML table).
 *
 * @param {DocumentController} doc
 * @param {string} sheetId
 * @param {CellRange | string} range
 * @param {{ dlp?: { documentId: string, classificationStore: any, policy: any } }} [options]
 * @returns {ClipboardWritePayload}
 */
export function copyRangeToClipboardPayload(doc, sheetId, range, options = {}) {
  const r = typeof range === "string" ? parseRangeA1(range) : normalizeRange(range);
  const dlp = options.dlp;

  if (dlp && typeof dlp === "object") {
    enforceClipboardCopy({
      documentId: dlp.documentId,
      sheetId,
      range: r,
      classificationStore: dlp.classificationStore,
      policy: dlp.policy,
    });
  }

  const grid = getCellGridFromRange(doc, sheetId, r);
  return serializeCellGridToClipboardPayload(grid);
}

/**
 * Paste clipboard content into the document controller at a start cell.
 *
 * @param {DocumentController} doc
 * @param {string} sheetId
 * @param {CellCoord | string} start
 * @param {ClipboardContent} content
 * @param {PasteOptions} [options]
 * @returns {boolean} whether a paste occurred
 */
export function pasteClipboardContent(doc, sheetId, start, content, options = {}) {
  const grid = parseClipboardContentToCellGrid(content);
  if (!grid) return false;

  const mode = options.mode ?? "all";

  const values = grid.map((row) =>
    row.map((cell) => {
      if (mode === "values") return cell.value ?? null;
      if (mode === "formulas") return cell.formula != null ? { formula: cell.formula } : cell.value ?? null;
      if (mode === "formats") return { format: cell.format ?? null };

      // mode === "all"
      if (cell.formula != null) return { formula: cell.formula, value: null, format: cell.format ?? null };
      return { value: cell.value ?? null, format: cell.format ?? null };
    })
  );

  const startCoord = typeof start === "string" ? parseA1(start) : start;
  doc.setRangeValues(sheetId, startCoord, values, { label: t("clipboard.paste") });
  return true;
}
