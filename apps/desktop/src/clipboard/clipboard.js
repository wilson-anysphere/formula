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
  /** @type {Map<number, any>} */
  const formatCache = new Map();

  for (let row = r.start.row; row <= r.end.row; row++) {
    /** @type {CellState[]} */
    const outRow = [];
    for (let col = r.start.col; col <= r.end.col; col++) {
      const cell = doc.getCell(sheetId, { row, col });
      const styleId = typeof cell?.styleId === "number" ? cell.styleId : 0;
      let format = null;
      if (styleId !== 0) {
        if (formatCache.has(styleId)) {
          format = formatCache.get(styleId) ?? null;
        } else {
          const style =
            doc?.styleTable?.get && typeof doc.styleTable.get === "function" ? doc.styleTable.get(styleId) : null;
          format = styleToClipboardFormat(style);
          formatCache.set(styleId, format);
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
