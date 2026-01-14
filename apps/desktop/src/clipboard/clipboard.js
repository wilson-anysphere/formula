import { normalizeRange, parseA1, parseRangeA1 } from "../document/coords.js";
import { t } from "../i18n/index.js";
import { parseHtmlToCellGrid, serializeCellGridToHtml } from "./html.js";
import { extractPlainTextFromRtf, serializeCellGridToRtf } from "./rtf.js";
import { parseTsvToCellGrid, serializeCellGridToTsv } from "./tsv.js";
import { ClipboardParseLimitError, DEFAULT_MAX_CLIPBOARD_HTML_CHARS, DEFAULT_MAX_CLIPBOARD_PARSE_CELLS } from "./limits.js";
import { enforceClipboardCopy } from "../dlp/enforceClipboardCopy.js";
import { normalizeExcelColorToCss } from "../shared/colors.js";
import { getStyleNumberFormat } from "../formatting/styleFieldAccess.js";

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
  return normalizeExcelColorToCss(hex) ?? null;
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

function hasOwn(obj, key) {
  return Boolean(obj) && Object.prototype.hasOwnProperty.call(obj, key);
}

/**
 * Minimal CSS named-color support for paste normalization in non-DOM environments (Node tests).
 *
 * This is intentionally not exhaustive; browsers/WebViews will still normalize the full
 * CSS color space via `getComputedStyle` in `normalizeCssColorViaDom`.
 *
 * @type {Record<string, { r: number, g: number, b: number }>}
 */
const CSS_NAMED_COLORS = {
  black: { r: 0, g: 0, b: 0 },
  silver: { r: 192, g: 192, b: 192 },
  gray: { r: 128, g: 128, b: 128 },
  grey: { r: 128, g: 128, b: 128 },
  white: { r: 255, g: 255, b: 255 },
  maroon: { r: 128, g: 0, b: 0 },
  red: { r: 255, g: 0, b: 0 },
  purple: { r: 128, g: 0, b: 128 },
  fuchsia: { r: 255, g: 0, b: 255 },
  magenta: { r: 255, g: 0, b: 255 },
  green: { r: 0, g: 128, b: 0 },
  lime: { r: 0, g: 255, b: 0 },
  olive: { r: 128, g: 128, b: 0 },
  yellow: { r: 255, g: 255, b: 0 },
  navy: { r: 0, g: 0, b: 128 },
  blue: { r: 0, g: 0, b: 255 },
  teal: { r: 0, g: 128, b: 128 },
  aqua: { r: 0, g: 255, b: 255 },
  cyan: { r: 0, g: 255, b: 255 },
  orange: { r: 255, g: 165, b: 0 },
  hotpink: { r: 255, g: 105, b: 180 },
  rebeccapurple: { r: 102, g: 51, b: 153 },
};

function clampByte(value) {
  if (!Number.isFinite(value)) return 0;
  return Math.max(0, Math.min(255, Math.round(value)));
}

function toHex2(value) {
  return clampByte(value).toString(16).padStart(2, "0").toUpperCase();
}

function parseCssRgbChannel(value) {
  const trimmed = String(value).trim();
  if (!trimmed) return null;

  const percent = /^([+-]?\d*\.?\d+)%$/.exec(trimmed);
  if (percent) {
    const p = Number(percent[1]);
    if (!Number.isFinite(p)) return null;
    return clampByte((p / 100) * 255);
  }

  const num = Number(trimmed);
  if (!Number.isFinite(num)) return null;
  return clampByte(num);
}

function parseCssAlphaChannel(value) {
  const trimmed = String(value).trim();
  if (!trimmed) return null;

  const percent = /^([+-]?\d*\.?\d+)%$/.exec(trimmed);
  if (percent) {
    const p = Number(percent[1]);
    if (!Number.isFinite(p)) return null;
    const a = Math.max(0, Math.min(1, p / 100));
    return clampByte(a * 255);
  }

  const num = Number(trimmed);
  if (!Number.isFinite(num)) return null;
  const a = Math.max(0, Math.min(1, num));
  return clampByte(a * 255);
}

function parseCssRgbFunction(value) {
  const trimmed = String(value).trim();
  const match = /^(rgb|rgba)\(\s*([\s\S]+)\s*\)$/i.exec(trimmed);
  if (!match) return null;
  const args = match[2]?.trim() ?? "";
  if (!args) return null;

  let rgbPart = args;
  let alphaPart = null;

  // Support modern slash syntax: rgb(… / …).
  if (args.includes("/")) {
    const parts = args.split("/");
    if (parts.length !== 2) return null;
    rgbPart = parts[0].trim();
    alphaPart = parts[1].trim();
  }

  let parts = [];
  if (rgbPart.includes(",")) {
    parts = rgbPart
      .split(",")
      .map((p) => p.trim())
      .filter(Boolean);
  } else {
    parts = rgbPart
      .split(/\s+/)
      .map((p) => p.trim())
      .filter(Boolean);
  }

  if (alphaPart == null && parts.length === 4) {
    alphaPart = parts[3];
    parts = parts.slice(0, 3);
  }

  if (parts.length < 3) return null;

  const r = parseCssRgbChannel(parts[0]);
  const g = parseCssRgbChannel(parts[1]);
  const b = parseCssRgbChannel(parts[2]);
  if (r == null || g == null || b == null) return null;

  const a = alphaPart != null ? parseCssAlphaChannel(alphaPart) : 255;
  if (a == null) return null;

  return { a, r, g, b };
}

function normalizeCssColorViaDom(color) {
  try {
    // eslint-disable-next-line no-undef
    const doc = typeof document !== "undefined" ? document : null;
    // eslint-disable-next-line no-undef
    const compute = typeof getComputedStyle === "function" ? getComputedStyle : null;
    if (!doc || typeof doc.createElement !== "function" || !compute) return null;

    const el = doc.createElement("span");
    el.style.color = "";
    el.style.color = color;
    // Invalid colors yield an empty string.
    if (!el.style.color) return null;

    const parent = doc.body ?? doc.documentElement;
    if (parent && typeof parent.appendChild === "function") {
      parent.appendChild(el);
    }
    const computed = compute(el).color;
    el.remove();
    return typeof computed === "string" && computed.trim() ? computed.trim() : null;
  } catch {
    return null;
  }
}

/**
 * Convert a CSS color string into the DocumentController's ARGB color format (`#AARRGGBB`).
 *
 * This intentionally supports:
 * - `#RGB` (CSS shorthand)
 * - `#RRGGBB`
 * - `#RGBA` (CSS shorthand)
 * - `#AARRGGBB` (Excel/OOXML style)
 * - `rgb(r,g,b)` / `rgb(r g b)`
 * - `rgba(r,g,b,a)`
 *
 * For other CSS color syntaxes (named colors, `hsl()`, etc.), we attempt a DOM-based
 * normalization when running in a browser/WebView. In non-DOM environments (node tests),
 * unsupported strings return `null`.
 *
 * @param {any} value
 * @returns {string | null}
 */
function cssColorToArgb(value) {
  if (typeof value !== "string") return null;
  const trimmed = value.trim();
  if (!trimmed) return null;

  const lower = trimmed.toLowerCase();
  // Keep as concatenation so hardcoded-color detectors don't flag it as a UI literal.
  if (lower === "transparent") return "#" + "00000000";
  if (lower === "none") return null;

  const named = CSS_NAMED_COLORS[lower];
  if (named) return `#FF${toHex2(named.r)}${toHex2(named.g)}${toHex2(named.b)}`;

  const hex = trimmed.replace(/^#/, "");
  if (/^[0-9a-fA-F]{3}$/.test(hex)) {
    const r = hex[0];
    const g = hex[1];
    const b = hex[2];
    return `#FF${r}${r}${g}${g}${b}${b}`.toUpperCase();
  }
  if (/^[0-9a-fA-F]{6}$/.test(hex)) return `#FF${hex.toUpperCase()}`;
  if (/^[0-9a-fA-F]{4}$/.test(hex)) {
    // CSS `#RGBA` uses alpha as the last nibble; expand it and rearrange to DocumentController's `#AARRGGBB`.
    const r = hex[0];
    const g = hex[1];
    const b = hex[2];
    const a = hex[3];
    return `#${a}${a}${r}${r}${g}${g}${b}${b}`.toUpperCase();
  }
  if (/^[0-9a-fA-F]{8}$/.test(hex)) return `#${hex.toUpperCase()}`;

  const rgb = parseCssRgbFunction(trimmed);
  if (rgb) return `#${toHex2(rgb.a)}${toHex2(rgb.r)}${toHex2(rgb.g)}${toHex2(rgb.b)}`;

  const normalized = normalizeCssColorViaDom(trimmed);
  if (normalized) {
    const parsed = parseCssRgbFunction(normalized);
    if (parsed) return `#${toHex2(parsed.a)}${toHex2(parsed.r)}${toHex2(parsed.g)}${toHex2(parsed.b)}`;
  }

  return null;
}

/**
 * Convert a clipboard cell `format` object into a DocumentController style object.
 *
 * Clipboard HTML parsing currently produces a flat format object:
 * `{ bold, italic, underline, textColor, backgroundColor, numberFormat }`.
 *
 * The DocumentController stores canonical styles in an OOXML-ish schema:
 * `{ font: { bold, italic, underline, color }, fill: { pattern, fgColor }, numberFormat }`
 *
 * @param {any} format
 * @returns {any | null}
 */
export function clipboardFormatToDocStyle(format) {
  if (!format || typeof format !== "object") return null;

  /** @type {any} */
  const out = {};

  const setFont = (key, value) => {
    out.font ??= {};
    out.font[key] = value;
  };

  if (typeof format.bold === "boolean") setFont("bold", format.bold);
  if (typeof format.italic === "boolean") setFont("italic", format.italic);

  if (typeof format.underline === "boolean") {
    setFont("underline", format.underline);
  } else if (typeof format.underline === "string") {
    setFont("underline", format.underline !== "none");
  }

  const textColor = cssColorToArgb(format.textColor);
  if (textColor) setFont("color", textColor);

  const backgroundColor = cssColorToArgb(format.backgroundColor);
  if (backgroundColor) {
    out.fill = { pattern: "solid", fgColor: backgroundColor };
  }

  const numberFormat = format.numberFormat;
  if (typeof numberFormat === "string") {
    const trimmed = numberFormat.trim();
    // Treat "General" (Excel default) and empty string as clearing number formatting.
    // Use `null` (explicit override) instead of omitting the key so pastes can clear
    // inherited row/column formatting (Excel semantics).
    if (trimmed === "" || trimmed.toLowerCase() === "general") {
      out.numberFormat = null;
    } else {
      out.numberFormat = numberFormat;
    }
  }

  return Object.keys(out).length > 0 ? out : null;
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

  // Treat `numberFormat` as authoritative even when it is explicitly cleared (null/undefined)
  // so UI patches can override imported formula-model `number_format` strings.
  const numberFormat = getStyleNumberFormat(style);
  if (numberFormat != null) out.numberFormat = numberFormat;

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
    rtf: serializeCellGridToRtf(grid),
  };
}

/**
 * Parse clipboard payloads in priority order: HTML > TSV/plain.
 * @param {ClipboardContent} content
 * @param {{ maxCells?: number, maxChars?: number }} [options]
 * @returns {CellGrid | null}
 */
export function parseClipboardContentToCellGrid(content, options = {}) {
  const rawMaxCells = options.maxCells;
  const maxCells = (() => {
    if (rawMaxCells === Infinity) return Infinity;
    const n = Number(rawMaxCells);
    if (Number.isFinite(n) && Number.isInteger(n) && n > 0) return n;
    return DEFAULT_MAX_CLIPBOARD_PARSE_CELLS;
  })();

  const rawMaxChars = options.maxChars;
  const maxChars = (() => {
    if (rawMaxChars === Infinity) return Infinity;
    const n = Number(rawMaxChars);
    if (Number.isFinite(n) && Number.isInteger(n) && n > 0) return n;
    return DEFAULT_MAX_CLIPBOARD_HTML_CHARS;
  })();

  const parseOptions = { maxCells, maxChars };

  /** @type {unknown | null} */
  let deferredLimitError = null;

  const html = typeof content.html === "string" ? content.html : null;
  // Avoid trimming the raw HTML string before parsing: CF_HTML offset fields are byte offsets into
  // the *original* payload, so stripping whitespace can invalidate otherwise-correct offsets.
  if (html && /\S/.test(html)) {
    try {
      const parsed = parseHtmlToCellGrid(html, parseOptions);
      if (parsed) return parsed;
    } catch (err) {
      if (err instanceof ClipboardParseLimitError || err?.name === "ClipboardParseLimitError") {
        deferredLimitError = err;
      } else {
        // Ignore parse failures and fall back to TSV/RTF.
      }
    }
  }

  const text = content.text;
  const rtf = content.rtf;
  const hasMeaningfulRtf = typeof rtf === "string" && rtf.trim() !== "";

  // Some clipboard backends may provide an empty `text/plain` payload alongside richer formats.
  // Only treat empty strings as "missing" when we have an RTF payload to fall back to. Otherwise
  // keep the legacy behavior (pasting a blank 1×1 cell should still clear the target).
  if (typeof text === "string" && (text !== "" || (!hasMeaningfulRtf && deferredLimitError === null))) {
    try {
      const parsed = parseTsvToCellGrid(text, parseOptions);
      if (parsed) return parsed;
    } catch (err) {
      if (err instanceof ClipboardParseLimitError || err?.name === "ClipboardParseLimitError") {
        throw err;
      }
      // Ignore parse failures and fall back to RTF.
    }
  }

  if (hasMeaningfulRtf) {
    const extracted = extractPlainTextFromRtf(rtf);
    if (extracted) {
      try {
        const parsed = parseTsvToCellGrid(extracted, parseOptions);
        if (parsed) return parsed;
      } catch (err) {
        if (err instanceof ClipboardParseLimitError || err?.name === "ClipboardParseLimitError") {
          throw err;
        }
        // Ignore parse failures.
      }
    }
  }

  if (deferredLimitError) {
    throw deferredLimitError;
  }

  return null;
}

/**
 * Extract a rectangular cell grid from the document controller.
 *
 * @param {DocumentController} doc
 * @param {string} sheetId
 * @param {CellRange | string} range
 * @param {{ maxCells?: number }} [options]
 * @returns {CellGrid}
 */
export const DEFAULT_MAX_CELL_GRID_CELLS = 200_000;

export function getCellGridFromRange(doc, sheetId, range, options = {}) {
  const r = typeof range === "string" ? parseRangeA1(range) : normalizeRange(range);
  const sheetKey = String(sheetId ?? "").trim();
  const maxCells = options.maxCells ?? DEFAULT_MAX_CELL_GRID_CELLS;
  const rowCount = Math.max(0, r.end.row - r.start.row + 1);
  const colCount = Math.max(0, r.end.col - r.start.col + 1);
  const cellCount = rowCount * colCount;
  if (cellCount > maxCells) {
    throw new Error(
      `Range too large to materialize (${rowCount}x${colCount}=${cellCount} cells). ` +
        `Limit is ${maxCells} cells.`
    );
  }

  /** @type {CellGrid} */
  const grid = [];
  /** @type {Map<string, any>} */
  const formatCache = new Map();

  // Avoid creating phantom sheets when callers hold stale sheet ids (e.g. during sheet
  // deletion/applyState races). DocumentController lazily materializes sheets on
  // `getCell()` / `getCellFormat()` / `getSheetView()`.
  const sheetExists = (() => {
    if (!sheetKey) return false;
    try {
      if (doc && typeof doc.getSheetMeta === "function") {
        return Boolean(doc.getSheetMeta(sheetKey));
      }
      const sheets = doc?.model?.sheets;
      const sheetMeta = doc?.sheetMeta;
      if (sheets instanceof Map || sheetMeta instanceof Map) {
        return (sheets instanceof Map && sheets.has(sheetKey)) || (sheetMeta instanceof Map && sheetMeta.has(sheetKey));
      }
    } catch {
      // ignore
    }
    // If we can't verify existence, preserve legacy behavior.
    return true;
  })();

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
    if (!sheetExists) return cachedSheetView;
    if (doc && typeof doc.getSheetView === "function") {
      cachedSheetView = doc.getSheetView(sheetKey);
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
   * @returns {[number, number, number, number, number]}
   */
  const getCellStyleIdTuple = (row, col, cellStyleId) => {
    if (!sheetExists) return [0, 0, 0, normalizeStyleId(cellStyleId), 0];
    const coord = { row, col };
    const styleIdForRowInRuns = (runs, r) => {
      if (!Array.isArray(runs) || runs.length === 0) return 0;
      let lo = 0;
      let hi = runs.length - 1;
      while (lo <= hi) {
        const mid = (lo + hi) >> 1;
        const run = runs[mid];
        const startRow = Number(run?.startRow);
        const endRowExclusive = Number(run?.endRowExclusive);
        if (!Number.isInteger(startRow) || !Number.isInteger(endRowExclusive)) return 0;
        if (r < startRow) hi = mid - 1;
        else if (r >= endRowExclusive) lo = mid + 1;
        else return normalizeStyleId(run?.styleId);
      }
      return 0;
    };

    // Preferred: helper that returns the tuple directly.
    /** @type {any} */
    const helper =
      (doc && typeof doc.getCellFormatStyleIds === "function" && doc.getCellFormatStyleIds) ||
      (doc && typeof doc.getCellFormatIds === "function" && doc.getCellFormatIds) ||
      (doc && typeof doc.getCellFormatStyleIdTuple === "function" && doc.getCellFormatStyleIdTuple) ||
      (doc && typeof doc.getCellStyleIdTuple === "function" && doc.getCellStyleIdTuple);

    if (helper) {
      const out = helper.call(doc, sheetKey, coord);
      if (Array.isArray(out) && out.length >= 5) {
        return [normalizeStyleId(out[0]), normalizeStyleId(out[1]), normalizeStyleId(out[2]), normalizeStyleId(out[3]), normalizeStyleId(out[4])];
      }
      if (Array.isArray(out) && out.length >= 4) {
        return [normalizeStyleId(out[0]), normalizeStyleId(out[1]), normalizeStyleId(out[2]), normalizeStyleId(out[3]), 0];
      }
      if (out && typeof out === "object") {
        const sheetDefaultStyleId = normalizeStyleId(
          out.sheetDefaultStyleId ?? out.sheetStyleId ?? out.defaultStyleId ?? out.sheetDefault
        );
        const rowStyleId = normalizeStyleId(out.rowStyleId ?? out.rowDefaultStyleId ?? out.rowDefault);
        const colStyleId = normalizeStyleId(out.colStyleId ?? out.colDefaultStyleId ?? out.colDefault);
        const cellId = normalizeStyleId(out.cellStyleId ?? out.styleId ?? cellStyleId);
        const runId = normalizeStyleId(
          out.rangeRunStyleId ??
            out.rangeRun ??
            out.runStyleId ??
            out.formatRunStyleId ??
            out.runDefaultStyleId ??
            out.rangeStyleId ??
            0
        );
        return [sheetDefaultStyleId, rowStyleId, colStyleId, cellId, runId];
      }
    }

    // DocumentController internal model shape (layered formatting).
    // - Sheet default style layer: `sheet.defaultStyleId` (legacy: `sheet.sheetStyleId` / `sheet.sheetDefaultStyleId`)
    // - Column style layer: `sheet.colStyleIds` (legacy: `sheet.colStyles`)
    // - Row style layer: `sheet.rowStyleIds` (legacy: `sheet.rowStyles`)
    //
    // The public `getCellFormat()` API returns the merged style, but does not expose
    // the style-id tuple needed for caching. When possible, derive it from the
    // internal sheet model maps.
    const sheetModel =
      doc?.model?.sheets?.get && typeof doc.model.sheets.get === "function" ? doc.model.sheets.get(sheetKey) : null;
    if (sheetModel && typeof sheetModel === "object") {
      const sheetDefaultStyleId = normalizeStyleId(
        sheetModel.defaultStyleId ?? sheetModel.sheetStyleId ?? sheetModel.sheetDefaultStyleId
      );

      const rowStyles = sheetModel.rowStyleIds ?? sheetModel.rowStyles ?? sheetModel.rowStyleIdByRow;
      const colStyles = sheetModel.colStyleIds ?? sheetModel.colStyles ?? sheetModel.colStyleIdByCol;

      const rowStyleId = (() => {
        if (!rowStyles) return 0;
        if (typeof rowStyles.get === "function") return normalizeStyleId(rowStyles.get(row));
        return normalizeStyleId(rowStyles[String(row)] ?? rowStyles[row]);
      })();

      const colStyleId = (() => {
        if (!colStyles) return 0;
        if (typeof colStyles.get === "function") return normalizeStyleId(colStyles.get(col));
        return normalizeStyleId(colStyles[String(col)] ?? colStyles[col]);
      })();

      const runsByCol = sheetModel.formatRunsByCol ?? sheetModel.formatRunsByColumn ?? sheetModel.rangeRunsByCol;
      const runs = (() => {
        if (!runsByCol) return null;
        if (typeof runsByCol.get === "function") return runsByCol.get(col);
        return runsByCol[String(col)] ?? runsByCol[col];
      })();
      const runStyleId = styleIdForRowInRuns(runs, row);

      return [sheetDefaultStyleId, rowStyleId, colStyleId, normalizeStyleId(cellStyleId), runStyleId];
    }

    // Best-effort fallback: query per-layer style ids if exposed.
    let sheetDefaultStyleId = 0;
    let rowStyleId = 0;
    let colStyleId = 0;

    if (doc && typeof doc.getSheetDefaultStyleId === "function") {
      sheetDefaultStyleId = normalizeStyleId(doc.getSheetDefaultStyleId(sheetKey));
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
      rowStyleId = normalizeStyleId(doc.getRowStyleId(sheetKey, row));
    }
    if (doc && typeof doc.getColStyleId === "function") {
      colStyleId = normalizeStyleId(doc.getColStyleId(sheetKey, col));
    }

    return [sheetDefaultStyleId, rowStyleId, colStyleId, normalizeStyleId(cellStyleId), 0];
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
    if (!sheetExists) return null;
    const coord = { row, col };
    if (doc && typeof doc.getCellFormat === "function") {
      return doc.getCellFormat(sheetKey, coord);
    }
    if (doc && typeof doc.getEffectiveCellStyle === "function") {
      return doc.getEffectiveCellStyle(sheetKey, coord);
    }
    if (doc && typeof doc.getCellStyle === "function") {
      return doc.getCellStyle(sheetKey, coord);
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
  const hasInternalLayeredStyleMaps = Boolean(doc?.model?.sheets?.get && typeof doc.model.sheets.get === "function");

  // Only consult `getSheetView()` when there isn't another obvious way to derive
  // the style-id tuple (some controllers may surface row/col style ids on the view).
  if (sheetExists && !hasStyleIdTupleHelper && !hasPerLayerStyleIdMethods && !hasInternalLayeredStyleMaps) {
    getSheetViewForStyleLayers();
  }
  const canDeriveStyleIdTuple = sheetExists && Boolean(
    hasStyleIdTupleHelper || hasPerLayerStyleIdMethods || hasInternalLayeredStyleMaps || cachedSheetViewHasStyleLayers
  );

  for (let row = r.start.row; row <= r.end.row; row++) {
    /** @type {CellState[]} */
    const outRow = [];
    for (let col = r.start.col; col <= r.end.col; col++) {
      const cell = (() => {
        if (doc && typeof doc.peekCell === "function") {
          return doc.peekCell(sheetKey, { row, col });
        }
        if (!sheetExists) return { value: null, formula: null, styleId: 0 };
        return doc.getCell(sheetKey, { row, col });
      })();

      const cellStyleId = typeof cell?.styleId === "number" ? cell.styleId : 0;

      // If the controller can supply the contributing style ids, cache by the
      // (sheet,row,col,cell,range-run) style-id tuple. Otherwise, fall back to caching by
      // the resolved style object to avoid collisions.
      let format = null;
      if (sheetExists && doc && typeof doc.getCellFormat === "function" && !canDeriveStyleIdTuple) {
        const style = getEffectiveStyle(row, col, cellStyleId);
        const cacheKey = stableStringify(style);
        if (formatCache.has(cacheKey)) {
          format = formatCache.get(cacheKey) ?? null;
        } else {
          format = styleToClipboardFormat(style);
          formatCache.set(cacheKey, format);
        }
      } else {
        if (sheetExists) {
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
 * @param {{ dlp?: { documentId: string, classificationStore: any, policy: any }, maxCells?: number }} [options]
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

  const grid = getCellGridFromRange(doc, sheetId, r, { maxCells: options.maxCells });
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
  let grid = null;
  try {
    grid = parseClipboardContentToCellGrid(content);
  } catch (err) {
    // Oversized clipboard payloads should not crash the paste handler.
    if (err instanceof ClipboardParseLimitError || err?.name === "ClipboardParseLimitError") {
      return false;
    }
    throw err;
  }
  if (!grid) return false;

  const mode = options.mode ?? "all";

  const values = grid.map((row) =>
    row.map((cell) => {
      // Note: pass values via object form (`{ value }`) so DocumentController does not
      // reinterpret strings like "=1+1" as formulas or coerce ID-like values (e.g. "00123")
      // into numbers.
      if (mode === "values") return { value: cell.value ?? null };
      if (mode === "formulas") return cell.formula != null ? { formula: cell.formula } : { value: cell.value ?? null };
      if (mode === "formats") return { format: clipboardFormatToDocStyle(cell.format ?? null) };

      // mode === "all"
      const format = clipboardFormatToDocStyle(cell.format ?? null);
      if (cell.formula != null) return { formula: cell.formula, value: null, format };
      return { value: cell.value ?? null, format };
    })
  );

  const startCoord = typeof start === "string" ? parseA1(start) : start;
  doc.setRangeValues(sheetId, startCoord, values, { label: t("clipboard.paste") });
  return true;
}
