import { excelSerialToDate, parseScalar } from "../shared/valueParsing.js";
import { ClipboardParseLimitError, DEFAULT_MAX_CLIPBOARD_PARSE_CELLS } from "./limits.js";

/**
 * @typedef {import("./types.js").CellGrid} CellGrid
 * @typedef {import("../document/cell.js").CellState} CellState
 */

function isLikelyDateNumberFormat(fmt) {
  if (typeof fmt !== "string") return false;
  const lower = fmt.toLowerCase();
  return lower.includes("yyyy-mm-dd") || lower.includes("m/d/yyyy");
}

function isLikelyTimeNumberFormat(fmt) {
  if (typeof fmt !== "string") return false;
  const compact = fmt.toLowerCase().replace(/\s+/g, "");
  return /^h{1,2}:m{1,2}(:s{1,2})?$/.test(compact);
}

function pad2(value) {
  return String(value).padStart(2, "0");
}

function formatExcelTime(serial, fmt) {
  const date = excelSerialToDate(serial);
  const hh = date.getUTCHours();
  const mm = date.getUTCMinutes();
  const ss = date.getUTCSeconds();
  const compact = String(fmt).toLowerCase().replace(/\s+/g, "");
  const hasSeconds = compact.includes(":s");
  return hasSeconds ? `${pad2(hh)}:${pad2(mm)}:${pad2(ss)}` : `${pad2(hh)}:${pad2(mm)}`;
}

/**
 * @param {CellState} cell
 * @returns {string}
 */
function cellValueToPlainText(cell) {
  const value = cell.value;
  if (value == null) {
    const formula = cell.formula;
    if (typeof formula === "string" && formula.trim() !== "") {
      // When we don't have a cached/display value, fall back to copying the formula text
      // (keeps the payload useful for spreadsheet-to-spreadsheet pastes).
      return formula;
    }
    return "";
  }

  // DocumentController stores rich text as `{ text, runs }`. Clipboard payloads should
  // round-trip as plain text (like Excel/Sheets) rather than `[object Object]`.
  if (typeof value === "object" && typeof value.text === "string") {
    return value.text;
  }

  if (typeof value === "string" && (value.trimStart().startsWith("=") || value.startsWith("'"))) {
    // Escape TSV so we don't accidentally treat literal strings (e.g. "=literal") as formulas
    // when parsing the clipboard text back into a cell grid.
    return `'${value}`;
  }

  const numberFormat = cell.format?.numberFormat;
  if (typeof value === "number" && isLikelyDateNumberFormat(numberFormat)) {
    const date = excelSerialToDate(value);
    const lower = numberFormat.toLowerCase();
    return lower.includes("h") ? date.toISOString() : date.toISOString().slice(0, 10);
  }

  if (typeof value === "number" && isLikelyTimeNumberFormat(numberFormat)) {
    return formatExcelTime(value, numberFormat);
  }

  return String(value);
}

/**
 * @param {CellGrid} grid
 * @returns {string}
 */
export function serializeCellGridToTsv(grid) {
  return grid
    .map((row) => row.map(cellValueToPlainText).join("\t"))
    .join("\n");
}

/**
 * @param {string} tsv
 * @param {{ maxCells?: number }} [options]
 * @returns {CellGrid}
 */
export function parseTsvToCellGrid(tsv, options = {}) {
  const rawMaxCells = options.maxCells;
  const maxCells = (() => {
    if (rawMaxCells === Infinity) return Infinity;
    const n = Number(rawMaxCells);
    if (Number.isFinite(n) && Number.isInteger(n) && n > 0) return n;
    return DEFAULT_MAX_CLIPBOARD_PARSE_CELLS;
  })();

  const text = String(tsv ?? "");

  /** @type {CellGrid} */
  const grid = [];
  /** @type {CellState[]} */
  let row = [];
  let cellStart = 0;
  let cellCount = 0;

  const parseCell = (raw) => {
    if (raw.startsWith("'")) {
      // Excel convention: a leading apostrophe forces text.
      return { value: raw.slice(1), formula: null, format: null };
    }

    const trimmed = raw.trimStart();
    if (trimmed.startsWith("=")) {
      return { value: null, formula: trimmed, format: null };
    }

    const parsed = parseScalar(raw);
    const format = parsed.type === "datetime" ? { numberFormat: parsed.numberFormat } : null;
    return { value: parsed.value, formula: null, format };
  };

  const pushCell = (end) => {
    row.push(parseCell(text.slice(cellStart, end)));
    cellCount += 1;
    if (Number.isFinite(maxCells) && cellCount > maxCells) {
      throw new ClipboardParseLimitError(
        `Clipboard TSV too large to parse (${cellCount.toLocaleString()} cells; max=${maxCells.toLocaleString()}).`
      );
    }
  };

  const pushRow = () => {
    grid.push(row);
    row = [];
  };

  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    // tab, newline, carriage return
    if (code !== 9 && code !== 10 && code !== 13) continue;

    pushCell(i);

    if (code === 9) {
      // tab
      cellStart = i + 1;
      continue;
    }

    // newline / CRLF
    pushRow();

    if (code === 13 && text.charCodeAt(i + 1) === 10) {
      // CRLF -> skip the LF
      i += 1;
    }
    cellStart = i + 1;
  }

  // Add the final cell/row (if any). When the payload ends with a newline, `row` will be
  // empty here, mirroring the prior behavior that dropped the final empty record.
  if (cellStart <= text.length) {
    // Note: for an empty string, this produces a 1x1 empty cell, matching `"".split("\n")`.
    if (row.length > 0 || cellStart < text.length || grid.length === 0) {
      pushCell(text.length);
      pushRow();
    }
  }

  return grid;
}
