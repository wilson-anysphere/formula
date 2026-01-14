import { colToIndex, parseA1Range, splitSheetQualifier } from "./a1.js";
import { normalizeName } from "./workbook.js";

function isLikelyA1(ref) {
  return /^(\$?[A-Za-z]{1,3}\$?\d+)(:\$?[A-Za-z]{1,3}\$?\d+)?$/.test(String(ref).trim());
}

const DEFAULT_MAX_ROWS = 1_048_576;
const DEFAULT_MAX_COLS = 16_384;

function parseA1RowOrColRange(ref) {
  const s = String(ref).trim();

  // Excel-style column range: A:A, A:C, $A:$C, etc.
  const colMatch = /^(\$?[A-Za-z]{1,3})\s*:\s*(\$?[A-Za-z]{1,3})$/.exec(s);
  if (colMatch) {
    const a = colToIndex(colMatch[1].replaceAll("$", ""));
    const b = colToIndex(colMatch[2].replaceAll("$", ""));
    const startCol = Math.min(a, b);
    const endCol = Math.max(a, b);
    return { startRow: 0, endRow: DEFAULT_MAX_ROWS - 1, startCol, endCol };
  }

  // Excel-style row range: 1:1, 1:10, $1:$10, etc.
  const rowMatch = /^(\$?\d+)\s*:\s*(\$?\d+)$/.exec(s);
  if (rowMatch) {
    const parseRow = (token) => {
      const raw = Number.parseInt(String(token).replaceAll("$", ""), 10);
      if (!Number.isFinite(raw) || raw < 1) return null;
      return raw - 1;
    };
    const a = parseRow(rowMatch[1]);
    const b = parseRow(rowMatch[2]);
    if (a == null || b == null) return null;
    const startRow = Math.min(a, b);
    const endRow = Math.max(a, b);
    return { startRow, endRow, startCol: 0, endCol: DEFAULT_MAX_COLS - 1 };
  }

  return null;
}

function parseStructuredRef(input) {
  const s = String(input).trim();

  const firstBracket = s.indexOf("[");
  if (firstBracket <= 0) return null;

  const tableName = s.slice(0, firstBracket);
  if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(tableName)) return null;

  const suffix = s.slice(firstBracket);
  if (!suffix) return null;

  const KNOWN_SELECTOR_ITEMS = new Set(["#ALL", "#HEADERS", "#DATA", "#TOTALS", "#THIS ROW"]);
  const normalizeSelector = (value) => String(value ?? "").trim().replace(/\s+/g, " ").toUpperCase();
  const isSelectorItem = (value) => KNOWN_SELECTOR_ITEMS.has(normalizeSelector(value));

  function findMatchingStructuredRefBracketEnd(src, start) {
    // Structured references escape closing brackets inside items by doubling: `]]` -> literal `]`.
    // That makes naive depth counting incorrect (it will pop twice for an escaped bracket).
    //
    // Use a small backtracking parser:
    // - On `[` increase depth.
    // - On `]]`, prefer treating it as an escape (consume both, depth unchanged), but remember
    //   a choice point. If we later fail to close all brackets, backtrack and reinterpret that
    //   `]]` as a real closing bracket.
    if (src[start] !== "[") return null;

    let i = start;
    let depth = 0;
    const escapeChoices = [];

    const backtrack = () => {
      const choice = escapeChoices.pop();
      if (!choice) return false;
      i = choice.i;
      depth = choice.depth;
      // Reinterpret the first `]` of the `]]` pair as a real closing bracket.
      depth -= 1;
      i += 1;
      return true;
    };

    while (true) {
      if (i >= src.length) {
        if (!backtrack()) return null;
        if (depth === 0) return i;
        continue;
      }

      const ch = src[i];
      if (ch === "[") {
        depth += 1;
        i += 1;
        continue;
      }

      if (ch === "]") {
        if (src[i + 1] === "]" && depth > 0) {
          escapeChoices.push({ i, depth });
          i += 2;
          continue;
        }

        depth -= 1;
        i += 1;
        if (depth === 0) return i;
        if (depth < 0) {
          if (!backtrack()) return null;
          if (depth === 0) return i;
        }
        continue;
      }

      i += 1;
    }
  }

  function parseNested(nested) {
    if (!nested.startsWith("[[")) return null;
    const items = [];
    const seps = [];
    let pos = 1; // start at the second '['
    while (true) {
      if (nested[pos] !== "[") return null;
      const end = findMatchingStructuredRefBracketEnd(nested, pos);
      if (!end) return null;
      const raw = nested.slice(pos + 1, end - 1);
      // Excel escapes `]` inside structured reference items by doubling it: `]]`.
      items.push(raw.replaceAll("]]", "]").trim());
      pos = end;
      while (pos < nested.length && /\s/.test(nested[pos])) pos += 1;
      const ch = nested[pos];
      if (ch === "," || ch === ":") {
        seps.push(ch);
        pos += 1;
        while (pos < nested.length && /\s/.test(nested[pos])) pos += 1;
        continue;
      }
      if (ch === "]") {
        pos += 1;
        while (pos < nested.length && /\s/.test(nested[pos])) pos += 1;
        if (pos !== nested.length) return null;
        break;
      }
      return null;
    }

    if (items.length === 0) return null;

    let selectorCount = 0;
    while (selectorCount < items.length && isSelectorItem(items[selectorCount])) {
      if (selectorCount < seps.length && seps[selectorCount] === ":") return null;
      selectorCount += 1;
    }

    if (selectorCount > 0 && selectorCount < items.length) {
      if (seps[selectorCount - 1] !== ",") return null;
    }

    return {
      selectors: items.slice(0, selectorCount),
      columnItems: items.slice(selectorCount),
      columnSeparators: seps.slice(selectorCount),
    };
  }

  const nested = parseNested(suffix);
  if (nested) {
    return {
      tableName,
      selectors: nested.selectors,
      columnItems: nested.columnItems,
      columnSeparators: nested.columnSeparators,
    };
  }

  if (suffix.startsWith("[[")) return null;

  const simpleMatch = suffix.match(/^\[\s*((?:[^\]]|]])+)\s*\]$/);
  if (simpleMatch) {
    const item = simpleMatch[1].replaceAll("]]", "]").trim();
    if (isSelectorItem(item)) {
      return { tableName, selectors: [item], columnItems: [], columnSeparators: [] };
    }
    return { tableName, selectors: [], columnItems: [item], columnSeparators: [] };
  }

  return null;
}

function resolveSheetName(name, workbook) {
  const sheetName = String(name ?? "").trim();
  if (!sheetName) return sheetName;
  if (!workbook || typeof workbook.getSheet !== "function") return sheetName;

  const sheet = workbook.getSheet(sheetName);
  if (!sheet) throw new Error(`Unknown sheet: ${sheetName}`);
  const canonical = typeof sheet.name === "string" ? String(sheet.name).trim() : "";
  return canonical || sheetName;
}

/**
 * Parse Go To / Name box input.
 *
 * Supports:
 * - A1 references: `A1`, `A1:B2`
 * - Full column/row ranges: `A:A`, `1:1`
 * - Sheet-qualified: `Sheet2!C3`, `'My Sheet'!A1`
 * - Named ranges: `MyName`
 * - Table structured refs:
 *   - `Table1[Column]`
 *   - `Table1[#All]` / `Table1[#Headers]` / `Table1[#Data]` / `Table1[#Totals]`
 *   - `Table1[[#All],[Column]]` (and other selectors like `#Data`/`#Headers`/`#Totals`)
 *   - `Table1[[#Headers],[#Data],[Column]]` (multi-item selector unions; only when the resulting range is rectangular)
 *   - `Table1[[#All],[#Totals]]` (multi-item selector unions without an explicit column list; rectangular only)
 *   - `Table1[[#All],[Col1],[Col2]]` (multi-column; contiguous columns only)
 *   - `Table1[[#All],[Col1]:[Col3]]` (column-range)
 */
export function parseGoTo(input, { workbook, currentSheetName } = {}) {
  if (!workbook) throw new Error("parseGoTo: workbook is required");
  if (!currentSheetName) throw new Error("parseGoTo: currentSheetName is required");

  const raw = String(input).trim();
  if (raw === "") throw new Error("parseGoTo: empty input");

  const { sheetName: qualifiedSheet, ref } = splitSheetQualifier(raw);
  const sheetName = resolveSheetName(qualifiedSheet ?? currentSheetName, workbook);

  // Structured reference
  const structured = parseStructuredRef(ref);
  if (structured) {
    const table = workbook.getTable(structured.tableName);
    if (!table) throw new Error(`Unknown table: ${structured.tableName}`);
    const tableSheetName = resolveSheetName(table.sheetName, workbook);

    const normalizeSelector = (value) => String(value ?? "").trim().replace(/\s+/g, " ").toUpperCase();
    const selectors =
      Array.isArray(structured.selectors) && structured.selectors.length > 0 ? structured.selectors.map(normalizeSelector) : ["#ALL"];

    const unsupportedSelector = selectors.find((s) => s !== "#ALL" && s !== "#HEADERS" && s !== "#DATA" && s !== "#TOTALS");
    if (unsupportedSelector) {
      // Selectors like `#This Row` are relative to an active row context, which `parseGoTo` does not have.
      // Reject them explicitly rather than returning a misleading full-table range.
      throw new Error(`Unsupported structured reference selector: ${unsupportedSelector}`);
    }

    // Resolve row union. `parseGoTo` only returns a single rectangle; reject discontiguous unions.
    const rowIntervals = selectors.map((s) => {
      if (s === "#HEADERS") return { startRow: table.startRow, endRow: table.startRow };
      if (s === "#TOTALS") return { startRow: table.endRow, endRow: table.endRow };
      if (s === "#DATA") {
        const startRow = table.endRow > table.startRow ? table.startRow + 1 : table.startRow;
        return { startRow, endRow: table.endRow };
      }
      // #ALL
      return { startRow: table.startRow, endRow: table.endRow };
    });

    rowIntervals.sort((a, b) => a.startRow - b.startRow);
    const mergedRows = [];
    for (const interval of rowIntervals) {
      const last = mergedRows[mergedRows.length - 1];
      if (!last) {
        mergedRows.push({ ...interval });
        continue;
      }
      if (interval.startRow <= last.endRow + 1) {
        last.endRow = Math.max(last.endRow, interval.endRow);
      } else {
        mergedRows.push({ ...interval });
      }
    }
    if (mergedRows.length !== 1) {
      throw new Error(`Discontiguous structured reference selectors are not supported: ${structured.tableName}`);
    }
    const { startRow, endRow } = mergedRows[0];

    const columns = table.columns ?? [];
    const findColumnIndex = (name) => columns.findIndex((c) => normalizeName(c) === normalizeName(name));

    let startCol = table.startCol;
    let endCol = table.endCol;

    const columnItems = Array.isArray(structured.columnItems) ? structured.columnItems : [];
    const columnSeps = Array.isArray(structured.columnSeparators) ? structured.columnSeparators : [];
    if (columnItems.length > 0) {
      const indices = new Set();
      let i = 0;
      while (i < columnItems.length) {
        const sep = i < columnSeps.length ? columnSeps[i] : null;
        if (sep === ":") {
          if (i + 1 >= columnItems.length) throw new Error(`Invalid structured reference: ${input}`);
          const a = findColumnIndex(columnItems[i]);
          const b = findColumnIndex(columnItems[i + 1]);
          if (a === -1 || b === -1) throw new Error(`Unknown table column: ${structured.tableName}`);
          const lo = Math.min(a, b);
          const hi = Math.max(a, b);
          for (let idx = lo; idx <= hi; idx += 1) indices.add(idx);
          i += 2;
          if (i < columnItems.length && columnSeps[i - 1] !== ",") throw new Error(`Invalid structured reference: ${input}`);
          continue;
        }

        const idx = findColumnIndex(columnItems[i]);
        if (idx === -1) throw new Error(`Unknown table column: ${structured.tableName}`);
        indices.add(idx);
        i += 1;
        if (i < columnItems.length && columnSeps[i - 1] !== ",") throw new Error(`Invalid structured reference: ${input}`);
      }

      const uniq = Array.from(indices).sort((a, b) => a - b);
      if (uniq.length === 0) throw new Error(`Invalid structured reference: ${input}`);
      const minIdx = uniq[0];
      const maxIdx = uniq[uniq.length - 1];

      // Multi-column structured refs can represent unions. Only resolve contiguous columns to avoid
      // selecting a misleading bounding rectangle.
      if (maxIdx - minIdx + 1 !== uniq.length) {
        throw new Error(`Non-contiguous structured reference columns are not supported: ${structured.tableName}`);
      }

      startCol = table.startCol + minIdx;
      endCol = table.startCol + maxIdx;
    }

    return { type: "range", source: "table", sheetName: tableSheetName, range: { startRow, endRow, startCol, endCol } };
  }

  // A1 reference
  if (isLikelyA1(ref)) {
    return { type: "range", source: "a1", sheetName, range: parseA1Range(ref) };
  }

  // Excel-style row/column references (A:A, 1:1).
  const rowOrCol = parseA1RowOrColRange(ref);
  if (rowOrCol) {
    return { type: "range", source: "a1", sheetName, range: rowOrCol };
  }

  // Named range
  const named = workbook.getName(ref);
  if (named) {
    const targetSheet = resolveSheetName(named.sheetName ?? sheetName, workbook);
    return {
      type: "range",
      source: "name",
      sheetName: targetSheet,
      range: named.range,
    };
  }

  throw new Error(`Unrecognized Go To reference: ${input}`);
}
