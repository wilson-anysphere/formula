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

  // Match the subset supported by the formula tokenizer / reference extractor:
  // - TableName[ColumnName]
  // - TableName[#All] / [#Headers] / [#Data] / [#Totals]
  // - TableName[[#All],[ColumnName]] (and other selectors like #Data/#Headers/#Totals)
  // - TableName[[#All],[Col1],[Col2]] (multi-column; contiguous columns only)
  // - TableName[[#All],[Col1]:[Col3]] (column-range)

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

  function parseNestedItems(nested) {
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
      items.push(raw.replaceAll("]]", "]"));
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

    let selector = null;
    let columnItems = items;
    let columnSeps = seps;

    const first = String(items[0] ?? "").trim();
    if (first.startsWith("#")) {
      if (items.length < 2) return null;
      if (seps[0] !== ",") return null;
      selector = first;
      columnItems = items.slice(1);
      columnSeps = seps.slice(1);
    }

    if (columnItems.length === 1) {
      return { selector, columnName: String(columnItems[0] ?? "").trim(), columns: null, columnMode: null };
    }

    const hasColon = columnSeps.includes(":");
    if (hasColon) {
      if (columnItems.length !== 2) return null;
      if (columnSeps.length !== 1 || columnSeps[0] !== ":") return null;
      return { selector, columnName: null, columns: columnItems.map((c) => String(c ?? "").trim()), columnMode: "range" };
    }

    if (!columnSeps.every((sep) => sep === ",")) return null;
    return { selector, columnName: null, columns: columnItems.map((c) => String(c ?? "").trim()), columnMode: "list" };
  }

  const nested = parseNestedItems(suffix);
  if (nested) {
    return {
      tableName,
      selector: nested.selector,
      columnName: nested.columnName,
      columns: nested.columns,
      columnMode: nested.columnMode,
    };
  }

  // Avoid mis-parsing nested bracket groups like `[[#All],[Amount]]` as a single item.
  // If the nested parser fails, treat the reference as invalid rather than falling back to
  // a simple `[Column]` match.
  if (suffix.startsWith("[[")) return null;

  const simpleMatch = suffix.match(/^\[\s*((?:[^\]]|]])+)\s*\]$/);
  if (simpleMatch) {
    // Excel escapes `]` inside structured reference items by doubling it: `]]` -> `]`.
    const columnName = simpleMatch[1].replaceAll("]]", "]").trim();
    return { tableName, selector: null, columnName, columns: null, columnMode: null };
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

    const selectorNorm = structured.selector ? normalizeName(structured.selector) : null;
    if (selectorNorm && selectorNorm !== "#ALL" && selectorNorm !== "#HEADERS" && selectorNorm !== "#DATA" && selectorNorm !== "#TOTALS") {
      // Selectors like `#This Row` are relative to an active row context, which `parseGoTo` does not have.
      // Reject them explicitly rather than returning a misleading full-table range.
      throw new Error(`Unsupported structured reference selector: ${structured.selector}`);
    }

    const normalizeColumns = (names) =>
      Array.isArray(names)
        ? names.map((c) => normalizeName(c)).filter((c) => typeof c === "string" && c.length > 0)
        : [];

    const columns = table.columns ?? [];
    const findColumnIndex = (name) => columns.findIndex((c) => normalizeName(c) === normalizeName(name));

    if (structured.columns && structured.columnMode) {
      const requested = normalizeColumns(structured.columns);
      if (requested.length === 0) throw new Error(`Invalid structured reference: ${input}`);

      let indices = [];
      if (structured.columnMode === "range") {
        if (requested.length !== 2) throw new Error(`Invalid structured reference: ${input}`);
        const a = findColumnIndex(structured.columns[0]);
        const b = findColumnIndex(structured.columns[1]);
        if (a === -1 || b === -1) throw new Error(`Unknown table column: ${structured.tableName}[[${structured.columns.join("]:[")}]]`);
        indices = [a, b];
      } else {
        indices = structured.columns.map((name) => findColumnIndex(name));
        if (indices.some((idx) => idx === -1)) {
          throw new Error(`Unknown table column in structured reference: ${structured.tableName}`);
        }
      }

      const uniq = Array.from(new Set(indices)).sort((a, b) => a - b);
      const minIdx = uniq[0];
      const maxIdx = uniq[uniq.length - 1];

      // Comma-separated multi-column refs represent a union. `parseGoTo` only returns a single range,
      // so only support contiguous columns to avoid selecting a misleading bounding rectangle.
      if (structured.columnMode === "list" && maxIdx - minIdx + 1 !== uniq.length) {
        throw new Error(`Non-contiguous structured reference columns are not supported: ${structured.tableName}`);
      }

      const startCol = table.startCol + minIdx;
      const endCol = table.startCol + maxIdx;

      let startRow = table.startRow;
      let endRow = table.endRow;

      if (selectorNorm === "#HEADERS") {
        endRow = startRow;
      } else if (selectorNorm === "#TOTALS") {
        startRow = endRow;
      } else if (selectorNorm === "#DATA") {
        if (endRow > startRow) startRow = startRow + 1;
      } else if (selectorNorm === "#ALL") {
        // keep full range
      }

      return { type: "range", source: "table", sheetName: tableSheetName, range: { startRow, endRow, startCol, endCol } };
    }

    const columnNorm = normalizeName(structured.columnName);

    // Whole-table specifiers: `Table1[#All]`, `Table1[#Headers]`, `Table1[#Data]`, `Table1[#Totals]`.
    if (columnNorm.startsWith("#") && !structured.columns) {
      if (columnNorm === "#ALL") {
        return {
          type: "range",
          source: "table",
          sheetName: tableSheetName,
          range: { startRow: table.startRow, endRow: table.endRow, startCol: table.startCol, endCol: table.endCol },
        };
      }
      if (columnNorm === "#HEADERS") {
        return {
          type: "range",
          source: "table",
          sheetName: tableSheetName,
          range: { startRow: table.startRow, endRow: table.startRow, startCol: table.startCol, endCol: table.endCol },
        };
      }
      if (columnNorm === "#DATA") {
        const startRow = table.endRow > table.startRow ? table.startRow + 1 : table.startRow;
        return {
          type: "range",
          source: "table",
          sheetName: tableSheetName,
          range: { startRow, endRow: table.endRow, startCol: table.startCol, endCol: table.endCol },
        };
      }
      if (columnNorm === "#TOTALS") {
        return {
          type: "range",
          source: "table",
          sheetName: tableSheetName,
          range: { startRow: table.endRow, endRow: table.endRow, startCol: table.startCol, endCol: table.endCol },
        };
      }
    }

    // Column references: `Table1[Col2]` or selector-qualified `Table1[[#Headers],[Col2]]`.
    const idx = columns.findIndex((c) => normalizeName(c) === columnNorm);
    if (idx === -1) {
      const suffix = selectorNorm ? `[[${structured.selector}],[${structured.columnName}]]` : `[${structured.columnName}]`;
      throw new Error(`Unknown table column: ${structured.tableName}${suffix}`);
    }

    const col = table.startCol + idx;
    let startRow = table.startRow;
    let endRow = table.endRow;

    // For selector-qualified column refs, adjust the row span. Unqualified `Table1[Col]`
    // intentionally selects the full column including headers (legacy behavior).
    if (selectorNorm === "#HEADERS") {
      endRow = startRow;
    } else if (selectorNorm === "#TOTALS") {
      startRow = endRow;
    } else if (selectorNorm === "#DATA") {
      if (endRow > startRow) startRow = startRow + 1;
    } else if (selectorNorm === "#ALL") {
      // keep full range (including headers)
    }

    return { type: "range", source: "table", sheetName: tableSheetName, range: { startRow, endRow, startCol: col, endCol: col } };
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
