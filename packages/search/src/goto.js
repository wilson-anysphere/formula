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
  const qualifiedMatch = suffix.match(/^\[\[\s*(#[A-Za-z]+)\s*\]\s*,\s*\[\s*([^\]]+?)\s*\]\]$/i);
  if (qualifiedMatch) {
    return { tableName, selector: qualifiedMatch[1], columnName: qualifiedMatch[2] };
  }

  const simpleMatch = suffix.match(/^\[\s*([^\[\]]+?)\s*\]$/);
  if (simpleMatch) {
    return { tableName, selector: null, columnName: simpleMatch[1] };
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

    const columnNorm = normalizeName(structured.columnName);
    const selectorNorm = structured.selector ? normalizeName(structured.selector) : null;

    // Whole-table specifiers: `Table1[#All]`, `Table1[#Headers]`, `Table1[#Data]`, `Table1[#Totals]`.
    if (!selectorNorm && columnNorm.startsWith("#")) {
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
    const columns = table.columns ?? [];
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
