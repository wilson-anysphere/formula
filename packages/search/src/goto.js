import { parseA1Range, splitSheetQualifier } from "./a1.js";
import { normalizeName } from "./workbook.js";

function isLikelyA1(ref) {
  return /^(\$?[A-Za-z]{1,3}\$?\d+)(:\$?[A-Za-z]{1,3}\$?\d+)?$/.test(String(ref).trim());
}

function parseStructuredRef(input) {
  const s = String(input).trim();

  // Minimal subset: TableName[ColumnName] or TableName[#All]
  const m = s.match(/^([A-Za-z_][A-Za-z0-9_]*)\[(.+)\]$/);
  if (!m) return null;

  return { tableName: m[1], spec: m[2] };
}

/**
 * Parse Go To / Name box input.
 *
 * Supports:
 * - A1 references: `A1`, `A1:B2`
 * - Sheet-qualified: `Sheet2!C3`, `'My Sheet'!A1`
 * - Named ranges: `MyName`
 * - Table structured refs (minimal): `Table1[Column]`, `Table1[#All]`
 */
export function parseGoTo(input, { workbook, currentSheetName } = {}) {
  if (!workbook) throw new Error("parseGoTo: workbook is required");
  if (!currentSheetName) throw new Error("parseGoTo: currentSheetName is required");

  const raw = String(input).trim();
  if (raw === "") throw new Error("parseGoTo: empty input");

  const { sheetName: qualifiedSheet, ref } = splitSheetQualifier(raw);
  const sheetName = qualifiedSheet ?? currentSheetName;

  // Structured reference
  const structured = parseStructuredRef(ref);
  if (structured) {
    const table = workbook.getTable(structured.tableName);
    if (!table) throw new Error(`Unknown table: ${structured.tableName}`);

    const specNorm = normalizeName(structured.spec);
    if (specNorm === "#ALL") {
      return {
        type: "range",
        source: "table",
        sheetName: table.sheetName,
        range: {
          startRow: table.startRow,
          endRow: table.endRow,
          startCol: table.startCol,
          endCol: table.endCol,
        },
      };
    }

    const columns = table.columns ?? [];
    const idx = columns.findIndex((c) => normalizeName(c) === specNorm);
    if (idx === -1) {
      throw new Error(`Unknown table column: ${structured.tableName}[${structured.spec}]`);
    }

    const col = table.startCol + idx;
    return {
      type: "range",
      source: "table",
      sheetName: table.sheetName,
      range: { startRow: table.startRow, endRow: table.endRow, startCol: col, endCol: col },
    };
  }

  // A1 reference
  if (isLikelyA1(ref)) {
    return { type: "range", source: "a1", sheetName, range: parseA1Range(ref) };
  }

  // Named range
  const named = workbook.getName(ref);
  if (named) {
    return {
      type: "range",
      source: "name",
      sheetName: named.sheetName ?? sheetName,
      range: named.range,
    };
  }

  throw new Error(`Unrecognized Go To reference: ${input}`);
}
