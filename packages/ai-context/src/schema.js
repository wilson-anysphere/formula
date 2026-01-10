import { isCellEmpty, normalizeRange, rangeToA1 } from "./a1.js";

/**
 * @typedef {"empty"|"number"|"boolean"|"date"|"string"|"formula"|"mixed"} InferredType
 */

/**
 * @param {unknown} value
 * @returns {InferredType}
 */
export function inferCellType(value) {
  if (isCellEmpty(value)) return "empty";
  if (typeof value === "number" && Number.isFinite(value)) return "number";
  if (typeof value === "boolean") return "boolean";
  if (value instanceof Date && !Number.isNaN(value.getTime())) return "date";
  if (typeof value === "string") {
    const trimmed = value.trim();
    if (trimmed.startsWith("=")) return "formula";

    // Numeric-like strings are common in CSV imports. Treat them as numbers for schema purposes.
    if (/^[+-]?\d+(?:\.\d+)?$/.test(trimmed)) return "number";

    // ISO-like dates are also common.
    if (/^\d{4}-\d{2}-\d{2}/.test(trimmed)) {
      const parsed = new Date(trimmed);
      if (!Number.isNaN(parsed.getTime())) return "date";
    }

    return "string";
  }
  return "string";
}

/**
 * @param {unknown[]} values
 * @returns {InferredType}
 */
export function inferColumnType(values) {
  const types = new Set();
  for (const value of values) {
    const t = inferCellType(value);
    if (t !== "empty") types.add(t);
  }

  if (types.size === 0) return "empty";
  if (types.size === 1) return /** @type {InferredType} */ (types.values().next().value);

  // "formula" plus "number" is a common computed column.
  if (types.has("formula") && types.size === 2 && (types.has("number") || types.has("date") || types.has("string"))) {
    return "formula";
  }

  return "mixed";
}

/**
 * @param {unknown} value
 */
function isHeaderCandidateValue(value) {
  if (isCellEmpty(value)) return false;
  if (typeof value !== "string") return false;
  const trimmed = value.trim();
  if (!trimmed) return false;
  if (trimmed.startsWith("=")) return false;
  // Disqualify pure numbers masquerading as strings.
  if (/^[+-]?\d+(?:\.\d+)?$/.test(trimmed)) return false;
  return true;
}

/**
 * @param {unknown[]} rowValues
 * @param {unknown[] | undefined} nextRowValues
 */
export function isLikelyHeaderRow(rowValues, nextRowValues) {
  const nonEmpty = rowValues.filter((v) => !isCellEmpty(v));
  if (nonEmpty.length === 0) return false;

  const headerish = nonEmpty.filter(isHeaderCandidateValue);
  if (headerish.length / nonEmpty.length < 0.6) return false;

  const normalized = headerish.map((v) => String(v).trim().toLowerCase());
  const unique = new Set(normalized);
  if (unique.size !== normalized.length) return false;

  if (!nextRowValues) return true;
  const nextNonEmpty = nextRowValues.filter((v) => !isCellEmpty(v));
  if (nextNonEmpty.length === 0) return true;

  // If the next row is "more numeric" than the first row, it's likely data.
  const nextNumeric = nextNonEmpty.filter((v) => inferCellType(v) === "number").length;
  const nextStrings = nextNonEmpty.filter((v) => inferCellType(v) === "string").length;

  if (nextNumeric > 0) return true;
  if (nextStrings / nextNonEmpty.length < 0.6) return true;

  return false;
}

/**
 * @param {unknown[][]} values
 * @returns {{ startRow: number, startCol: number, endRow: number, endCol: number }[]}
 */
export function detectDataRegions(values) {
  const rowCount = values.length;
  const colCount = values.reduce((max, row) => Math.max(max, row?.length ?? 0), 0);

  /** @type {boolean[][]} */
  const visited = Array.from({ length: rowCount }, () => Array.from({ length: colCount }, () => false));

  /** @type {{ startRow: number, startCol: number, endRow: number, endCol: number }[]} */
  const regions = [];

  /**
   * @param {number} r
   * @param {number} c
   */
  function getValue(r, c) {
    return values[r]?.[c];
  }

  for (let r = 0; r < rowCount; r++) {
    for (let c = 0; c < colCount; c++) {
      if (visited[r][c]) continue;
      visited[r][c] = true;

      if (isCellEmpty(getValue(r, c))) continue;

      let minRow = r;
      let maxRow = r;
      let minCol = c;
      let maxCol = c;

      /** @type {[number, number][]} */
      const queue = [[r, c]];

      while (queue.length) {
        const [qr, qc] = queue.shift();
        if (qr < minRow) minRow = qr;
        if (qr > maxRow) maxRow = qr;
        if (qc < minCol) minCol = qc;
        if (qc > maxCol) maxCol = qc;

        const neighbors = [
          [qr - 1, qc],
          [qr + 1, qc],
          [qr, qc - 1],
          [qr, qc + 1],
        ];

        for (const [nr, nc] of neighbors) {
          if (nr < 0 || nr >= rowCount || nc < 0 || nc >= colCount) continue;
          if (visited[nr][nc]) continue;
          visited[nr][nc] = true;
          if (isCellEmpty(getValue(nr, nc))) continue;
          queue.push([nr, nc]);
        }
      }

      regions.push({ startRow: minRow, startCol: minCol, endRow: maxRow, endCol: maxCol });
    }
  }

  regions.sort((a, b) => (a.startRow - b.startRow) || (a.startCol - b.startCol));
  return regions;
}

/**
 * @typedef {{ name: string, range: string, columns: { name: string, type: InferredType, sampleValues: string[] }[], rowCount: number }} TableSchema
 * @typedef {{ name: string, range: string }} NamedRangeSchema
 * @typedef {{ range: string, hasHeader: boolean, headers: string[], inferredColumnTypes: InferredType[], rowCount: number, columnCount: number }} DataRegionSchema
 * @typedef {{ name: string, tables: TableSchema[], namedRanges: NamedRangeSchema[], dataRegions: DataRegionSchema[] }} SheetSchema
 */

/**
 * Extract a schema-first representation of a sheet suitable for LLM context.
 *
 * The input model is intentionally minimal: a 2D array of values plus optional metadata
 * (named ranges, structured tables). This makes the package usable before the full
 * spreadsheet engine is wired in.
 *
 * @param {{ name: string, values: unknown[][], namedRanges?: NamedRangeSchema[], tables?: { name: string, range: string }[] }} sheet
 * @returns {SheetSchema}
 */
export function extractSheetSchema(sheet) {
  const regions = detectDataRegions(sheet.values);

  /** @type {DataRegionSchema[]} */
  const dataRegions = [];
  /** @type {TableSchema[]} */
  const implicitTables = [];

  for (let i = 0; i < regions.length; i++) {
    const region = regions[i];
    const normalized = normalizeRange(region);
    const regionValues = slice2D(sheet.values, normalized);
    const headerRowValues = regionValues[0] ?? [];
    const nextRowValues = regionValues[1];
    const hasHeader = isLikelyHeaderRow(headerRowValues, nextRowValues);

    const headers = [];
    for (let c = 0; c < (headerRowValues?.length ?? 0); c++) {
      const raw = headerRowValues[c];
      const fallback = `Column${c + 1}`;
      headers.push(hasHeader && isHeaderCandidateValue(raw) ? String(raw).trim() : fallback);
    }

    const dataStartRow = hasHeader ? 1 : 0;
    const dataRows = regionValues.slice(dataStartRow);
    const columnCount = Math.max(...regionValues.map((row) => row.length), 0);

    /** @type {InferredType[]} */
    const inferredColumnTypes = [];
    /** @type {{ name: string, type: InferredType, sampleValues: string[] }[]} */
    const columns = [];

    for (let c = 0; c < columnCount; c++) {
      const colValues = dataRows.map((row) => row[c]).filter((v) => v !== undefined);
      const type = inferColumnType(colValues);
      inferredColumnTypes.push(type);

      const sampleValues = [];
      for (const v of colValues) {
        if (isCellEmpty(v)) continue;
        const s = String(v);
        if (!sampleValues.includes(s)) sampleValues.push(s);
        if (sampleValues.length >= 3) break;
      }

      columns.push({
        name: headers[c] ?? `Column${c + 1}`,
        type,
        sampleValues,
      });
    }

    const range = rangeToA1({ ...normalized, sheetName: sheet.name });
    const rowCount = Math.max(regionValues.length - (hasHeader ? 1 : 0), 0);

    dataRegions.push({
      range,
      hasHeader,
      headers,
      inferredColumnTypes,
      rowCount,
      columnCount,
    });

    implicitTables.push({
      name: `Region${i + 1}`,
      range,
      columns,
      rowCount,
    });
  }

  // TODO: When the spreadsheet engine provides explicit structured tables, reconcile
  // them with implicit region detection (avoid duplicates, prefer explicit names).

  return {
    name: sheet.name,
    tables: implicitTables,
    namedRanges: sheet.namedRanges ?? [],
    dataRegions,
  };
}

/**
 * @param {unknown[][]} values
 * @param {{ startRow: number, startCol: number, endRow: number, endCol: number }} range
 */
function slice2D(values, range) {
  /** @type {unknown[][]} */
  const out = [];
  for (let r = range.startRow; r <= range.endRow; r++) {
    const row = values[r] ?? [];
    out.push(row.slice(range.startCol, range.endCol + 1));
  }
  return out;
}
