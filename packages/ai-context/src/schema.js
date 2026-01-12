import { isCellEmpty, normalizeRange, parseA1Range, rangeToA1 } from "./a1.js";

// Detecting connected regions in a dense matrix requires an O(rows*cols) visited grid.
// Cap the scan to avoid catastrophic allocations when callers accidentally pass
// Excel-scale ranges (1,048,576 x 16,384).
const DEFAULT_DATA_REGION_SCAN_CELL_LIMIT = 200_000;

function createAbortError(message = "Aborted") {
  const err = new Error(message);
  err.name = "AbortError";
  return err;
}

function throwIfAborted(signal) {
  if (signal?.aborted) throw createAbortError();
}

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
 * @param {{ signal?: AbortSignal }} [options]
 * @returns {InferredType}
 */
export function inferColumnType(values, options = {}) {
  const signal = options.signal;
  const types = new Set();
  for (const value of values) {
    throwIfAborted(signal);
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
 * @param {{ maxCells?: number, signal?: AbortSignal }} [options]
 * @returns {{ startRow: number, startCol: number, endRow: number, endCol: number }[]}
 */
export function detectDataRegions(values, options = {}) {
  const signal = options.signal;
  throwIfAborted(signal);
  const maxCellsRaw = options.maxCells;
  const maxCells =
    typeof maxCellsRaw === "number" && Number.isFinite(maxCellsRaw) && maxCellsRaw > 0
      ? Math.floor(maxCellsRaw)
      : DEFAULT_DATA_REGION_SCAN_CELL_LIMIT;

  const rawRowCount = values.length;
  // Keep rows bounded so the outer loop can't become unbounded (even if the input
  // matrix is sparse / mostly empty).
  const rowCount = Math.max(0, Math.min(rawRowCount, maxCells));
  if (rowCount === 0) return [];

  // Only consider columns present in the scanned prefix of rows.
  let colCountRaw = 0;
  for (let r = 0; r < rowCount; r++) {
    colCountRaw = Math.max(colCountRaw, values[r]?.length ?? 0);
  }
  if (colCountRaw === 0) return [];

  // Prefer preserving rows (up to the safe cap) and clamp the column scan so the
  // visited bitmap stays within `maxCells`.
  const colCount = Math.max(0, Math.min(colCountRaw, Math.max(1, Math.floor(maxCells / rowCount))));

  if (rowCount === 0 || colCount === 0) return [];

  /**
   * `Array.shift()` is O(n) due to element re-indexing.
   * Dense regions (large contiguous blocks) would therefore take ~O(n^2) time to flood-fill.
   *
   * Use an index-based queue + typed visited grid to keep flood-fill linear in the number of
   * visited cells.
   *
   * Flattened indexing: `idx = row * colCount + col`.
   *
   * @type {Uint8Array}
   */
  const visited = new Uint8Array(rowCount * colCount);

  /** @type {{ startRow: number, startCol: number, endRow: number, endCol: number }[]} */
  const regions = [];

  for (let r = 0; r < rowCount; r++) {
    throwIfAborted(signal);
    const row = values[r];
    const rowOffset = r * colCount;
    for (let c = 0; c < colCount; c++) {
      throwIfAborted(signal);
      const startIdx = rowOffset + c;
      if (visited[startIdx]) continue;
      visited[startIdx] = 1;

      if (isCellEmpty(row?.[c])) continue;

      let minRow = r;
      let maxRow = r;
      let minCol = c;
      let maxCol = c;

      // Use a flat number queue to avoid allocating `[row, col]` tuples for every enqueued cell.
      /** @type {number[]} */
      const queue = [r, c];
      let head = 0;

      while (head < queue.length) {
        throwIfAborted(signal);
        const qr = queue[head++];
        const qc = queue[head++];
        if (qr < minRow) minRow = qr;
        if (qr > maxRow) maxRow = qr;
        if (qc < minCol) minCol = qc;
        if (qc > maxCol) maxCol = qc;

        const qRowOffset = qr * colCount;

        // Inline neighbor exploration to avoid per-cell allocations.
        // Up
        if (qr > 0) {
          const idx = qRowOffset - colCount + qc;
          if (!visited[idx]) {
            visited[idx] = 1;
          if (!isCellEmpty(values[qr - 1]?.[qc])) queue.push(qr - 1, qc);
          }
        }
        // Down
        if (qr + 1 < rowCount) {
          const idx = qRowOffset + colCount + qc;
          if (!visited[idx]) {
            visited[idx] = 1;
            if (!isCellEmpty(values[qr + 1]?.[qc])) queue.push(qr + 1, qc);
          }
        }
        // Left
        if (qc > 0) {
          const idx = qRowOffset + qc - 1;
          if (!visited[idx]) {
            visited[idx] = 1;
            if (!isCellEmpty(values[qr]?.[qc - 1])) queue.push(qr, qc - 1);
          }
        }
        // Right
        if (qc + 1 < colCount) {
          const idx = qRowOffset + qc + 1;
          if (!visited[idx]) {
            visited[idx] = 1;
            if (!isCellEmpty(values[qr]?.[qc + 1])) queue.push(qr, qc + 1);
          }
        }
      }

      regions.push({ startRow: minRow, startCol: minCol, endRow: maxRow, endCol: maxCol });
    }
  }

  throwIfAborted(signal);
  regions.sort((a, b) => (a.startRow - b.startRow) || (a.startCol - b.startCol));
  throwIfAborted(signal);
  return regions;
}

/**
 * @param {unknown[][]} sheetValues
 * @param {{ startRow: number, startCol: number, endRow: number, endCol: number }} normalized
 * @returns {{
 *   hasHeader: boolean,
 *   headers: string[],
 *   inferredColumnTypes: InferredType[],
 *   columns: { name: string, type: InferredType, sampleValues: string[] }[],
 *   rowCount: number,
 *   columnCount: number,
 * }}
 */
function analyzeRegion(sheetValues, normalized, signal) {
  const regionValues = slice2D(sheetValues, normalized, signal);
  const headerRowValues = regionValues[0] ?? [];
  const nextRowValues = regionValues[1];
  const hasHeader = isLikelyHeaderRow(headerRowValues, nextRowValues);

  const dataStartRow = hasHeader ? 1 : 0;
  const dataRows = regionValues.slice(dataStartRow);
  let columnCount = 0;
  for (const row of regionValues) {
    throwIfAborted(signal);
    columnCount = Math.max(columnCount, row.length);
  }

  const headers = [];
  for (let c = 0; c < columnCount; c++) {
    throwIfAborted(signal);
    const raw = headerRowValues[c];
    const fallback = `Column${c + 1}`;
    headers.push(hasHeader && isHeaderCandidateValue(raw) ? String(raw).trim() : fallback);
  }

  /** @type {InferredType[]} */
  const inferredColumnTypes = [];
  /** @type {{ name: string, type: InferredType, sampleValues: string[] }[]} */
  const columns = [];

  for (let c = 0; c < columnCount; c++) {
    throwIfAborted(signal);
    /** @type {unknown[]} */
    const colValues = [];
    for (const row of dataRows) {
      throwIfAborted(signal);
      const v = row?.[c];
      if (v !== undefined) colValues.push(v);
    }
    const type = inferColumnType(colValues, { signal });
    inferredColumnTypes.push(type);

    const sampleValues = [];
    for (const v of colValues) {
      throwIfAborted(signal);
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

  const rowCount = Math.max(regionValues.length - (hasHeader ? 1 : 0), 0);

  return {
    hasHeader,
    headers,
    inferredColumnTypes,
    columns,
    rowCount,
    columnCount,
  };
}

/**
 * @param {{ startRow: number, startCol: number, endRow: number, endCol: number }} outer
 * @param {{ startRow: number, startCol: number, endRow: number, endCol: number }} inner
 */
function rangeContains(outer, inner) {
  return (
    outer.startRow <= inner.startRow &&
    outer.startCol <= inner.startCol &&
    outer.endRow >= inner.endRow &&
    outer.endCol >= inner.endCol
  );
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
 * @param {{
 *   name: string,
 *   values: unknown[][],
 *   /**
 *    * Optional coordinate origin (0-based) for the provided `values` matrix.
 *    *
 *    * When `values` is a cropped window of a larger sheet (e.g. a capped used-range
 *    * sample), `origin` lets schema extraction produce correct absolute A1 ranges.
 *    *\/
 *   origin?: { row: number, col: number },
 *   namedRanges?: NamedRangeSchema[],
 *   tables?: { name: string, range: string }[]
 * }} sheet
 * @param {{ signal?: AbortSignal }} [options]
 * @returns {SheetSchema}
 */
export function extractSheetSchema(sheet, options = {}) {
  const signal = options.signal;
  throwIfAborted(signal);
  const origin = sheet && typeof sheet === "object" && sheet.origin && typeof sheet.origin === "object"
    ? {
        row: Number.isInteger(sheet.origin.row) && sheet.origin.row >= 0 ? sheet.origin.row : 0,
        col: Number.isInteger(sheet.origin.col) && sheet.origin.col >= 0 ? sheet.origin.col : 0,
      }
    : { row: 0, col: 0 };
  throwIfAborted(signal);
  const matrixRowCount = sheet.values.length;
  let matrixColCount = 0;
  for (const row of sheet.values) {
    throwIfAborted(signal);
    matrixColCount = Math.max(matrixColCount, row?.length ?? 0);
  }

  /**
   * Clamp a rect (0-based) to the bounds of `sheet.values`.
   *
   * Returns null when the rect does not intersect the provided matrix at all.
   *
   * @param {{ startRow: number, startCol: number, endRow: number, endCol: number }} rect
   */
  function clampRectToMatrix(rect) {
    if (matrixRowCount === 0 || matrixColCount === 0) return null;
    if (rect.endRow < 0 || rect.endCol < 0) return null;
    if (rect.startRow >= matrixRowCount || rect.startCol >= matrixColCount) return null;

    const startRow = Math.max(0, Math.min(rect.startRow, matrixRowCount - 1));
    const endRow = Math.max(0, Math.min(rect.endRow, matrixRowCount - 1));
    const startCol = Math.max(0, Math.min(rect.startCol, matrixColCount - 1));
    const endCol = Math.max(0, Math.min(rect.endCol, matrixColCount - 1));
    return { startRow, startCol, endRow, endCol };
  }
  const regions = detectDataRegions(sheet.values, { signal });

  /** @type {DataRegionSchema[]} */
  const dataRegions = [];
  /** @type {TableSchema[]} */
  const implicitTables = [];
  /** @type {{ startRow: number, startCol: number, endRow: number, endCol: number }[]} */
  const implicitTableRects = [];

  for (let i = 0; i < regions.length; i++) {
    throwIfAborted(signal);
    const region = regions[i];
    const normalized = normalizeRange(region);
    const analyzed = analyzeRegion(sheet.values, normalized, signal);
    const rect = {
      startRow: normalized.startRow + origin.row,
      endRow: normalized.endRow + origin.row,
      startCol: normalized.startCol + origin.col,
      endCol: normalized.endCol + origin.col,
    };
    const range = rangeToA1({ ...rect, sheetName: sheet.name });

    dataRegions.push({
      range,
      hasHeader: analyzed.hasHeader,
      headers: analyzed.headers,
      inferredColumnTypes: analyzed.inferredColumnTypes,
      rowCount: analyzed.rowCount,
      columnCount: analyzed.columnCount,
    });

    implicitTables.push({
      name: `Region${i + 1}`,
      range,
      columns: analyzed.columns,
      rowCount: analyzed.rowCount,
    });
    // Track the rect in absolute sheet coordinates so explicit table metadata (A1 ranges)
    // can be reconciled even when `values` is a cropped window (origin offset).
    implicitTableRects.push(rect);
  }

  /** @type {{ name: string, range: string, rect: { startRow: number, startCol: number, endRow: number, endCol: number } }[]} */
  const explicitDefs = [];

  if (sheet.tables?.length) {
    const seenRanges = new Set();
    for (const table of sheet.tables) {
      throwIfAborted(signal);
      if (!table || typeof table !== "object") continue;
      if (typeof table.range !== "string" || typeof table.name !== "string") continue;

      let parsed;
      try {
        parsed = parseA1Range(table.range);
      } catch {
        continue;
      }

      if (parsed.sheetName && parsed.sheetName !== sheet.name) continue;
      const rect = normalizeRange(parsed);
      const canonicalRange = rangeToA1({ ...rect, sheetName: sheet.name });

      if (seenRanges.has(canonicalRange)) continue;
      seenRanges.add(canonicalRange);
      explicitDefs.push({ name: table.name, range: canonicalRange, rect });
    }
  }

  /** @type {{ table: TableSchema, rect: { startRow: number, startCol: number, endRow: number, endCol: number } }[]} */
  const tableEntries = [];

  if (explicitDefs.length) {
    const coveredImplicit = new Set();
    for (let i = 0; i < implicitTableRects.length; i++) {
      throwIfAborted(signal);
      const implicitRect = implicitTableRects[i];
      for (const explicit of explicitDefs) {
        throwIfAborted(signal);
        if (rangeContains(explicit.rect, implicitRect)) {
          coveredImplicit.add(i);
          break;
        }
      }
    }

    const implicitUncovered = [];
    for (let i = 0; i < implicitTables.length; i++) {
      throwIfAborted(signal);
      if (coveredImplicit.has(i)) continue;
      implicitUncovered.push({ table: implicitTables[i], rect: implicitTableRects[i] });
    }

    // Re-number implicit regions so we don't end up with confusing gaps when some
    // regions are replaced by explicit tables.
    for (let i = 0; i < implicitUncovered.length; i++) {
      throwIfAborted(signal);
      implicitUncovered[i].table.name = `Region${i + 1}`;
    }

    tableEntries.push(...implicitUncovered);

    for (const explicit of explicitDefs) {
      throwIfAborted(signal);
      const localRect =
        origin.row === 0 && origin.col === 0
          ? explicit.rect
          : {
              startRow: explicit.rect.startRow - origin.row,
              endRow: explicit.rect.endRow - origin.row,
              startCol: explicit.rect.startCol - origin.col,
              endCol: explicit.rect.endCol - origin.col,
            };
      const clamped = clampRectToMatrix(localRect);
      const analyzed = clamped ? analyzeRegion(sheet.values, clamped, signal) : { columns: [], rowCount: 0 };
      tableEntries.push({
        table: {
          name: explicit.name,
          range: explicit.range,
          columns: analyzed.columns,
          rowCount: analyzed.rowCount,
        },
        rect: explicit.rect,
      });
    }

    throwIfAborted(signal);
    tableEntries.sort((a, b) => (a.rect.startRow - b.rect.startRow) || (a.rect.startCol - b.rect.startCol));
  } else {
    tableEntries.push(...implicitTables.map((t, idx) => ({ table: t, rect: implicitTableRects[idx] })));
  }

  throwIfAborted(signal);
  return {
    name: sheet.name,
    tables: tableEntries.map((t) => t.table),
    namedRanges: sheet.namedRanges ?? [],
    dataRegions,
  };
}

/**
 * @param {unknown[][]} values
 * @param {{ startRow: number, startCol: number, endRow: number, endCol: number }} range
 */
function slice2D(values, range, signal) {
  /** @type {unknown[][]} */
  const out = [];
  for (let r = range.startRow; r <= range.endRow; r++) {
    throwIfAborted(signal);
    const row = values[r] ?? [];
    out.push(row.slice(range.startCol, range.endCol + 1));
  }
  return out;
}
