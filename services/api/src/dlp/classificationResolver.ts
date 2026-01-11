import { DEFAULT_CLASSIFICATION, classificationRank, maxClassification, normalizeClassification, normalizeRange } from "../../../../shared/dlp-core";
import { selectorKey as coreSelectorKey } from "../../../../shared/dlp-core";
import type { Classification } from "../../../../shared/dlp-core";
import type { Pool } from "pg";

export type DbClient = Pick<Pool, "query">;

export type ClassificationResolutionSource = {
  scope: string;
  selectorKey: string;
};

export type ClassificationResolutionResult = {
  classification: Classification;
  source: ClassificationResolutionSource;
};

export type NormalizedSelectorColumns = {
  scope: string;
  sheetId: string | null;
  tableId: string | null;
  row: number | null;
  col: number | null;
  startRow: number | null;
  startCol: number | null;
  endRow: number | null;
  endCol: number | null;
  columnIndex: number | null;
  columnId: string | null;
};

const DEFAULT_RESOLUTION: ClassificationResolutionResult = {
  classification: { ...DEFAULT_CLASSIFICATION },
  source: { scope: "default", selectorKey: "default" }
};

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function requireNonEmptyString(value: unknown, name: string): string {
  if (typeof value !== "string" || value.trim().length === 0) throw new Error(`${name} must be a non-empty string`);
  return value;
}

function requireNonNegativeInt(value: unknown, name: string): number {
  if (typeof value !== "number" || !Number.isInteger(value) || value < 0) {
    throw new Error(`${name} must be a non-negative integer`);
  }
  return value;
}

/**
 * Normalize a selector into the dedicated columns stored in `document_classifications`.
 *
 * The API writes these columns on upsert to make containment/overlap queries fast.
 */
export function normalizeSelectorColumns(selector: unknown): NormalizedSelectorColumns {
  if (!isRecord(selector)) throw new Error("Selector must be an object");
  const scope = String(selector.scope ?? "");
  if (!scope) throw new Error("Selector.scope must be a string");

  switch (scope) {
    case "document":
      return {
        scope,
        sheetId: null,
        tableId: null,
        row: null,
        col: null,
        startRow: null,
        startCol: null,
        endRow: null,
        endCol: null,
        columnIndex: null,
        columnId: null
      };
    case "sheet":
      return {
        scope,
        sheetId: requireNonEmptyString(selector.sheetId, "Selector.sheetId"),
        tableId: null,
        row: null,
        col: null,
        startRow: null,
        startCol: null,
        endRow: null,
        endCol: null,
        columnIndex: null,
        columnId: null
      };
    case "column":
      {
        const sheetId = requireNonEmptyString(selector.sheetId, "Selector.sheetId");
        const tableId = selector.tableId == null ? null : requireNonEmptyString(selector.tableId, "Selector.tableId");

        const columnIndex = selector.columnIndex === undefined ? null : requireNonNegativeInt(selector.columnIndex, "Selector.columnIndex");
        const columnId = selector.columnId == null ? null : requireNonEmptyString(selector.columnId, "Selector.columnId");
        if (columnIndex === null && columnId === null) {
          throw new Error("Column selector must include columnIndex or columnId");
        }

        return {
        scope,
          sheetId,
          tableId,
        row: null,
        col: null,
        startRow: null,
        startCol: null,
        endRow: null,
        endCol: null,
          columnIndex,
          columnId
        };
      }
    case "cell":
      {
        const sheetId = requireNonEmptyString(selector.sheetId, "Selector.sheetId");
        const row = requireNonNegativeInt(selector.row, "Selector.row");
        const col = requireNonNegativeInt(selector.col, "Selector.col");

        return {
        scope,
          sheetId,
        tableId: null,
          row,
          col,
        startRow: null,
        startCol: null,
        endRow: null,
        endCol: null,
        columnIndex: null,
        columnId: null
        };
      }
    case "range": {
      const sheetId = requireNonEmptyString(selector.sheetId, "Selector.sheetId");

      if (!isRecord(selector.range)) throw new Error("Selector.range must be an object");
      if (!isRecord(selector.range.start)) throw new Error("Selector.range.start must be an object");
      if (!isRecord(selector.range.end)) throw new Error("Selector.range.end must be an object");

      // Validate/normalize using the shared helper once the coordinates are known-good.
      requireNonNegativeInt(selector.range.start.row, "Selector.range.start.row");
      requireNonNegativeInt(selector.range.start.col, "Selector.range.start.col");
      requireNonNegativeInt(selector.range.end.row, "Selector.range.end.row");
      requireNonNegativeInt(selector.range.end.col, "Selector.range.end.col");

      const range = normalizeRange(selector.range as any);
      return {
        scope,
        sheetId,
        tableId: null,
        row: null,
        col: null,
        startRow: range.start.row,
        startCol: range.start.col,
        endRow: range.end.row,
        endCol: range.end.col,
        columnIndex: null,
        columnId: null
      };
    }
    default:
      throw new Error(`Unknown selector scope: ${scope}`);
  }
}

type RangeRow = {
  selector_key: string;
  classification: unknown;
  start_row: number;
  start_col: number;
  end_row: number;
  end_col: number;
};

function rangeArea(row: RangeRow): number {
  const rows = row.end_row - row.start_row + 1;
  const cols = row.end_col - row.start_col + 1;
  return rows * cols;
}

function maxRank(level: string): number {
  return classificationRank(level as any);
}

function resolveFromExactRows(rows: Array<{ selector_key: string; classification: unknown; scope: string }>): ClassificationResolutionResult | null {
  if (rows.length === 0) return null;

  let best = { ...DEFAULT_CLASSIFICATION };
  let bestRankValue = -1;
  let bestKey: string | null = null;
  let bestScope: string | null = null;

  for (const row of rows) {
    const normalized = normalizeClassification(row.classification);
    best = maxClassification(best, normalized);
    const rank = maxRank(normalized.level);
    if (rank > bestRankValue || (rank === bestRankValue && bestKey !== null && row.selector_key < bestKey)) {
      bestRankValue = rank;
      bestKey = row.selector_key;
      bestScope = row.scope;
    }
    if (best.level === "Restricted") break;
  }

  return {
    classification: best,
    source: {
      scope: bestScope ?? rows[0].scope,
      selectorKey: bestKey ?? rows[0].selector_key
    }
  };
}

function resolveFromContainingRanges(rows: RangeRow[]): ClassificationResolutionResult | null {
  if (rows.length === 0) return null;

  let minArea = Infinity;
  for (const row of rows) {
    const area = rangeArea(row);
    if (area < minArea) minArea = area;
  }

  let best = { ...DEFAULT_CLASSIFICATION };
  let bestRankValue = -1;
  let bestKey: string | null = null;

  for (const row of rows) {
    if (rangeArea(row) !== minArea) continue;
    const normalized = normalizeClassification(row.classification);
    best = maxClassification(best, normalized);
    const rank = maxRank(normalized.level);
    if (rank > bestRankValue || (rank === bestRankValue && bestKey !== null && row.selector_key < bestKey)) {
      bestRankValue = rank;
      bestKey = row.selector_key;
    }
    if (best.level === "Restricted") break;
  }

  if (!bestKey) {
    // Should be impossible, but keep the resolver deterministic.
    bestKey = rows[0].selector_key;
  }

  return { classification: best, source: { scope: "range", selectorKey: bestKey } };
}

async function queryDocument(db: DbClient, docId: string): Promise<ClassificationResolutionResult | null> {
  const res = await db.query(
    `
      SELECT selector_key, classification
      FROM document_classifications
      WHERE document_id = $1 AND scope = 'document'
    `,
    [docId]
  );
  const rows = res.rows.map((row) => ({ selector_key: row.selector_key as string, classification: row.classification, scope: "document" }));
  return resolveFromExactRows(rows);
}

async function querySheet(db: DbClient, docId: string, sheetId: string): Promise<ClassificationResolutionResult | null> {
  const res = await db.query(
    `
      SELECT selector_key, classification
      FROM document_classifications
      WHERE document_id = $1 AND scope = 'sheet' AND sheet_id = $2
    `,
    [docId, sheetId]
  );
  const rows = res.rows.map((row) => ({ selector_key: row.selector_key as string, classification: row.classification, scope: "sheet" }));
  return resolveFromExactRows(rows);
}

async function queryColumn(params: {
  db: DbClient;
  docId: string;
  sheetId: string;
  columnIndex?: number;
  columnId?: string;
  tableId?: string | null;
  allowTableColumns: boolean;
}): Promise<ClassificationResolutionResult | null> {
  const { db, docId, sheetId, columnIndex, columnId, tableId, allowTableColumns } = params;
  const tableClause = allowTableColumns
    ? tableId
      ? { sql: "AND table_id = $4", args: [tableId] }
      : { sql: "AND table_id IS NULL", args: [] }
    : { sql: "AND table_id IS NULL", args: [] };

  if (typeof columnIndex === "number") {
    const args = [docId, sheetId, columnIndex];
    const res = await db.query(
      `
        SELECT selector_key, classification
        FROM document_classifications
        WHERE document_id = $1 AND scope = 'column' AND sheet_id = $2 AND column_index = $3
        ${tableClause.sql}
      `,
      tableClause.args.length ? [...args, ...tableClause.args] : args
    );
    const rows = res.rows.map((row) => ({ selector_key: row.selector_key as string, classification: row.classification, scope: "column" }));
    return resolveFromExactRows(rows);
  }

  if (columnId) {
    const args = [docId, sheetId, columnId];
    const res = await db.query(
      `
        SELECT selector_key, classification
        FROM document_classifications
        WHERE document_id = $1 AND scope = 'column' AND sheet_id = $2 AND column_id = $3
        ${tableClause.sql}
      `,
      tableClause.args.length ? [...args, ...tableClause.args] : args
    );
    const rows = res.rows.map((row) => ({ selector_key: row.selector_key as string, classification: row.classification, scope: "column" }));
    return resolveFromExactRows(rows);
  }

  return null;
}

async function queryCell(
  db: DbClient,
  docId: string,
  sheetId: string,
  row: number,
  col: number
): Promise<ClassificationResolutionResult | null> {
  const res = await db.query(
    `
      SELECT selector_key, classification
      FROM document_classifications
      WHERE document_id = $1 AND scope = 'cell' AND sheet_id = $2 AND row = $3 AND col = $4
    `,
    [docId, sheetId, row, col]
  );
  const rows = res.rows.map((r) => ({ selector_key: r.selector_key as string, classification: r.classification, scope: "cell" }));
  return resolveFromExactRows(rows);
}

async function queryContainingRangesForCell(
  db: DbClient,
  docId: string,
  sheetId: string,
  row: number,
  col: number
): Promise<ClassificationResolutionResult | null> {
  const res = await db.query(
    `
      SELECT selector_key, classification, start_row, start_col, end_row, end_col
      FROM document_classifications
      WHERE
        document_id = $1
        AND scope = 'range'
        AND sheet_id = $2
        AND start_row <= $3 AND end_row >= $3
        AND start_col <= $4 AND end_col >= $4
    `,
    [docId, sheetId, row, col]
  );

  const rows = res.rows.map((r) => ({
    selector_key: r.selector_key as string,
    classification: r.classification,
    start_row: r.start_row as number,
    start_col: r.start_col as number,
    end_row: r.end_row as number,
    end_col: r.end_col as number
  }));

  return resolveFromContainingRanges(rows);
}

async function queryContainingRangesForRange(
  db: DbClient,
  docId: string,
  sheetId: string,
  startRow: number,
  startCol: number,
  endRow: number,
  endCol: number
): Promise<ClassificationResolutionResult | null> {
  const res = await db.query(
    `
      SELECT selector_key, classification, start_row, start_col, end_row, end_col
      FROM document_classifications
      WHERE
        document_id = $1
        AND scope = 'range'
        AND sheet_id = $2
        AND start_row <= $3 AND end_row >= $4
        AND start_col <= $5 AND end_col >= $6
    `,
    [docId, sheetId, startRow, endRow, startCol, endCol]
  );

  const rows = res.rows.map((r) => ({
    selector_key: r.selector_key as string,
    classification: r.classification,
    start_row: r.start_row as number,
    start_col: r.start_col as number,
    end_row: r.end_row as number,
    end_col: r.end_col as number
  }));

  return resolveFromContainingRanges(rows);
}

/**
 * Resolve the effective classification for a selector.
 *
 * Precedence (most specific wins):
 * 1) cell (exact)
 * 2) range (smallest containing range; tie -> maxClassification)
 * 3) column
 * 4) sheet
 * 5) document
 * 6) default = Public
 */
export async function getEffectiveClassificationForSelector(
  db: DbClient,
  docId: string,
  selector: unknown
): Promise<ClassificationResolutionResult> {
  if (!isRecord(selector)) throw new Error("Selector must be an object");
  if (selector.documentId !== docId) throw new Error("Selector documentId must match docId");

  const normalized = normalizeSelectorColumns(selector);

  switch (normalized.scope) {
    case "document": {
      return (await queryDocument(db, docId)) ?? DEFAULT_RESOLUTION;
    }
    case "sheet": {
      return (
        (await querySheet(db, docId, normalized.sheetId!)) ??
        (await queryDocument(db, docId)) ??
        DEFAULT_RESOLUTION
      );
    }
    case "column": {
      const result =
        (await queryColumn({
          db,
          docId,
          sheetId: normalized.sheetId!,
          columnIndex: normalized.columnIndex ?? undefined,
          columnId: normalized.columnId ?? undefined,
          tableId: normalized.tableId,
          allowTableColumns: true
        })) ??
        (await querySheet(db, docId, normalized.sheetId!)) ??
        (await queryDocument(db, docId));
      return result ?? DEFAULT_RESOLUTION;
    }
    case "cell": {
      const result =
        (await queryCell(db, docId, normalized.sheetId!, normalized.row!, normalized.col!)) ??
        (await queryContainingRangesForCell(db, docId, normalized.sheetId!, normalized.row!, normalized.col!)) ??
        (await queryColumn({
          db,
          docId,
          sheetId: normalized.sheetId!,
          columnIndex: normalized.col!,
          allowTableColumns: false
        })) ??
        (await querySheet(db, docId, normalized.sheetId!)) ??
        (await queryDocument(db, docId));
      return result ?? DEFAULT_RESOLUTION;
    }
    case "range": {
      if (normalized.startRow === normalized.endRow && normalized.startCol === normalized.endCol) {
        return getEffectiveClassificationForSelector(db, docId, {
          scope: "cell",
          documentId: docId,
          sheetId: normalized.sheetId,
          row: normalized.startRow,
          col: normalized.startCol
        });
      }

      const columnFallback = normalized.startCol === normalized.endCol;
      const result =
        (await queryContainingRangesForRange(
          db,
          docId,
          normalized.sheetId!,
          normalized.startRow!,
          normalized.startCol!,
          normalized.endRow!,
          normalized.endCol!
        )) ??
        (columnFallback
          ? await queryColumn({
              db,
              docId,
              sheetId: normalized.sheetId!,
              columnIndex: normalized.startCol!,
              allowTableColumns: false
            })
          : null) ??
        (await querySheet(db, docId, normalized.sheetId!)) ??
        (await queryDocument(db, docId));
      return result ?? DEFAULT_RESOLUTION;
    }
    default:
      throw new Error(`Unknown selector scope: ${normalized.scope}`);
  }
}

export async function getAggregateClassificationForRange(
  db: DbClient,
  docId: string,
  sheetId: string,
  startRow: number,
  startCol: number,
  endRow: number,
  endCol: number
): Promise<Classification> {
  const range = normalizeRange({ start: { row: startRow, col: startCol }, end: { row: endRow, col: endCol } });

  let result: Classification = { ...DEFAULT_CLASSIFICATION };

  const applyRows = (rows: Array<{ classification: unknown }>) => {
    for (const row of rows) {
      result = maxClassification(result, normalizeClassification(row.classification));
      if (result.level === "Restricted") return true;
    }
    return false;
  };

  const docRes = await db.query(
    `
      SELECT classification
      FROM document_classifications
      WHERE document_id = $1 AND scope = 'document'
    `,
    [docId]
  );
  if (applyRows(docRes.rows as any)) return result;

  const sheetRes = await db.query(
    `
      SELECT classification
      FROM document_classifications
      WHERE document_id = $1 AND scope = 'sheet' AND sheet_id = $2
    `,
    [docId, sheetId]
  );
  if (applyRows(sheetRes.rows as any)) return result;

  const colRes = await db.query(
    `
      SELECT classification
      FROM document_classifications
      WHERE
        document_id = $1
        AND scope = 'column'
        AND sheet_id = $2
        AND table_id IS NULL
        AND column_index BETWEEN $3 AND $4
    `,
    [docId, sheetId, range.start.col, range.end.col]
  );
  if (applyRows(colRes.rows as any)) return result;

  const rangeRes = await db.query(
    `
      SELECT classification
      FROM document_classifications
      WHERE
        document_id = $1
        AND scope = 'range'
        AND sheet_id = $2
        AND start_row <= $3 AND end_row >= $4
        AND start_col <= $5 AND end_col >= $6
    `,
    [docId, sheetId, range.end.row, range.start.row, range.end.col, range.start.col]
  );
  if (applyRows(rangeRes.rows as any)) return result;

  const cellRes = await db.query(
    `
      SELECT classification
      FROM document_classifications
      WHERE
        document_id = $1
        AND scope = 'cell'
        AND sheet_id = $2
        AND row BETWEEN $3 AND $4
        AND col BETWEEN $5 AND $6
    `,
    [docId, sheetId, range.start.row, range.end.row, range.start.col, range.end.col]
  );
  applyRows(cellRes.rows as any);

  return result;
}

/**
 * Convenience helper for APIs that need to return selector_key for a selector object.
 * Prefer this over duplicating the shared selectorKey call site.
 */
export function selectorKey(selector: unknown): string {
  return coreSelectorKey(selector as any);
}
