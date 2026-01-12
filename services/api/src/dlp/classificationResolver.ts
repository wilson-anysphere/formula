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
        const tableId = selector.tableId == null ? null : requireNonEmptyString(selector.tableId, "Selector.tableId");
        const columnId = selector.columnId == null ? null : requireNonEmptyString(selector.columnId, "Selector.columnId");

        return {
          scope,
          sheetId,
          tableId,
          row,
          col,
          startRow: null,
          startCol: null,
          endRow: null,
          endCol: null,
          columnIndex: null,
          columnId
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

type ExactCandidateRow = { scope: string; selector_key: string; classification: unknown };

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

async function queryCandidatesForSheet(db: DbClient, docId: string, sheetId: string): Promise<ExactCandidateRow[]> {
  const res = await db.query(
    `
      SELECT scope, selector_key, classification
      FROM document_classifications
      WHERE document_id = $1 AND scope = 'sheet' AND sheet_id = $2
      UNION ALL
      SELECT scope, selector_key, classification
      FROM document_classifications
      WHERE document_id = $1 AND scope = 'document'
    `,
    [docId, sheetId]
  );
  return res.rows as ExactCandidateRow[];
}

async function queryCandidatesForColumn(params: {
  db: DbClient;
  docId: string;
  sheetId: string;
  columnIndex?: number;
  columnId?: string;
  tableId?: string | null;
}): Promise<ExactCandidateRow[]> {
  const { db, docId, sheetId, columnIndex, columnId, tableId } = params;
  const tableClause = tableId ? "AND table_id = $4" : "AND table_id IS NULL";

  if (typeof columnIndex === "number") {
    const res = await db.query(
      `
        SELECT scope, selector_key, classification
        FROM document_classifications
        WHERE
          document_id = $1
          AND scope = 'column'
          AND sheet_id = $2
          AND column_index = $3
          ${tableClause}
        UNION ALL
        SELECT scope, selector_key, classification
        FROM document_classifications
        WHERE document_id = $1 AND scope = 'sheet' AND sheet_id = $2
        UNION ALL
        SELECT scope, selector_key, classification
        FROM document_classifications
        WHERE document_id = $1 AND scope = 'document'
      `,
      tableId ? [docId, sheetId, columnIndex, tableId] : [docId, sheetId, columnIndex]
    );
    return res.rows as ExactCandidateRow[];
  }

  if (columnId) {
    const res = await db.query(
      `
        SELECT scope, selector_key, classification
        FROM document_classifications
        WHERE
          document_id = $1
          AND scope = 'column'
          AND sheet_id = $2
          AND column_id = $3
          ${tableClause}
        UNION ALL
        SELECT scope, selector_key, classification
        FROM document_classifications
        WHERE document_id = $1 AND scope = 'sheet' AND sheet_id = $2
        UNION ALL
        SELECT scope, selector_key, classification
        FROM document_classifications
        WHERE document_id = $1 AND scope = 'document'
      `,
      tableId ? [docId, sheetId, columnId, tableId] : [docId, sheetId, columnId]
    );
    return res.rows as ExactCandidateRow[];
  }

  return [];
}

type CellResolutionCandidateRow = {
  scope: string;
  selector_key: string;
  classification: unknown;
  table_id: string | null;
  start_row: number | null;
  start_col: number | null;
  end_row: number | null;
  end_col: number | null;
};

async function queryCandidatesForCell(
  db: DbClient,
  docId: string,
  sheetId: string,
  row: number,
  col: number,
  tableId: string | null,
  columnId: string | null
): Promise<CellResolutionCandidateRow[]> {
  const res = await db.query(
    `
      SELECT scope, selector_key, classification, table_id, start_row, start_col, end_row, end_col
      FROM document_classifications
      WHERE document_id = $1 AND scope = 'cell' AND sheet_id = $2 AND row = $3 AND col = $4
      UNION ALL
      SELECT scope, selector_key, classification, table_id, start_row, start_col, end_row, end_col
      FROM document_classifications
      WHERE
        document_id = $1
        AND scope = 'range'
        AND sheet_id = $2
        AND start_row <= $3 AND end_row >= $3
        AND start_col <= $4 AND end_col >= $4
      UNION ALL
      SELECT scope, selector_key, classification, table_id, start_row, start_col, end_row, end_col
      FROM document_classifications
      WHERE
        document_id = $1
        AND scope = 'column'
        AND sheet_id = $2
        AND table_id IS NULL
        AND column_index = $4
      UNION ALL
      SELECT scope, selector_key, classification, table_id, start_row, start_col, end_row, end_col
      FROM document_classifications
      WHERE
        document_id = $1
        AND scope = 'column'
        AND sheet_id = $2
        AND table_id IS NULL
        AND column_id = $6
      UNION ALL
      SELECT scope, selector_key, classification, table_id, start_row, start_col, end_row, end_col
      FROM document_classifications
      WHERE
        document_id = $1
        AND scope = 'column'
        AND sheet_id = $2
        AND table_id = $5
        AND column_index = $4
      UNION ALL
      SELECT scope, selector_key, classification, table_id, start_row, start_col, end_row, end_col
      FROM document_classifications
      WHERE
        document_id = $1
        AND scope = 'column'
        AND sheet_id = $2
        AND table_id = $5
        AND column_id = $6
      UNION ALL
      SELECT scope, selector_key, classification, table_id, start_row, start_col, end_row, end_col
      FROM document_classifications
      WHERE document_id = $1 AND scope = 'sheet' AND sheet_id = $2
      UNION ALL
      SELECT scope, selector_key, classification, table_id, start_row, start_col, end_row, end_col
      FROM document_classifications
      WHERE document_id = $1 AND scope = 'document'
    `,
    [docId, sheetId, row, col, tableId, columnId]
  );

  return res.rows as CellResolutionCandidateRow[];
}

type RangeResolutionCandidateRow = {
  scope: string;
  selector_key: string;
  classification: unknown;
  start_row: number | null;
  start_col: number | null;
  end_row: number | null;
  end_col: number | null;
};

async function queryCandidatesForRange(
  db: DbClient,
  docId: string,
  sheetId: string,
  startRow: number,
  startCol: number,
  endRow: number,
  endCol: number
): Promise<RangeResolutionCandidateRow[]> {
  const res = await db.query(
    `
      SELECT scope, selector_key, classification, start_row, start_col, end_row, end_col
      FROM document_classifications
      WHERE
        document_id = $1
        AND scope = 'range'
        AND sheet_id = $2
        AND start_row <= $3 AND end_row >= $4
        AND start_col <= $5 AND end_col >= $6
      UNION ALL
      SELECT scope, selector_key, classification, start_row, start_col, end_row, end_col
      FROM document_classifications
      WHERE
        document_id = $1
        AND scope = 'column'
        AND sheet_id = $2
        AND table_id IS NULL
        AND column_index = $5
      UNION ALL
      SELECT scope, selector_key, classification, start_row, start_col, end_row, end_col
      FROM document_classifications
      WHERE document_id = $1 AND scope = 'sheet' AND sheet_id = $2
      UNION ALL
      SELECT scope, selector_key, classification, start_row, start_col, end_row, end_col
      FROM document_classifications
      WHERE document_id = $1 AND scope = 'document'
    `,
    [docId, sheetId, startRow, endRow, startCol, endCol]
  );

  return res.rows as RangeResolutionCandidateRow[];
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
      const candidates = await queryCandidatesForSheet(db, docId, normalized.sheetId!);
      const sheetRows = candidates
        .filter((row) => row.scope === "sheet")
        .map((row) => ({ selector_key: row.selector_key, classification: row.classification, scope: "sheet" }));
      const sheetResolved = resolveFromExactRows(sheetRows);
      if (sheetResolved) return sheetResolved;

      const docRows = candidates
        .filter((row) => row.scope === "document")
        .map((row) => ({ selector_key: row.selector_key, classification: row.classification, scope: "document" }));
      const docResolved = resolveFromExactRows(docRows);
      return docResolved ?? DEFAULT_RESOLUTION;
    }
    case "column": {
      const candidates = await queryCandidatesForColumn({
        db,
        docId,
        sheetId: normalized.sheetId!,
        columnIndex: normalized.columnIndex ?? undefined,
        columnId: normalized.columnId ?? undefined,
        tableId: normalized.tableId,
      });

      const columnRows = candidates
        .filter((row) => row.scope === "column")
        .map((row) => ({ selector_key: row.selector_key, classification: row.classification, scope: "column" }));
      const columnResolved = resolveFromExactRows(columnRows);
      if (columnResolved) return columnResolved;

      const sheetRows = candidates
        .filter((row) => row.scope === "sheet")
        .map((row) => ({ selector_key: row.selector_key, classification: row.classification, scope: "sheet" }));
      const sheetResolved = resolveFromExactRows(sheetRows);
      if (sheetResolved) return sheetResolved;

      const docRows = candidates
        .filter((row) => row.scope === "document")
        .map((row) => ({ selector_key: row.selector_key, classification: row.classification, scope: "document" }));
      const docResolved = resolveFromExactRows(docRows);
      return docResolved ?? DEFAULT_RESOLUTION;
    }
    case "cell": {
      const candidateRows = await queryCandidatesForCell(
        db,
        docId,
        normalized.sheetId!,
        normalized.row!,
        normalized.col!,
        normalized.tableId,
        normalized.columnId
      );

      const cellRows = candidateRows
        .filter((row) => row.scope === "cell")
        .map((row) => ({ selector_key: row.selector_key, classification: row.classification, scope: "cell" }));
      const exact = resolveFromExactRows(cellRows);
      if (exact) return exact;

      const rangeRows = candidateRows
        .filter((row) => row.scope === "range")
        .map((row) => ({
          selector_key: row.selector_key,
          classification: row.classification,
          start_row: row.start_row as number,
          start_col: row.start_col as number,
          end_row: row.end_row as number,
          end_col: row.end_col as number
        }));
      const rangeResolved = resolveFromContainingRanges(rangeRows);
      if (rangeResolved) return rangeResolved;

      const columnCandidates = candidateRows.filter((row) => row.scope === "column");
      const tableColumnRows = columnCandidates
        .filter((row) => row.table_id !== null)
        .map((row) => ({ selector_key: row.selector_key, classification: row.classification, scope: "column" }));
      const sheetColumnRows = columnCandidates
        .filter((row) => row.table_id === null)
        .map((row) => ({ selector_key: row.selector_key, classification: row.classification, scope: "column" }));

      if (normalized.tableId) {
        const tableColumnResolved = resolveFromExactRows(tableColumnRows);
        if (tableColumnResolved) return tableColumnResolved;
      }

      const columnResolved = resolveFromExactRows(sheetColumnRows);
      if (columnResolved) return columnResolved;

      const sheetRows = candidateRows
        .filter((row) => row.scope === "sheet")
        .map((row) => ({ selector_key: row.selector_key, classification: row.classification, scope: "sheet" }));
      const sheetResolved = resolveFromExactRows(sheetRows);
      if (sheetResolved) return sheetResolved;

      const docRows = candidateRows
        .filter((row) => row.scope === "document")
        .map((row) => ({ selector_key: row.selector_key, classification: row.classification, scope: "document" }));
      const docResolved = resolveFromExactRows(docRows);
      return docResolved ?? DEFAULT_RESOLUTION;
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
      const candidateRows = await queryCandidatesForRange(
        db,
        docId,
        normalized.sheetId!,
        normalized.startRow!,
        normalized.startCol!,
        normalized.endRow!,
        normalized.endCol!
      );

      const rangeRows = candidateRows
        .filter((row) => row.scope === "range")
        .map((row) => ({
          selector_key: row.selector_key,
          classification: row.classification,
          start_row: row.start_row as number,
          start_col: row.start_col as number,
          end_row: row.end_row as number,
          end_col: row.end_col as number
        }));
      const rangeResolved = resolveFromContainingRanges(rangeRows);
      if (rangeResolved) return rangeResolved;

      if (columnFallback) {
        const columnRows = candidateRows
          .filter((row) => row.scope === "column")
          .map((row) => ({ selector_key: row.selector_key, classification: row.classification, scope: "column" }));
        const columnResolved = resolveFromExactRows(columnRows);
        if (columnResolved) return columnResolved;
      }

      const sheetRows = candidateRows
        .filter((row) => row.scope === "sheet")
        .map((row) => ({ selector_key: row.selector_key, classification: row.classification, scope: "sheet" }));
      const sheetResolved = resolveFromExactRows(sheetRows);
      if (sheetResolved) return sheetResolved;

      const docRows = candidateRows
        .filter((row) => row.scope === "document")
        .map((row) => ({ selector_key: row.selector_key, classification: row.classification, scope: "document" }));
      const docResolved = resolveFromExactRows(docRows);
      return docResolved ?? DEFAULT_RESOLUTION;
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

  const res = await db.query(
    `
      SELECT classification
      FROM document_classifications
      WHERE document_id = $1 AND scope = 'document'
      UNION ALL
      SELECT classification
      FROM document_classifications
      WHERE document_id = $1 AND scope = 'sheet' AND sheet_id = $2
      UNION ALL
      SELECT classification
      FROM document_classifications
      WHERE
        document_id = $1
        AND scope = 'column'
        AND sheet_id = $2
        AND table_id IS NULL
        AND column_index BETWEEN $3 AND $4
      UNION ALL
      SELECT classification
      FROM document_classifications
      WHERE
        document_id = $1
        AND scope = 'range'
        AND sheet_id = $2
        AND start_row <= $5 AND end_row >= $6
        AND start_col <= $4 AND end_col >= $3
      UNION ALL
      SELECT classification
      FROM document_classifications
      WHERE
        document_id = $1
        AND scope = 'cell'
        AND sheet_id = $2
        AND row BETWEEN $6 AND $5
        AND col BETWEEN $3 AND $4
    `,
    [docId, sheetId, range.start.col, range.end.col, range.end.row, range.start.row]
  );

  let result: Classification = { ...DEFAULT_CLASSIFICATION };
  for (const row of res.rows as Array<{ classification: unknown }>) {
    result = maxClassification(result, normalizeClassification(row.classification));
  }
  return result;
}

/**
 * Convenience helper for APIs that need to return selector_key for a selector object.
 * Prefer this over duplicating the shared selectorKey call site.
 */
export function selectorKey(selector: unknown): string {
  return coreSelectorKey(selector as any);
}
