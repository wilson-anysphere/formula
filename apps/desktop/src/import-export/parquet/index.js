import {
  ArrowColumnarSheet,
  arrowTableFromColumns,
  arrowTableToGridBatches,
  arrowTableToParquet,
  parquetToArrowTable,
} from "@formula/data-io";
import {
  columnIndexToName,
  normalizeRange,
  parseA1,
  parseRangeA1,
} from "../../document/coords.js";

/**
 * Import a Parquet file into a columnar sheet backing store.
 *
 * If `onBatch` is provided, this will also emit row batches (including a header row) so a
 * cell-based grid can be populated progressively.
 *
 * @param {File} file
 * @param {{ batchSize?: number; onBatch?: (batch: {rowOffset: number; values: any[][]}) => Promise<void> | void }} [options]
 */
export async function importParquetFile(file, options = {}) {
  const batchSize = options.batchSize ?? 1024;
  const parquetBytes = new Uint8Array(await file.arrayBuffer());

  const table = await parquetToArrowTable(parquetBytes, { batchSize });
  const sheet = new ArrowColumnarSheet(table);

  if (options.onBatch) {
    for await (const batch of arrowTableToGridBatches(table, {
      batchSize,
      includeHeader: true,
    })) {
      await options.onBatch(batch);
    }
  }

  return sheet;
}

/**
 * Convenience helper to import Parquet directly into a document at a start cell.
 *
 * Values are inserted in batches via `DocumentController.setRangeValues`, wrapped in
 * `beginBatch/endBatch` so the import is a single undo step.
 *
 * @param {import("../../document/documentController.js").DocumentController} doc
 * @param {string} sheetId
 * @param {import("../../document/coords.js").CellCoord | string} start
 * @param {File | Uint8Array} source
 * @param {{ batchSize?: number; includeHeader?: boolean; onBatch?: (batch: {rowOffset: number; values: any[][]}) => Promise<void> | void }} [options]
 */
export async function importParquetIntoDocument(doc, sheetId, start, source, options = {}) {
  const batchSize = options.batchSize ?? 1024;
  const startCoord = typeof start === "string" ? parseA1(start) : start;
  const parquetBytes = source instanceof Uint8Array ? source : new Uint8Array(await source.arrayBuffer());

  const table = await parquetToArrowTable(parquetBytes, { batchSize });
  const sheet = new ArrowColumnarSheet(table);

  doc.beginBatch({ label: "Import Parquet" });
  try {
    for await (const batch of arrowTableToGridBatches(table, {
      batchSize,
      includeHeader: options.includeHeader ?? true,
    })) {
      doc.setRangeValues(
        sheetId,
        { row: startCoord.row + batch.rowOffset, col: startCoord.col },
        batch.values
      );
      if (options.onBatch) {
        await options.onBatch(batch);
      }
    }
  } finally {
    doc.endBatch();
  }

  return sheet;
}

/**
 * Export a columnar sheet (or range) to Parquet bytes.
 *
 * @param {ArrowColumnarSheet} sheet
 * @param {{ range?: { startRow: number; endRow: number; startCol: number; endCol: number }; compression?: any }} [options]
 */
export async function exportParquetFromSheet(sheet, options = {}) {
  const selection = options.range ? sheet.slice(options.range) : sheet;
  return arrowTableToParquet(selection.table, {
    compression: options.compression ?? "ZSTD",
  });
}

function makeUniqueColumnNames(names) {
  /** @type {Map<string, number>} */
  const seen = new Map();
  return names.map((raw) => {
    const base = String(raw ?? "").trim() || "Column";
    const count = seen.get(base) ?? 0;
    seen.set(base, count + 1);
    if (count === 0) return base;
    return `${base}_${count + 1}`;
  });
}

/**
 * Export a document range to a Parquet file.
 *
 * By default the first row of the range becomes the Parquet schema (column names) and is excluded
 * from data rows.
 *
 * @param {import("../../document/documentController.js").DocumentController} doc
 * @param {string} sheetId
 * @param {import("../../document/coords.js").CellRange | string} range
 * @param {{ headerRow?: boolean; compression?: any }} [options]
 */
export async function exportDocumentRangeToParquet(doc, sheetId, range, options = {}) {
  const r = typeof range === "string" ? parseRangeA1(range) : normalizeRange(range);
  const headerRow = options.headerRow ?? true;

  const columnNames = [];
  for (let col = r.start.col; col <= r.end.col; col++) {
    if (headerRow) {
      const cell = doc.getCell(sheetId, { row: r.start.row, col });
      const value = cell?.value;
      const name = value == null || String(value).trim() === "" ? columnIndexToName(col) : String(value);
      columnNames.push(name);
    } else {
      columnNames.push(columnIndexToName(col));
    }
  }

  const uniqueNames = makeUniqueColumnNames(columnNames);
  const dataStartRow = headerRow ? r.start.row + 1 : r.start.row;
  const dataRowCount = Math.max(0, r.end.row - dataStartRow + 1);

  /** @type {Record<string, any[]>} */
  const columns = {};

  for (let i = 0; i < uniqueNames.length; i++) {
    columns[uniqueNames[i]] = new Array(dataRowCount);
  }

  for (let row = dataStartRow; row <= r.end.row; row++) {
    for (let col = r.start.col; col <= r.end.col; col++) {
      const cell = doc.getCell(sheetId, { row, col });
      columns[uniqueNames[col - r.start.col]][row - dataStartRow] = cell?.value ?? null;
    }
  }

  const table = arrowTableFromColumns(columns);
  return arrowTableToParquet(table, {
    compression: options.compression ?? "ZSTD",
  });
}
