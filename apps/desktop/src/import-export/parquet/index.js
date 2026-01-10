import {
  ArrowColumnarSheet,
  arrowTableToGridBatches,
  arrowTableToParquet,
  parquetToArrowTable,
} from "@formula/data-io";
import { parseA1 } from "../../document/coords.js";

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
