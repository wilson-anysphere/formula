import type { Query } from "../../../../packages/power-query/src/model.js";
import type { QueryExecutionContext, QueryEngine } from "../../../../packages/power-query/src/engine.js";
import { ArrowTableAdapter } from "../../../../packages/power-query/src/arrowTable.js";
import { DataTable } from "../../../../packages/power-query/src/table.js";

import type { DocumentController } from "../document/documentController.js";
import { dateToExcelSerial } from "../shared/valueParsing.js";

export type QuerySheetDestination = {
  sheetId: string;
  start: { row: number; col: number };
  includeHeader: boolean;
  /**
   * When true, Formula will clear the last output rectangle (if known) before
   * writing the new results. This helps avoid leftover cells when the refreshed
   * result is smaller than the previous refresh.
   */
  clearExisting?: boolean;
  /**
   * Internal bookkeeping to support `clearExisting` on subsequent refreshes.
   * This field is owned by the desktop layer (not the core power-query model).
   */
  lastOutputSize?: { rows: number; cols: number };
};

export type ApplyToDocumentResult = {
  /** Number of rows written (including header row if `destination.includeHeader`). */
  rows: number;
  /** Number of columns written. */
  cols: number;
};

export type ApplyToDocumentProgressEvent =
  | { type: "started"; queryId: string }
  | { type: "engine"; queryId: string; event: unknown }
  | { type: "batch"; queryId: string; rowOffset: number; rowCount: number; totalRowsWritten: number }
  | { type: "completed"; queryId: string; rows: number; cols: number };

export type ApplyToDocumentOptions = {
  engine: QueryEngine;
  context?: QueryExecutionContext;
  batchSize?: number;
  signal?: AbortSignal;
  onProgress?: (event: ApplyToDocumentProgressEvent) => void | Promise<void>;
  label?: string;
};

function abortError(message: string): Error {
  const err = new Error(message);
  // Match the DOM AbortError shape used throughout the codebase.
  (err as any).name = "AbortError";
  return err;
}

function throwIfAborted(signal?: AbortSignal): void {
  if (!signal?.aborted) return;
  throw abortError("Aborted");
}

type GridBatch = { rowOffset: number; values: unknown[][] };

function cellValueToDocumentInput(value: unknown): unknown {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    // Spreadsheet engines expect dates/times as numeric serials (Excel 1900 system).
    // Convert Date objects to serials so they behave like real Excel dates.
    return dateToExcelSerial(value);
  }

  // `DocumentController.setRangeValues` interprets string primitives with:
  // - leading "=" as a formula
  // - leading "'" as a literal escape (and it strips the apostrophe)
  //
  // Power Query output should always be loaded as *values*, even if it looks like a formula.
  // Wrap these strings in `{ value }` to bypass the string parsing path.
  if (typeof value === "string") {
    const trimmed = value.trimStart();
    if ((trimmed.startsWith("=") && trimmed.length > 1) || value.startsWith("'")) {
      return { value };
    }
  }
  return value === undefined ? null : value;
}

function gridToDocumentInputs(grid: unknown[][]): unknown[][] {
  let anyChanged = false;
  const out = grid.map((row) => {
    if (!Array.isArray(row) || row.length === 0) return row;
    let changed = false;
    const mappedRow = row.map((cell) => {
      const mapped = cellValueToDocumentInput(cell);
      if (mapped !== cell) changed = true;
      return mapped;
    });
    if (changed) anyChanged = true;
    return changed ? mappedRow : row;
  });
  return anyChanged ? out : grid;
}

async function* tableToGridBatches(
  table: DataTable | ArrowTableAdapter,
  options: { batchSize: number; includeHeader: boolean },
): AsyncGenerator<GridBatch> {
  if (table instanceof ArrowTableAdapter) {
    const batchSize = options.batchSize;
    const includeHeader = options.includeHeader;
    const baseOffset = includeHeader ? 1 : 0;

    if (includeHeader) {
      yield { rowOffset: 0, values: [table.columns.map((c) => c.name)] };
    }

    try {
      // Prefer the optimized Arrow grid conversion when available.
      const mod = await import("@formula/data-io");
      const arrowTableToGridBatches: any = (mod as any).arrowTableToGridBatches;
      if (typeof arrowTableToGridBatches === "function") {
        for await (const batch of arrowTableToGridBatches(table.table, { batchSize, includeHeader: false })) {
          yield { rowOffset: baseOffset + batch.rowOffset, values: batch.values };
        }
        return;
      }
    } catch {
      // Fall back to row iteration (slower) when Arrow helpers aren't available.
    }

    for (let rowStart = 0; rowStart < table.rowCount; rowStart += batchSize) {
      const end = Math.min(table.rowCount, rowStart + batchSize);
      const slice = [];
      for (let rowIndex = rowStart; rowIndex < end; rowIndex++) {
        slice.push(table.getRow(rowIndex));
      }
      yield { rowOffset: baseOffset + rowStart, values: slice };
    }
    return;
  }

  const batchSize = options.batchSize;
  const includeHeader = options.includeHeader;

  if (includeHeader) {
    yield { rowOffset: 0, values: [table.columns.map((c) => c.name)] };
  }

  const baseOffset = includeHeader ? 1 : 0;
  for (let rowStart = 0; rowStart < table.rowCount; rowStart += batchSize) {
    const end = Math.min(table.rowCount, rowStart + batchSize);
    const slice = [];
    for (let rowIndex = rowStart; rowIndex < end; rowIndex++) {
      slice.push(table.getRow(rowIndex));
    }
    yield { rowOffset: baseOffset + rowStart, values: slice };
  }
}

async function applyBatchesToDocument(
  doc: DocumentController,
  destination: QuerySheetDestination,
  batches: AsyncIterable<GridBatch>,
  options: { signal?: AbortSignal; onProgress?: ApplyToDocumentOptions["onProgress"]; queryId: string },
): Promise<ApplyToDocumentResult> {
  let rowsWritten = 0;
  let colsWritten = 0;

  for await (const batch of batches) {
    throwIfAborted(options.signal);
    const values = gridToDocumentInputs(batch.values);

    doc.setRangeValues(
      destination.sheetId,
      { row: destination.start.row + batch.rowOffset, col: destination.start.col },
      values,
    );

    rowsWritten = Math.max(rowsWritten, batch.rowOffset + batch.values.length);
    for (const row of batch.values) {
      colsWritten = Math.max(colsWritten, Array.isArray(row) ? row.length : 0);
    }

    await options.onProgress?.({
      type: "batch",
      queryId: options.queryId,
      rowOffset: batch.rowOffset,
      rowCount: batch.values.length,
      totalRowsWritten: rowsWritten,
    });
  }

  return { rows: rowsWritten, cols: colsWritten };
}

function clearExistingOutput(doc: DocumentController, destination: QuerySheetDestination): void {
  const size = destination.lastOutputSize;
  if (!destination.clearExisting || !size) return;
  if (size.rows <= 0 || size.cols <= 0) return;

  doc.clearRange(destination.sheetId, {
    start: destination.start,
    end: { row: destination.start.row + size.rows - 1, col: destination.start.col + size.cols - 1 },
  });
}

/**
 * Apply a Power Query `Query` to a sheet destination.
 *
 * This executes the query via `QueryEngine.executeQueryStreaming` and progressively writes
 * grid batches into the `DocumentController`.
 */
export async function applyQueryToDocument(
  doc: DocumentController,
  query: Query,
  destination: QuerySheetDestination,
  options: ApplyToDocumentOptions,
): Promise<ApplyToDocumentResult> {
  const batchSize = options.batchSize ?? 1024;
  const initialDepth = (doc as any).batchDepth ?? 0;

  doc.beginBatch({ label: options.label ?? `Load query: ${query.name}` });

  try {
    throwIfAborted(options.signal);
    await options.onProgress?.({ type: "started", queryId: query.id });

    clearExistingOutput(doc, destination);

    let rowsWritten = 0;
    let colsWritten = 0;
    await options.engine.executeQueryStreaming(query, options.context ?? {}, {
      batchSize,
      includeHeader: destination.includeHeader,
      materialize: false,
      signal: options.signal,
      onProgress: async (evt) => {
        await options.onProgress?.({ type: "engine", queryId: query.id, event: evt });
      },
      onBatch: async (batch) => {
        // `executeQueryStreaming` doesn't currently consult the signal between batches.
        // Make cancellation responsive by explicitly checking before writing each batch.
        throwIfAborted(options.signal);

        const values = gridToDocumentInputs(batch.values);
        doc.setRangeValues(
          destination.sheetId,
          { row: destination.start.row + batch.rowOffset, col: destination.start.col },
          values,
        );

        rowsWritten = Math.max(rowsWritten, batch.rowOffset + batch.values.length);
        for (const row of batch.values) {
          colsWritten = Math.max(colsWritten, Array.isArray(row) ? row.length : 0);
        }

        await options.onProgress?.({
          type: "batch",
          queryId: query.id,
          rowOffset: batch.rowOffset,
          rowCount: batch.values.length,
          totalRowsWritten: rowsWritten,
        });
      },
    });

    destination.lastOutputSize = { rows: rowsWritten, cols: colsWritten };
    await options.onProgress?.({ type: "completed", queryId: query.id, rows: rowsWritten, cols: colsWritten });

    doc.endBatch();
    return { rows: rowsWritten, cols: colsWritten };
  } catch (err) {
    if (initialDepth === 0) {
      // Revert partial writes (including any cleared cells) when we're not nested inside another batch.
      doc.cancelBatch();
    } else {
      // Always unwind our batch depth so callers aren't left in a broken state.
      doc.endBatch();
    }
    throw err;
  }
}

export type ApplyTableToDocumentOptions = {
  batchSize?: number;
  includeHeader?: boolean;
  signal?: AbortSignal;
  onProgress?: (event: ApplyToDocumentProgressEvent) => void | Promise<void>;
  label?: string;
  queryId?: string;
};

export async function applyTableToDocument(
  doc: DocumentController,
  table: DataTable | ArrowTableAdapter,
  destination: QuerySheetDestination,
  options: ApplyTableToDocumentOptions = {},
): Promise<ApplyToDocumentResult> {
  const batchSize = options.batchSize ?? 1024;
  const includeHeader = options.includeHeader ?? destination.includeHeader;
  const initialDepth = (doc as any).batchDepth ?? 0;
  const queryId = options.queryId ?? "<table>";

  doc.beginBatch({ label: options.label ?? "Load query results" });

  try {
    throwIfAborted(options.signal);
    clearExistingOutput(doc, destination);
    await options.onProgress?.({ type: "started", queryId });

    const result = await applyBatchesToDocument(doc, destination, tableToGridBatches(table, { batchSize, includeHeader }), {
      signal: options.signal,
      onProgress: options.onProgress,
      queryId,
    });

    destination.lastOutputSize = { rows: result.rows, cols: result.cols };
    await options.onProgress?.({ type: "completed", queryId, rows: result.rows, cols: result.cols });

    doc.endBatch();
    return result;
  } catch (err) {
    if (initialDepth === 0) {
      doc.cancelBatch();
    } else {
      doc.endBatch();
    }
    throw err;
  }
}
