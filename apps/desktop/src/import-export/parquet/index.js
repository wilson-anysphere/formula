import {
  ArrowColumnarSheet,
  arrowTableFromColumns,
  arrowTableToGridBatches,
  arrowTableToParquet,
  parquetFileToArrowTable,
  parquetToArrowTable,
} from "@formula/data-io";
import {
  columnIndexToName,
  normalizeRange,
  parseA1,
  parseRangeA1,
} from "../../document/coords.js";
import { parseImageCellValue } from "../../shared/imageCellValue.js";

// Exporting a rectangular range requires materializing a full columnar representation
// (`Record<string, any[]>`) in JS memory. Keep this bounded so Excel-scale selections
// can't accidentally allocate millions of values.
export const DEFAULT_MAX_PARQUET_EXPORT_CELLS = 200_000;

const PARQUET_IMAGE_MARKER = Symbol("parquetImageValue");

/**
 * DocumentController creates sheets lazily when referenced by `getCell()`.
 *
 * Parquet export is a read-only operation and should never recreate a deleted sheet if callers
 * pass a stale id (e.g. during sheet deletion/undo races). We treat a sheet as "known missing"
 * when the workbook already has *some* sheets, but the requested id is present in neither the
 * materialized `model.sheets` map nor the `sheetMeta` map.
 *
 * When the workbook has no sheets yet, we treat the id as "unknown" and preserve the historical
 * lazy-creation behavior.
 *
 * @param {any} doc
 * @param {string} sheetId
 */
function isSheetKnownMissing(doc, sheetId) {
  const id = String(sheetId ?? "").trim();
  if (!id) return true;

  const sheets = doc?.model?.sheets;
  const sheetMeta = doc?.sheetMeta;
  if (
    sheets &&
    typeof sheets.has === "function" &&
    typeof sheets.size === "number" &&
    sheetMeta &&
    typeof sheetMeta.has === "function" &&
    typeof sheetMeta.size === "number"
  ) {
    const workbookHasAnySheets = sheets.size > 0 || sheetMeta.size > 0;
    if (!workbookHasAnySheets) return false;
    return !sheets.has(id) && !sheetMeta.has(id);
  }
  return false;
}

function wrapImageValueForParquet(image) {
  return {
    [PARQUET_IMAGE_MARKER]: true,
    altText: image?.altText ?? null,
  };
}

function isWrappedImageValue(value) {
  return Boolean(value) && typeof value === "object" && value[PARQUET_IMAGE_MARKER] === true;
}

function cellValueToHeaderText(value) {
  if (value == null) return "";
  if (typeof value === "object" && typeof value.text === "string") return value.text;
  const image = parseImageCellValue(value);
  if (image) return image.altText ?? "";
  return String(value);
}

function coerceCellValueForParquet(value) {
  if (value == null) return null;
  if (typeof value === "object" && typeof value.text === "string") return value.text;
  const image = parseImageCellValue(value);
  if (image) return wrapImageValueForParquet(image);
  return value;
}

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
  const table = await parquetFileToArrowTable(file, {
    batchSize,
    gridBatchSize: batchSize,
    includeHeader: Boolean(options.onBatch),
    onGridBatch: options.onBatch,
  });
  const sheet = new ArrowColumnarSheet(table);

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
  const includeHeader = options.includeHeader ?? true;

  /** @type {import("apache-arrow").Table} */
  let table;

  doc.beginBatch({ label: "Import Parquet" });
  try {
    if (source instanceof Uint8Array) {
      table = await parquetToArrowTable(source, { batchSize });

      for await (const batch of arrowTableToGridBatches(table, {
        batchSize,
        includeHeader,
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
    } else {
      table = await parquetFileToArrowTable(source, {
        batchSize,
        gridBatchSize: batchSize,
        includeHeader,
        onGridBatch: async (batch) => {
          doc.setRangeValues(
            sheetId,
            { row: startCoord.row + batch.rowOffset, col: startCoord.col },
            batch.values
          );
          if (options.onBatch) {
            await options.onBatch(batch);
          }
        },
      });
    }
  } finally {
    doc.endBatch();
  }

  const sheet = new ArrowColumnarSheet(table);

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
 * @param {{ headerRow?: boolean; compression?: any; maxCells?: number }} [options]
 */
export async function exportDocumentRangeToParquet(doc, sheetId, range, options = {}) {
  const r = typeof range === "string" ? parseRangeA1(range) : normalizeRange(range);
  const sheetKey = String(sheetId ?? "").trim();
  if (!sheetKey) {
    throw new Error("Sheet id is required.");
  }
  if (isSheetKnownMissing(doc, sheetKey)) {
    throw new Error(`Unknown sheet: ${sheetKey}`);
  }
  const headerRow = options.headerRow ?? true;
  const maxCells = options.maxCells ?? DEFAULT_MAX_PARQUET_EXPORT_CELLS;

  const rowCount = Math.max(0, r.end.row - r.start.row + 1);
  const colCount = Math.max(0, r.end.col - r.start.col + 1);
  const cellCount = rowCount * colCount;
  if (cellCount > maxCells) {
    throw new Error(
      `Range too large to export to Parquet (${rowCount}x${colCount}=${cellCount} cells). ` +
        `Limit is ${maxCells} cells.`,
    );
  }

  const columnNames = [];
  for (let col = r.start.col; col <= r.end.col; col++) {
    if (headerRow) {
      const cell = doc.getCell(sheetKey, { row: r.start.row, col });
      const value = cell?.value;
      const raw = cellValueToHeaderText(value);
      const name = raw.trim() === "" ? columnIndexToName(col) : raw;
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
      const cell = doc.getCell(sheetKey, { row, col });
      columns[uniqueNames[col - r.start.col]][row - dataStartRow] = coerceCellValueForParquet(cell?.value);
    }
  }

  // Arrow cannot infer a type from arbitrary objects. In-cell image values are stored as JSON-ish
  // objects (from XLSX RichData extraction); normalize them into a scalar value so Parquet export
  // doesn't crash or produce "[object Object]".
  for (const name of uniqueNames) {
    const values = columns[name];
    if (!Array.isArray(values) || values.length === 0) continue;

    let kind = null;
    for (const v of values) {
      if (v == null) continue;
      if (isWrappedImageValue(v)) continue;
      if (typeof v === "number") {
        kind = "number";
        break;
      }
      if (typeof v === "boolean") {
        kind = "boolean";
        break;
      }
      if (v instanceof Date) {
        kind = "date";
        break;
      }
      kind = "string";
      break;
    }

    const imageAsString = kind == null || kind === "string";
    for (let i = 0; i < values.length; i++) {
      const v = values[i];
      if (!isWrappedImageValue(v)) continue;
      values[i] = imageAsString ? v.altText ?? "[Image]" : null;
    }
  }

  const table = arrowTableFromColumns(columns);
  return arrowTableToParquet(table, {
    compression: options.compression ?? "ZSTD",
  });
}
