import * as arrow from 'apache-arrow';
let parquetWasmLoadPromise;

async function getParquetWasm() {
  parquetWasmLoadPromise ??= (async () => {
    const parquet = await import('parquet-wasm/esm');

    // Browser: the wasm file will be resolved relative to the module URL.
    const isNode = typeof process !== 'undefined' && Boolean(process.versions?.node);
    if (!isNode) {
      await parquet.default();
      return parquet;
    }

    // Node: avoid `fetch(file://...)` by loading the wasm file via `fs`.
    const [{ createRequire }, { readFile }] = await Promise.all([
      import('node:module'),
      import('node:fs/promises'),
    ]);

    const require = createRequire(import.meta.url);
    const wasmPath = require.resolve('parquet-wasm/esm/parquet_wasm_bg.wasm');
    const wasmBytes = await readFile(wasmPath);

    await parquet.default({ module_or_path: wasmBytes });
    return parquet;
  })();

  return parquetWasmLoadPromise;
}

function uint8ArrayToBase64(bytes) {
  if (typeof Buffer !== 'undefined') {
    return Buffer.from(bytes).toString('base64');
  }

  let binary = '';
  for (let i = 0; i < bytes.length; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  // eslint-disable-next-line no-undef
  return btoa(binary);
}

function arrowValueToCellValue(value, type) {
  if (value === null || value === undefined) return null;

  if (type?.typeId === arrow.Type.Date || type?.typeId === arrow.Type.Timestamp) {
    if (value instanceof Date) return value;
    const raw =
      typeof value === 'number' && Number.isFinite(value)
        ? value
        : typeof value === 'bigint' &&
            value <= Number.MAX_SAFE_INTEGER &&
            value >= Number.MIN_SAFE_INTEGER
          ? Number(value)
          : null;
    if (raw != null) {
      let ms = raw;
      if (type.typeId === arrow.Type.Date) {
        ms = type.unit === arrow.DateUnit.DAY ? raw * 86400000 : raw;
      } else {
        switch (type.unit) {
          case arrow.TimeUnit.SECOND:
            ms = raw * 1000;
            break;
          case arrow.TimeUnit.MILLISECOND:
            ms = raw;
            break;
          case arrow.TimeUnit.MICROSECOND:
            ms = raw / 1000;
            break;
          case arrow.TimeUnit.NANOSECOND:
            ms = raw / 1_000_000;
            break;
          default:
            ms = raw;
            break;
        }
      }
      const date = new Date(ms);
      if (!Number.isNaN(date.getTime())) return date;
    }
  }

  if (typeof value === 'bigint') {
    return value <= Number.MAX_SAFE_INTEGER && value >= Number.MIN_SAFE_INTEGER
      ? Number(value)
      : value.toString();
  }
  if (value instanceof Date) return value;
  if (value instanceof Uint8Array) return uint8ArrayToBase64(value);
  return value;
}

/**
 * Read Parquet bytes into an Arrow JS Table.
 *
 * @param {Uint8Array} parquetBytes
 * @param {import('parquet-wasm').ReaderOptions} [options]
 */
export async function parquetToArrowTable(parquetBytes, options) {
  const parquet = await getParquetWasm();
  const wasmTable = parquet.readParquet(parquetBytes, options ?? null);
  return arrow.tableFromIPC(wasmTable.intoIPCStream());
}

/**
 * Read a Parquet File/Blob into an Arrow JS Table without first materializing the entire file into
 * a single ArrayBuffer.
 *
 * Optionally emits grid batches as data is decoded so callers can progressively populate a
 * cell-based UI while still building the final Arrow table.
 *
 * @param {Blob} handle
 * @param {import('parquet-wasm').ReaderOptions & {
 *   gridBatchSize?: number;
 *   includeHeader?: boolean;
 *   onGridBatch?: (batch: {rowOffset: number; values: any[][]}) => Promise<void> | void;
 * }} [options]
 */
export async function parquetFileToArrowTable(handle, options = {}) {
  const parquet = await getParquetWasm();
  const parquetFile = await parquet.ParquetFile.fromFile(handle);

  const {
    gridBatchSize = 1024,
    includeHeader = true,
    onGridBatch,
    ...readerOptions
  } = options ?? {};

  try {
    let globalRowOffset = 0;
    let emittedHeader = false;

    if (onGridBatch && includeHeader) {
      const wasmSchema = parquetFile.schema();
      const schemaTable = arrow.tableFromIPC(wasmSchema.intoIPCStream());
      const columnNames = schemaTable.schema.fields.map((field) => field.name);
      await onGridBatch({ rowOffset: 0, values: [columnNames] });
      emittedHeader = true;
      globalRowOffset = 1;
    }

    const stream = await parquetFile.stream(readerOptions ?? null);

    /** @type {import('apache-arrow').RecordBatch[]} */
    const recordBatches = [];

    for await (const wasmRecordBatch of stream) {
      const table = arrow.tableFromIPC(wasmRecordBatch.intoIPCStream());
      recordBatches.push(...table.batches);

      if (onGridBatch) {
        if (!emittedHeader && includeHeader) {
          const columnNames = table.schema.fields.map((field) => field.name);
          await onGridBatch({ rowOffset: 0, values: [columnNames] });
          emittedHeader = true;
          globalRowOffset = 1;
        }

        for await (const batch of arrowTableToGridBatches(table, {
          batchSize: gridBatchSize,
          includeHeader: false,
        })) {
          await onGridBatch({ rowOffset: globalRowOffset + batch.rowOffset, values: batch.values });
        }
        globalRowOffset += table.numRows;
      }
    }

    return new arrow.Table(recordBatches);
  } finally {
    parquetFile.free();
  }
}

/**
 * Write an Arrow JS Table into Parquet bytes.
 *
 * @param {arrow.Table} table
 * @param {{ compression?: keyof typeof Compression | Compression | null }} [options]
 */
export async function arrowTableToParquet(table, options = {}) {
  const parquet = await getParquetWasm();
  const wasmTable = parquet.Table.fromIPCStream(arrow.tableToIPC(table, 'stream'));

  /** @type {import('parquet-wasm').WriterProperties | null} */
  let writerProperties = null;
  let builder = new parquet.WriterPropertiesBuilder();

  try {
    const compression = options.compression ?? null;
    if (compression !== null && compression !== undefined) {
      const codec =
        typeof compression === 'string'
          ? parquet.Compression[compression]
          : compression;
      if (codec !== undefined) {
        builder = builder.setCompression(codec);
      }
    }

    writerProperties = builder.build();
    return parquet.writeParquet(wasmTable, writerProperties);
  } finally {
    // `WriterPropertiesBuilder.build()` consumes the builder, and `writeParquet` consumes both the
    // table and the writer properties, so there is nothing left to free here.
  }
}

/**
 * Serialize an Arrow JS Table into an Arrow IPC stream.
 *
 * This is useful for caching/interchange when you want to preserve Arrow's full
 * type information (timestamps, decimals, dictionary encoding, etc) without
 * going through Parquet.
 *
 * @param {arrow.Table} table
 * @returns {Uint8Array}
 */
export function arrowTableToIPC(table) {
  return arrow.tableToIPC(table, 'stream');
}

/**
 * Deserialize an Arrow IPC stream into an Arrow JS Table.
 *
 * @param {Uint8Array | ArrayBuffer} bytes
 * @returns {arrow.Table}
 */
export function arrowTableFromIPC(bytes) {
  return arrow.tableFromIPC(bytes);
}

/**
 * Construct an Arrow JS Table from column arrays.
 *
 * This is a small wrapper around Arrow JS's `tableFromArrays` to avoid consumers needing to depend
 * on `apache-arrow` directly when they already depend on `@formula/data-io`.
 *
 * @param {Record<string, any[] | ArrayLike<any>>} columns
 */
export function arrowTableFromColumns(columns) {
  return arrow.tableFromArrays(columns);
}

/**
 * Yield a 2D grid representation of an Arrow Table in batches suitable for progressive insertion
 * into a cell-based spreadsheet model.
 *
 * `rowOffset` is a grid row index where row 0 is the header.
 *
 * @param {arrow.Table} table
 * @param {{ batchSize?: number; includeHeader?: boolean }} [options]
 */
export async function* arrowTableToGridBatches(
  table,
  { batchSize = 1024, includeHeader = true } = {}
) {
  const columnNames = table.schema.fields.map((field) => field.name);
  const columnCount = columnNames.length;
  const dataBaseRowOffset = includeHeader ? 1 : 0;

  if (includeHeader) {
    yield { rowOffset: 0, values: [columnNames] };
  }

  let dataRowOffset = 0;
  for (const recordBatch of table.batches) {
    for (let batchStart = 0; batchStart < recordBatch.numRows; batchStart += batchSize) {
      const batchEnd = Math.min(recordBatch.numRows, batchStart + batchSize);
      const rows = new Array(batchEnd - batchStart);

      for (let rowIndex = batchStart; rowIndex < batchEnd; rowIndex++) {
        const row = new Array(columnCount);
        for (let colIndex = 0; colIndex < columnCount; colIndex++) {
          const vector = recordBatch.getChildAt(colIndex);
          row[colIndex] = arrowValueToCellValue(vector.get(rowIndex), vector.type);
        }
        rows[rowIndex - batchStart] = row;
      }

      yield { rowOffset: dataBaseRowOffset + dataRowOffset + batchStart, values: rows };
    }

    dataRowOffset += recordBatch.numRows;
  }
}

/**
 * A lightweight columnar backing store for a grid, backed by an Arrow Table.
 */
export class ArrowColumnarSheet {
  /**
   * @param {arrow.Table} table
   */
  constructor(table) {
    this.table = table;
    this.columnNames = table.schema.fields.map((field) => field.name);
  }

  get rowCount() {
    return this.table.numRows + 1;
  }

  get columnCount() {
    return this.columnNames.length;
  }

  /**
   * Grid coordinates, where row 0 is the header row.
   *
   * @param {number} row
   * @param {number} col
   */
  getCell(row, col) {
    if (row === 0) return this.columnNames[col] ?? null;
    const vector = this.table.getChildAt(col);
    return arrowValueToCellValue(vector?.get(row - 1), vector?.type);
  }

  /**
   * Select a sub-range of the Arrow table for export.
   *
   * @param {{ startRow: number; endRow: number; startCol: number; endCol: number }} range
   */
  slice(range) {
    const rowStart = Math.max(0, range.startRow - 1);
    const rowEndExclusive = Math.max(rowStart, range.endRow);
    const colIndices = [];
    for (let col = range.startCol; col <= range.endCol; col++) {
      colIndices.push(col);
    }

    const sliced = this.table.slice(rowStart, rowEndExclusive).selectAt(colIndices);
    return new ArrowColumnarSheet(sliced);
  }
}
