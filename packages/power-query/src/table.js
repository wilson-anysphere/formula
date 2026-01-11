/**
 * A minimal, in-memory columnar table abstraction used by the Power Query-style
 * transformation engine.
 *
 * Rows are stored as arrays for performance and predictable column ordering.
 *
 * This intentionally avoids any spreadsheet-specific concepts (A1 addresses,
 * formatting, etc.) so it can run in both the UI preview and a backend worker.
 */

/**
 * @typedef {"any" | "string" | "number" | "boolean" | "date"} DataType
 */

/**
 * @typedef {{ name: string, type: DataType }} Column
 */

/**
 * A minimal column-vector interface shared by row-backed and Arrow-backed tables.
 *
 * @typedef {{
 *   length: number;
 *   get: (index: number) => unknown;
 * }} ColumnVector
 */

/**
 * Shared table interface used by the transformation engine. Implemented by `DataTable`
 * and `ArrowTableAdapter`.
 *
 * @typedef {{
 *   columns: Column[];
 *   readonly rowCount: number;
 *   readonly columnCount: number;
 *   getColumnIndex: (name: string) => number;
 *   getColumnVector: (index: number) => ColumnVector;
 *   getCell: (rowIndex: number, colIndex: number) => unknown;
 *   getRow: (rowIndex: number) => unknown[];
 *   iterRows: () => IterableIterator<unknown[]>;
 *   toGrid: (options?: { includeHeader?: boolean }) => unknown[][];
 *   head: (limit: number) => any;
 * }} ITable
 */

/**
 * Convert a user-provided value into a printable column name.
 * @param {unknown} value
 * @returns {string}
 */
function normalizeColumnName(value) {
  const text = value == null ? "" : String(value).trim();
  return text;
}

/**
 * Power Query style column name uniquing: `A`, `A.1`, `A.2`, ...
 * Empty names become `Column1`, `Column2`, ...
 * @param {unknown[]} rawNames
 * @returns {string[]}
 */
export function makeUniqueColumnNames(rawNames) {
  const baseNames = rawNames.map((name, idx) => {
    const normalized = normalizeColumnName(name);
    return normalized === "" ? `Column${idx + 1}` : normalized;
  });

  const counts = new Map();
  return baseNames.map((base) => {
    const seen = counts.get(base) ?? 0;
    counts.set(base, seen + 1);
    return seen === 0 ? base : `${base}.${seen}`;
  });
}

/**
 * @param {unknown} value
 * @returns {value is Date}
 */
function isDate(value) {
  return value instanceof Date && !Number.isNaN(value.getTime());
}

/**
 * @param {unknown[]} values
 * @returns {DataType}
 */
export function inferColumnType(values) {
  let sawNumber = false;
  let sawBoolean = false;
  let sawDate = false;
  let sawOther = false;

  for (const value of values) {
    if (value == null) continue;
    if (typeof value === "number" && Number.isFinite(value)) {
      sawNumber = true;
      continue;
    }
    if (typeof value === "boolean") {
      sawBoolean = true;
      continue;
    }
    if (isDate(value)) {
      sawDate = true;
      continue;
    }
    sawOther = true;
  }

  const kinds = [sawNumber, sawBoolean, sawDate, sawOther].filter(Boolean).length;
  if (kinds === 0) return "any";
  if (kinds > 1) return "any";
  if (sawNumber) return "number";
  if (sawBoolean) return "boolean";
  if (sawDate) return "date";
  return "string";
}

/**
 * @param {unknown} value
 * @returns {unknown}
 */
function normalizeCellValue(value) {
  return value === undefined ? null : value;
}

export class DataTable {
  /**
   * @param {Column[]} columns
   * @param {unknown[][]} rows
   */
  constructor(columns, rows) {
    /** @type {Column[]} */
    this.columns = columns.map((c) => ({ name: c.name, type: c.type ?? "any" }));

    /** @type {Map<string, number>} */
    this.columnIndex = new Map();
    this.columns.forEach((col, idx) => {
      if (this.columnIndex.has(col.name)) {
        throw new Error(`Duplicate column name '${col.name}'`);
      }
      this.columnIndex.set(col.name, idx);
    });

    const width = this.columns.length;
    /** @type {unknown[][]} */
    this.rows = rows.map((row) => {
      const normalized = Array.from({ length: width }, (_, i) =>
        normalizeCellValue(row?.[i]),
      );
      return normalized;
    });
  }

  /**
   * Create a table from a 2D grid.
   * @param {unknown[][]} grid
   * @param {{ hasHeaders?: boolean, inferTypes?: boolean }} [options]
   * @returns {DataTable}
   */
  static fromGrid(grid, options = {}) {
    const hasHeaders = options.hasHeaders ?? true;
    const inferTypes = options.inferTypes ?? true;

    if (!Array.isArray(grid) || grid.length === 0) {
      return new DataTable([], []);
    }

    const headerRow = hasHeaders ? grid[0] : null;
    const dataRows = hasHeaders ? grid.slice(1) : grid;
    // Avoid `Math.max(...rows.map(...))` because spreading large arrays can exceed the
    // VM argument limit / call stack on big datasets.
    let width = 0;
    for (const row of grid) {
      if (!Array.isArray(row)) continue;
      if (row.length > width) width = row.length;
    }

    const rawNames = headerRow ?? Array.from({ length: width }, (_, i) => `Column${i + 1}`);
    const names = makeUniqueColumnNames(rawNames);

    const columns = names.map((name) => ({ name, type: "any" }));

    const normalizedRows = dataRows.map((row) =>
      Array.from({ length: width }, (_, i) => normalizeCellValue(row?.[i])),
    );

    if (!inferTypes) {
      return new DataTable(columns, normalizedRows);
    }

    const typedColumns = columns.map((col, idx) => ({
      name: col.name,
      type: inferColumnType(normalizedRows.map((row) => row[idx])),
    }));

    return new DataTable(typedColumns, normalizedRows);
  }

  get rowCount() {
    return this.rows.length;
  }

  get columnCount() {
    return this.columns.length;
  }

  /**
   * @param {string} name
   * @returns {number}
   */
  getColumnIndex(name) {
    const idx = this.columnIndex.get(name);
    if (idx == null) {
      throw new Error(`Unknown column '${name}'. Available: ${this.columns.map((c) => c.name).join(", ")}`);
    }
    return idx;
  }

  /**
   * @param {number} index
   * @returns {ColumnVector}
   */
  getColumnVector(index) {
    if (index < 0 || index >= this.columns.length) {
      throw new Error(`Unknown column index ${index}`);
    }

    const rows = this.rows;
    return {
      length: rows.length,
      get: (rowIndex) => rows[rowIndex]?.[index] ?? null,
    };
  }

  /**
   * @param {number} rowIndex
   * @param {number} colIndex
   * @returns {unknown}
   */
  getCell(rowIndex, colIndex) {
    return this.rows[rowIndex]?.[colIndex] ?? null;
  }

  /**
   * @param {number} rowIndex
   * @returns {unknown[]}
   */
  getRow(rowIndex) {
    return this.rows[rowIndex] ?? [];
  }

  /**
   * @returns {IterableIterator<unknown[]>}
   */
  *iterRows() {
    yield* this.rows;
  }

  /**
   * @param {{ includeHeader?: boolean }} [options]
   * @returns {unknown[][]}
   */
  toGrid(options = {}) {
    const includeHeader = options.includeHeader ?? true;
    const header = this.columns.map((c) => c.name);
    if (!includeHeader) return this.rows.map((row) => row.slice());
    return [header, ...this.rows.map((row) => row.slice())];
  }

  /**
   * @param {number} limit
   * @returns {DataTable}
   */
  head(limit) {
    return new DataTable(this.columns, this.rows.slice(0, limit));
  }
}
