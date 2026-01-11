/**
 * @typedef {import("./table.js").Column} Column
 * @typedef {import("./table.js").DataType} DataType
 * @typedef {import("./table.js").ColumnVector} ColumnVector
 */

import { inferColumnType } from "./table.js";

/**
 * @param {unknown} value
 * @returns {value is Date}
 */
function isDate(value) {
  return value instanceof Date && !Number.isNaN(value.getTime());
}

/**
 * Normalize Arrow values into the primitive cell types we use elsewhere in Power Query.
 * This mostly matches `@formula/data-io`'s Arrow grid conversion logic.
 *
 * @param {unknown} value
 * @returns {unknown}
 */
/**
 * @param {string | undefined} typeHint
 * @returns {string | null}
 */
function parseArrowTypeParam(typeHint) {
  if (typeof typeHint !== "string") return null;
  const start = typeHint.indexOf("<");
  const end = typeHint.indexOf(">");
  if (start < 0 || end < 0 || end <= start) return null;
  const inside = typeHint.slice(start + 1, end);
  const [first] = inside.split(",", 1);
  return first?.trim() ?? null;
}

/**
 * @param {unknown} value
 * @param {string | undefined} typeHint
 * @returns {Date | null}
 */
function arrowTemporalValueToDate(value, typeHint) {
  if (value == null) return null;
  if (isDate(value)) return value;

  const raw =
    typeof value === "number" && Number.isFinite(value)
      ? value
      : typeof value === "bigint" && value <= Number.MAX_SAFE_INTEGER && value >= Number.MIN_SAFE_INTEGER
        ? Number(value)
        : null;
  if (raw == null) return null;

  if (typeHint?.startsWith("Date32")) {
    const d = new Date(raw * 86400000);
    return Number.isNaN(d.getTime()) ? null : d;
  }

  if (typeHint?.startsWith("Date64")) {
    const d = new Date(raw);
    return Number.isNaN(d.getTime()) ? null : d;
  }

  if (typeHint?.startsWith("Timestamp")) {
    const unit = parseArrowTypeParam(typeHint) ?? "MILLISECOND";
    let ms = raw;
    switch (unit) {
      case "SECOND":
        ms = raw * 1000;
        break;
      case "MILLISECOND":
        ms = raw;
        break;
      case "MICROSECOND":
        ms = raw / 1000;
        break;
      case "NANOSECOND":
        ms = raw / 1_000_000;
        break;
      default:
        ms = raw;
        break;
    }
    const d = new Date(ms);
    return Number.isNaN(d.getTime()) ? null : d;
  }

  return null;
}

/**
 * @param {unknown} value
 * @param {string | undefined} [typeHint]
 * @returns {unknown}
 */
function arrowValueToCellValue(value, typeHint) {
  if (value === null || value === undefined) return null;

  if (typeof typeHint === "string" && (typeHint.startsWith("Timestamp") || typeHint.startsWith("Date32") || typeHint.startsWith("Date64"))) {
    const maybeDate = arrowTemporalValueToDate(value, typeHint);
    if (maybeDate) return maybeDate;
  }

  if (typeof value === "bigint") {
    return value <= Number.MAX_SAFE_INTEGER && value >= Number.MIN_SAFE_INTEGER ? Number(value) : value.toString();
  }
  if (value instanceof Uint8Array) {
    // Small helper: render binary as base64 so it is printable in a grid.
    if (typeof Buffer !== "undefined") {
      return Buffer.from(value).toString("base64");
    }
    let binary = "";
    for (let i = 0; i < value.length; i++) {
      binary += String.fromCharCode(value[i]);
    }
    // eslint-disable-next-line no-undef
    return btoa(binary);
  }
  if (isDate(value)) return value;
  return value;
}

/**
 * @param {string} typeHint
 * @returns {DataType}
 */
function dataTypeFromArrowType(typeHint) {
  if (typeHint.startsWith("Bool")) return "boolean";
  if (typeHint.includes("Utf8")) return "string";
  if (typeHint.includes("Binary")) return "string";
  if (typeHint.startsWith("Timestamp") || typeHint.startsWith("Date32") || typeHint.startsWith("Date64")) return "date";
  if (/^(?:Int|Uint|Float|Decimal)/.test(typeHint)) return "number";
  return "any";
}

/**
 * @param {import("apache-arrow").Table} table
 * @returns {Column[]}
 */
function columnsFromArrow(table) {
  const fields = table.schema?.fields ?? [];
  return fields.map((field, idx) => {
    const typeHint = String(field.type);
    const mapped = dataTypeFromArrowType(typeHint);
    if (mapped !== "any") {
      return { name: field.name, type: mapped };
    }

    const vector = table.getChildAt?.(idx);
    const sampleSize = Math.min(vector?.length ?? 0, 64);
    const sample = new Array(sampleSize);
    for (let i = 0; i < sampleSize; i++) {
      sample[i] = arrowValueToCellValue(vector?.get(i), typeHint);
    }
    return { name: field.name, type: inferColumnType(sample) };
  });
}

/**
 * A thin adapter that presents an Arrow JS `Table` through the Power Query `ITable` interface.
 *
 * The adapter does not eagerly materialize row arrays; callers should prefer column-vector access.
 */
export class ArrowTableAdapter {
  /**
   * @param {import("apache-arrow").Table} table
   * @param {Column[]} [columns]
   */
  constructor(table, columns) {
    this.table = table;
    this.arrowTypes = (table.schema?.fields ?? []).map((field) => String(field.type));

    const fieldCount = table.schema?.fields?.length;
    if (Array.isArray(columns) && typeof fieldCount === "number" && columns.length !== fieldCount) {
      throw new Error(`ArrowTableAdapter column metadata length mismatch (expected ${fieldCount}, got ${columns.length})`);
    }

    /** @type {Column[]} */
    this.columns = (columns ?? columnsFromArrow(table)).map((c) => ({ name: c.name, type: c.type ?? "any" }));

    /** @type {Map<string, number>} */
    this.columnIndex = new Map();
    this.columns.forEach((col, idx) => {
      if (this.columnIndex.has(col.name)) {
        throw new Error(`Duplicate column name '${col.name}'`);
      }
      this.columnIndex.set(col.name, idx);
    });
  }

  get rowCount() {
    return this.table.numRows;
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
    const vector = this.table.getChildAt(index);
    if (!vector) {
      throw new Error(`Unknown column index ${index}`);
    }
    return {
      length: this.rowCount,
      get: (rowIndex) => arrowValueToCellValue(vector.get(rowIndex), this.arrowTypes[index]),
    };
  }

  /**
   * @param {number} rowIndex
   * @param {number} colIndex
   * @returns {unknown}
   */
  getCell(rowIndex, colIndex) {
    if (colIndex < 0 || colIndex >= this.columnCount) return null;
    const vector = this.table.getChildAt(colIndex);
    return arrowValueToCellValue(vector?.get(rowIndex), this.arrowTypes[colIndex]);
  }

  /**
   * @param {number} rowIndex
   * @returns {unknown[]}
   */
  getRow(rowIndex) {
    const row = new Array(this.columnCount);
    for (let c = 0; c < this.columnCount; c++) {
      row[c] = this.getCell(rowIndex, c);
    }
    return row;
  }

  /**
   * Iterate row arrays when a row-oriented algorithm is unavoidable.
   * @returns {IterableIterator<unknown[]>}
   */
  *iterRows() {
    for (let i = 0; i < this.rowCount; i++) {
      yield this.getRow(i);
    }
  }

  /**
   * @param {{ includeHeader?: boolean }} [options]
   * @returns {unknown[][]}
   */
  toGrid(options = {}) {
    const includeHeader = options.includeHeader ?? true;
    const out = [];
    if (includeHeader) {
      out.push(this.columns.map((c) => c.name));
    }
    for (let i = 0; i < this.rowCount; i++) {
      out.push(this.getRow(i));
    }
    return out;
  }

  /**
   * @param {number} limit
   * @returns {ArrowTableAdapter}
   */
  head(limit) {
    return new ArrowTableAdapter(this.table.slice(0, limit), this.columns);
  }
}
