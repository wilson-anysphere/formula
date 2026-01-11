import { DataTable } from "../table.js";

/**
 * @typedef {import("../table.js").Column} Column
 */

const TYPE_KEY = "__pq_type";

/**
 * @param {unknown} value
 * @returns {unknown}
 */
function serializeCell(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return { [TYPE_KEY]: "date", value: value.toISOString() };
  }
  return value;
}

/**
 * @param {unknown} value
 * @returns {unknown}
 */
function deserializeCell(value) {
  if (value && typeof value === "object" && !Array.isArray(value)) {
    // @ts-ignore - runtime
    if (value[TYPE_KEY] === "date" && typeof value.value === "string") {
      const parsed = new Date(value.value);
      if (!Number.isNaN(parsed.getTime())) return parsed;
    }
  }
  return value;
}

/**
 * @typedef {Object} SerializedTable
 * @property {Column[]} columns
 * @property {unknown[][]} rows
 */

/**
 * @param {DataTable} table
 * @returns {SerializedTable}
 */
export function serializeTable(table) {
  return {
    columns: table.columns.map((c) => ({ name: c.name, type: c.type })),
    rows: table.rows.map((row) => row.map(serializeCell)),
  };
}

/**
 * @param {SerializedTable} data
 * @returns {DataTable}
 */
export function deserializeTable(data) {
  return new DataTable(
    data.columns.map((c) => ({ name: c.name, type: c.type })),
    data.rows.map((row) => row.map(deserializeCell)),
  );
}

