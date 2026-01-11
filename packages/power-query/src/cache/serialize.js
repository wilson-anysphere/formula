import { arrowTableFromIPC, arrowTableToIPC } from "../../../data-io/src/index.js";

import { ArrowTableAdapter } from "../arrowTable.js";
import { DataTable } from "../table.js";

/**
 * @typedef {import("../table.js").Column} Column
 * @typedef {import("../table.js").ITable} ITable
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

/**
 * @typedef {{
 *   kind: "data";
 *   data: SerializedTable;
 * }} SerializedAnyDataTable
 *
 * @typedef {{
 *   kind: "arrow";
 *   format: "ipc";
 *   columns: Column[];
 *   bytes: Uint8Array;
 * }} SerializedAnyArrowTable
 *
 * @typedef {SerializedAnyDataTable | SerializedAnyArrowTable} SerializedAnyTable
 */

/**
 * Serialize a Power Query `ITable` into a versioned payload that can be stored
 * in the cache.
 *
 * v2 supports:
 * - DataTable (row arrays; JSON-friendly, date-tagged)
 * - ArrowTableAdapter (columnar; Arrow IPC bytes)
 *
 * @param {ITable} table
 * @returns {SerializedAnyTable}
 */
export function serializeAnyTable(table) {
  if (table instanceof ArrowTableAdapter) {
    return {
      kind: "arrow",
      format: "ipc",
      columns: table.columns.map((c) => ({ name: c.name, type: c.type })),
      bytes: arrowTableToIPC(table.table),
    };
  }

  const materialized = table instanceof DataTable ? table : new DataTable(table.columns, Array.from(table.iterRows()));
  return { kind: "data", data: serializeTable(materialized) };
}

/**
 * @param {SerializedAnyTable} payload
 * @returns {DataTable | ArrowTableAdapter}
 */
export function deserializeAnyTable(payload) {
  if (!payload || typeof payload !== "object") {
    throw new Error("Invalid cached table payload");
  }

  if (payload.kind === "arrow") {
    const bytes = payload.bytes instanceof Uint8Array ? payload.bytes : new Uint8Array(payload.bytes);
    const table = arrowTableFromIPC(bytes);
    return new ArrowTableAdapter(table, payload.columns);
  }

  if (payload.kind === "data") {
    return deserializeTable(payload.data);
  }

  /** @type {never} */
  const exhausted = payload;
  throw new Error(`Unsupported cached table kind '${exhausted.kind}'`);
}
