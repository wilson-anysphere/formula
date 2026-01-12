/** @type {((table: any) => Uint8Array) | null} */
let arrowTableToIPC = null;
/** @type {((bytes: Uint8Array | ArrayBuffer) => any) | null} */
let arrowTableFromIPC = null;
/** @type {((columns: Record<string, any[] | ArrayLike<any>>) => any) | null} */
let arrowTableFromColumns = null;

try {
  ({ arrowTableFromIPC, arrowTableToIPC, arrowTableFromColumns } = await import("@formula/data-io"));

  // `@formula/data-io` can be present without the heavy optional `apache-arrow`
  // dependency installed (e.g. in some test sandboxes). In that case the module
  // loads, but calling Arrow helpers throws. Treat that as "Arrow IPC unavailable"
  // so callers can fall back to row-backed serialization.
  if (typeof arrowTableToIPC === "function" && typeof arrowTableFromColumns === "function") {
    try {
      const probe = arrowTableFromColumns({ __probe: [1] });
      arrowTableToIPC(probe);
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      if (message.includes("optional 'apache-arrow'")) {
        arrowTableFromIPC = null;
        arrowTableToIPC = null;
        arrowTableFromColumns = null;
      } else {
        throw err;
      }
    }
  }
} catch (err) {
  // Arrow IPC caching is optional; allow Power Query to run without Arrow
  // dependencies installed (e.g. lightweight environments that only use SQL/CSV).
  if (!(err && typeof err === "object" && "code" in err && err.code === "ERR_MODULE_NOT_FOUND")) {
    throw err;
  }
  arrowTableFromIPC = null;
  arrowTableToIPC = null;
  arrowTableFromColumns = null;
}

import { ArrowTableAdapter } from "../arrowTable.js";
import { DataTable } from "../table.js";
import { PqDateTimeZone, PqDecimal, PqDuration, PqTime, hasUtcTimeComponent } from "../values.js";

/**
 * @typedef {import("../table.js").Column} Column
 * @typedef {import("../table.js").ITable} ITable
 */

const TYPE_KEY = "__pq_type";

/**
 * @param {Uint8Array} bytes
 * @returns {string}
 */
function bytesToBase64(bytes) {
  if (typeof Buffer !== "undefined") {
    return Buffer.from(bytes).toString("base64");
  }
  let binary = "";
  for (let i = 0; i < bytes.length; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  // eslint-disable-next-line no-undef
  return btoa(binary);
}

/**
 * @param {string} encoded
 * @returns {Uint8Array}
 */
function base64ToBytes(encoded) {
  if (typeof Buffer !== "undefined") {
    return new Uint8Array(Buffer.from(encoded, "base64"));
  }
  // eslint-disable-next-line no-undef
  const binary = atob(encoded);
  const out = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    out[i] = binary.charCodeAt(i);
  }
  return out;
}

/**
 * @param {unknown} value
 * @returns {unknown}
 */
function serializeCell(value) {
  if (value instanceof PqDateTimeZone) {
    return { [TYPE_KEY]: "datetimezone", value: value.toString() };
  }
  if (value instanceof PqTime) {
    return { [TYPE_KEY]: "time", value: value.toString() };
  }
  if (value instanceof PqDuration) {
    return { [TYPE_KEY]: "duration", value: value.toString() };
  }
  if (value instanceof PqDecimal) {
    return { [TYPE_KEY]: "decimal", value: value.toString() };
  }
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return { [TYPE_KEY]: hasUtcTimeComponent(value) ? "datetime" : "date", value: value.toISOString() };
  }
  if (value instanceof Uint8Array) {
    return { [TYPE_KEY]: "binary", value: bytesToBase64(value) };
  }
  if (value instanceof DataTable || value instanceof ArrowTableAdapter) {
    // Nested tables (e.g. Table.NestedJoin) can appear as cell values. Serialize them
    // with the same versioned table format used for top-level cache entries so
    // cached results roundtrip correctly.
    //
    // When Arrow IPC support is unavailable, materialize Arrow tables into row
    // arrays so lightweight environments can still cache nested-join results.
    const table =
      value instanceof ArrowTableAdapter && !arrowTableToIPC
        ? new DataTable(value.columns, Array.from(value.iterRows()))
        : value;
    return { [TYPE_KEY]: "table", value: serializeAnyTable(table) };
  }
  if (typeof value === "number" && !Number.isFinite(value)) {
    // JSON does not support NaN/Infinity; tag them so roundtrips preserve values.
    return { [TYPE_KEY]: "number", value: String(value) };
  }
  if (typeof value === "bigint") {
    // JSON.stringify throws on bigint; preserve it as a tagged string.
    return { [TYPE_KEY]: "bigint", value: value.toString() };
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
    if (value[TYPE_KEY] === "table" && value.value) {
      try {
        // @ts-ignore - runtime
        return deserializeAnyTable(value.value);
      } catch {
        return value;
      }
    }
    // @ts-ignore - runtime
    if ((value[TYPE_KEY] === "date" || value[TYPE_KEY] === "datetime") && typeof value.value === "string") {
      const parsed = new Date(value.value);
      if (!Number.isNaN(parsed.getTime())) return parsed;
    }
    // @ts-ignore - runtime
    if ((value[TYPE_KEY] === "binary" || value[TYPE_KEY] === "bytes") && typeof value.value === "string") {
      try {
        return base64ToBytes(value.value);
      } catch {
        return value;
      }
    }
    // @ts-ignore - runtime
    if (value[TYPE_KEY] === "datetimezone" && typeof value.value === "string") {
      const parsed = PqDateTimeZone.from(value.value);
      return parsed ?? value;
    }
    // @ts-ignore - runtime
    if (value[TYPE_KEY] === "time" && typeof value.value === "string") {
      const parsed = PqTime.from(value.value);
      return parsed ?? value;
    }
    // @ts-ignore - runtime
    if (value[TYPE_KEY] === "duration" && typeof value.value === "string") {
      const parsed = PqDuration.from(value.value);
      return parsed ?? value;
    }
    // @ts-ignore - runtime
    if (value[TYPE_KEY] === "decimal" && typeof value.value === "string") {
      return new PqDecimal(value.value);
    }
    // @ts-ignore - runtime
    if (value[TYPE_KEY] === "number" && typeof value.value === "string") {
      switch (value.value) {
        case "NaN":
          return Number.NaN;
        case "Infinity":
          return Number.POSITIVE_INFINITY;
        case "-Infinity":
          return Number.NEGATIVE_INFINITY;
        default: {
          const parsed = Number(value.value);
          return Number.isNaN(parsed) ? value : parsed;
        }
      }
    }
    // @ts-ignore - runtime
    if (value[TYPE_KEY] === "bigint" && typeof value.value === "string") {
      try {
        return BigInt(value.value);
      } catch {
        return value;
      }
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
    if (!arrowTableToIPC) {
      throw new Error("Arrow IPC cache serialization requires Arrow IPC support (install @formula/data-io dependencies)");
    }
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
    if (!arrowTableFromIPC) {
      throw new Error("Cached Arrow table requires Arrow IPC support (install @formula/data-io dependencies)");
    }
    const bytes = payload.bytes instanceof Uint8Array ? payload.bytes : new Uint8Array(payload.bytes);
    const table = arrowTableFromIPC(bytes);
    const fieldCount = table?.schema?.fields?.length ?? null;
    if (typeof fieldCount === "number" && Array.isArray(payload.columns) && payload.columns.length !== fieldCount) {
      throw new Error(`Cached Arrow payload is inconsistent (expected ${payload.columns.length} fields, got ${fieldCount})`);
    }
    return new ArrowTableAdapter(table, payload.columns);
  }

  if (payload.kind === "data") {
    return deserializeTable(payload.data);
  }

  /** @type {never} */
  const exhausted = payload;
  throw new Error(`Unsupported cached table kind '${exhausted.kind}'`);
}
