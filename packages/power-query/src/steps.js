import { ArrowTableAdapter } from "./arrowTable.js";
import { DataTable, inferColumnType, makeUniqueColumnNames } from "./table.js";
import { compilePredicate, compileRowPredicate } from "./predicate.js";
import { valueKey } from "./valueKey.js";
import { bindExprColumns, collectExprColumnRefs, evaluateExpr, parseFormula } from "./expr/index.js";
import { MS_PER_DAY, PqDateTimeZone, PqDecimal, PqDuration, PqTime, parseIsoLikeToUtcDate } from "./values.js";

/** @type {((columns: Record<string, any[] | ArrayLike<any>>) => any) | null} */
let arrowTableFromColumns = null;
try {
  ({ arrowTableFromColumns } = await import("../../data-io/src/index.js"));
} catch (err) {
  // Arrow-backed execution is optional; most of Power Query (including SQL folding)
  // can run without Arrow/Parquet dependencies installed.
  if (!(err && typeof err === "object" && "code" in err && err.code === "ERR_MODULE_NOT_FOUND")) {
    throw err;
  }
  arrowTableFromColumns = null;
}

/**
 * @typedef {import("./model.js").QueryOperation} QueryOperation
 * @typedef {import("./model.js").SortSpec} SortSpec
 * @typedef {import("./model.js").Aggregation} Aggregation
 * @typedef {import("./model.js").DataType} DataType
 * @typedef {import("./table.js").ITable} ITable
 */

/**
 * @param {unknown} value
 * @returns {value is Date}
 */
function isDate(value) {
  return value instanceof Date && !Number.isNaN(value.getTime());
}

/**
 * @param {unknown} value
 * @returns {boolean}
 */
function isNullish(value) {
  return value == null;
}

/**
 * @param {unknown} value
 * @returns {unknown}
 */
function normalizeMissing(value) {
  return value === undefined ? null : value;
}

/**
 * Compute a deterministic key for distinctness comparisons.
 *
 * Power Query's `Table.Distinct` is value-based (not referential) for composite
 * values like Dates and records. Reuse the shared `valueKey` helper so local
 * execution matches merge/grouping semantics.
 *
 * Note: `undefined` and `null` compare equal in Power Query tables.
 *
 * @param {unknown} value
 * @returns {string}
 */
function distinctKey(value) {
  return valueKey(normalizeMissing(value));
}

/**
 * @param {unknown} value
 * @returns {string}
 */
function valueToString(value) {
  if (value == null) return "";
  if (isDate(value)) return value.toISOString();
  return String(value);
}

/**
 * @param {unknown} a
 * @param {unknown} b
 * @returns {number}
 */
function compareNonNull(a, b) {
  if (typeof a === "number" && typeof b === "number") return a - b;
  if (typeof a === "boolean" && typeof b === "boolean") return Number(a) - Number(b);
  if (isDate(a) && isDate(b)) return a.getTime() - b.getTime();
  if (a instanceof PqDateTimeZone && b instanceof PqDateTimeZone) return a.toDate().getTime() - b.toDate().getTime();
  if (a instanceof PqTime && b instanceof PqTime) return a.milliseconds - b.milliseconds;
  if (a instanceof PqDuration && b instanceof PqDuration) return a.milliseconds - b.milliseconds;
  if (a instanceof PqDecimal && b instanceof PqDecimal) {
    const aNum = Number(a.value);
    const bNum = Number(b.value);
    if (Number.isFinite(aNum) && Number.isFinite(bNum)) return aNum - bNum;
  }
  return valueToString(a).localeCompare(valueToString(b));
}

/**
 * @param {unknown} a
 * @param {unknown} b
 * @param {{ direction: "ascending" | "descending", nulls: "first" | "last" }} options
 * @returns {number}
 */
function compareValues(a, b, options) {
  const aNull = isNullish(a);
  const bNull = isNullish(b);
  if (aNull && bNull) return 0;
  if (aNull || bNull) {
    const nullCmp = aNull ? -1 : 1;
    const adjusted = options.nulls === "first" ? nullCmp : -nullCmp;
    return adjusted;
  }

  const baseCmp = compareNonNull(a, b);
  return options.direction === "ascending" ? baseCmp : -baseCmp;
}

/**
 * @param {ITable} table
 * @param {string[]} names
 * @returns {number[]}
 */
function indicesForColumns(table, names) {
  return names.map((name) => table.getColumnIndex(name));
}

/**
 * @param {ITable} table
 * @param {string[]} columns
 * @returns {ITable}
 */
function selectColumns(table, columns) {
  const indices = indicesForColumns(table, columns);
  const newColumns = indices.map((idx) => table.columns[idx]);

  if (table instanceof ArrowTableAdapter) {
    return new ArrowTableAdapter(table.table.selectAt(indices), newColumns);
  }

  const newRows = /** @type {DataTable} */ (table).rows.map((row) => indices.map((idx) => row[idx]));
  return new DataTable(newColumns, newRows);
}

/**
 * @param {ITable} table
 * @param {string[]} columns
 * @returns {ITable}
 */
function removeColumns(table, columns) {
  const remove = new Set(columns.map((name) => table.getColumnIndex(name)));
  const keepIndices = table.columns
    .map((_, idx) => idx)
    .filter((idx) => !remove.has(idx));

  const newColumns = keepIndices.map((idx) => table.columns[idx]);

  if (table instanceof ArrowTableAdapter) {
    return new ArrowTableAdapter(table.table.selectAt(keepIndices), newColumns);
  }

  const newRows = /** @type {DataTable} */ (table).rows.map((row) => keepIndices.map((idx) => row[idx]));
  return new DataTable(newColumns, newRows);
}

/**
 * @param {ITable} table
 * @param {string[] | null} columns
 * @returns {ITable}
 */
function distinctRows(table, columns) {
  const indices = columns && columns.length > 0 ? indicesForColumns(table, columns) : table.columns.map((_c, idx) => idx);
  const seen = new Set();

  if (table instanceof ArrowTableAdapter) {
    const vectors = table.columns.map((_c, idx) => table.getColumnVector(idx));
    /** @type {unknown[][]} */
    const outColumns = table.columns.map(() => []);

    for (let rowIndex = 0; rowIndex < table.rowCount; rowIndex++) {
      const keyValues = indices.map((idx) => distinctKey(vectors[idx].get(rowIndex)));
      const key = JSON.stringify(keyValues);
      if (seen.has(key)) continue;
      seen.add(key);
      for (let colIndex = 0; colIndex < vectors.length; colIndex++) {
        outColumns[colIndex].push(vectors[colIndex].get(rowIndex));
      }
    }

    const out = Object.fromEntries(table.columns.map((col, idx) => [col.name, outColumns[idx]]));
    return new ArrowTableAdapter(arrowTableFromColumns(out), table.columns);
  }

  const materialized = ensureDataTable(table);
  const rows = materialized.rows;
  const outRows = [];
  for (const row of rows) {
    const keyValues = indices.map((idx) => distinctKey(row[idx]));
    const key = JSON.stringify(keyValues);
    if (seen.has(key)) continue;
    seen.add(key);
    outRows.push(row);
  }
  return new DataTable(materialized.columns, outRows);
}

/**
 * @param {ITable} table
 * @param {string[] | null} columns
 * @returns {ITable}
 */
function removeRowsWithErrors(table, columns) {
  const indices = columns && columns.length > 0 ? indicesForColumns(table, columns) : table.columns.map((_c, idx) => idx);

  /**
   * @param {unknown} value
   * @returns {boolean}
   */
  const isErrorValue = (value) => value instanceof Error;

  if (table instanceof ArrowTableAdapter) {
    const vectors = table.columns.map((_c, idx) => table.getColumnVector(idx));
    /** @type {unknown[][]} */
    const outColumns = table.columns.map(() => []);

    for (let rowIndex = 0; rowIndex < table.rowCount; rowIndex++) {
      let hasError = false;
      for (const idx of indices) {
        if (isErrorValue(vectors[idx].get(rowIndex))) {
          hasError = true;
          break;
        }
      }
      if (hasError) continue;
      for (let colIndex = 0; colIndex < vectors.length; colIndex++) {
        outColumns[colIndex].push(vectors[colIndex].get(rowIndex));
      }
    }

    const out = Object.fromEntries(table.columns.map((col, idx) => [col.name, outColumns[idx]]));
    return new ArrowTableAdapter(arrowTableFromColumns(out), table.columns);
  }

  const materialized = ensureDataTable(table);
  const rows = materialized.rows;
  const outRows = [];
  for (const row of rows) {
    let hasError = false;
    for (const idx of indices) {
      if (isErrorValue(row[idx])) {
        hasError = true;
        break;
      }
    }
    if (!hasError) outRows.push(row);
  }
  return new DataTable(materialized.columns, outRows);
}

/**
 * @param {ITable} table
 * @param {import("./model.js").FilterPredicate} predicate
 * @returns {ITable}
 */
function filterRows(table, predicate) {
  const fn = compilePredicate(table, predicate);

  if (table instanceof ArrowTableAdapter) {
    if (!arrowTableFromColumns) {
      // Fall back to a row-backed table when Arrow helpers aren't available.
      const outRows = [];
      for (let rowIndex = 0; rowIndex < table.rowCount; rowIndex++) {
        if (fn(rowIndex)) outRows.push(table.getRow(rowIndex));
      }
      return new DataTable(table.columns, outRows);
    }

    /** @type {unknown[][]} */
    const outColumns = table.columns.map(() => []);
    const vectors = table.columns.map((_c, idx) => table.getColumnVector(idx));

    for (let rowIndex = 0; rowIndex < table.rowCount; rowIndex++) {
      if (!fn(rowIndex)) continue;
      for (let colIndex = 0; colIndex < vectors.length; colIndex++) {
        outColumns[colIndex].push(vectors[colIndex].get(rowIndex));
      }
    }

    const out = Object.fromEntries(table.columns.map((col, idx) => [col.name, outColumns[idx]]));
    return new ArrowTableAdapter(arrowTableFromColumns(out), table.columns);
  }

  const sourceRows = /** @type {DataTable} */ (table).rows;
  const newRows = [];
  for (let rowIndex = 0; rowIndex < sourceRows.length; rowIndex++) {
    if (fn(rowIndex)) newRows.push(sourceRows[rowIndex]);
  }
  return new DataTable(table.columns, newRows);
}

/**
 * @param {ITable} table
 * @param {SortSpec[]} sortBy
 * @returns {ITable}
 */
function sortRows(table, sortBy) {
  if (sortBy.length === 0) {
    if (table instanceof ArrowTableAdapter) return table;
    return new DataTable(table.columns, /** @type {DataTable} */ (table).rows);
  }

  const specs = sortBy.map((spec) => ({
    idx: table.getColumnIndex(spec.column),
    direction: spec.direction ?? "ascending",
    nulls: spec.nulls ?? "last",
  }));

  if (table instanceof ArrowTableAdapter) {
    if (!arrowTableFromColumns) {
      // Fall back to a row-backed table when Arrow helpers aren't available.
      const vectors = table.columns.map((_c, idx) => table.getColumnVector(idx));
      const indices = Array.from({ length: table.rowCount }, (_, i) => i);
      indices.sort((a, b) => {
        for (const spec of specs) {
          const cmp = compareValues(vectors[spec.idx].get(a), vectors[spec.idx].get(b), spec);
          if (cmp !== 0) return cmp;
        }
        return a - b;
      });
      return new DataTable(table.columns, indices.map((rowIdx) => table.getRow(rowIdx)));
    }

    const vectors = table.columns.map((_c, idx) => table.getColumnVector(idx));

    const indices = Array.from({ length: table.rowCount }, (_, i) => i);
    indices.sort((a, b) => {
      for (const spec of specs) {
        const cmp = compareValues(vectors[spec.idx].get(a), vectors[spec.idx].get(b), spec);
        if (cmp !== 0) return cmp;
      }
      return a - b;
    });

    /** @type {unknown[][]} */
    const outColumns = vectors.map(() => new Array(indices.length));
    for (let outRow = 0; outRow < indices.length; outRow++) {
      const srcRow = indices[outRow];
      for (let col = 0; col < vectors.length; col++) {
        outColumns[col][outRow] = vectors[col].get(srcRow);
      }
    }

    const out = Object.fromEntries(table.columns.map((col, idx) => [col.name, outColumns[idx]]));
    return new ArrowTableAdapter(arrowTableFromColumns(out), table.columns);
  }

  const decorated = /** @type {DataTable} */ (table).rows.map((row, originalIndex) => ({ row, originalIndex }));
  decorated.sort((a, b) => {
    for (const spec of specs) {
      const cmp = compareValues(a.row[spec.idx], b.row[spec.idx], spec);
      if (cmp !== 0) return cmp;
    }
    return a.originalIndex - b.originalIndex;
  });

  return new DataTable(table.columns, decorated.map((d) => d.row));
}

const NUMBER_TEXT_RE = /^[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?$/;

/**
 * @param {unknown} value
 * @returns {number | null}
 */
function toNumberOrNull(value) {
  if (value instanceof PqDecimal) {
    const parsed = Number(value.value);
    return Number.isFinite(parsed) ? parsed : null;
  }
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value === "string") {
    const trimmed = value.trim();
    if (trimmed === "") return null;
    if (!NUMBER_TEXT_RE.test(trimmed)) return null;
    const num = Number(trimmed);
    return Number.isFinite(num) ? num : null;
  }
  return null;
}

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
 * @param {ITable} table
 * @param {string[]} groupColumns
 * @param {Aggregation[]} aggregations
 * @returns {ITable}
 */
function groupBy(table, groupColumns, aggregations) {
  const groupIdx = indicesForColumns(table, groupColumns);
  const aggSpecs = aggregations.map((agg) => ({
    ...agg,
    idx: table.getColumnIndex(agg.column),
    as: agg.as ?? `${agg.op} of ${agg.column}`,
  }));

  /** @type {Map<string, { keyValues: unknown[], states: any[] }>} */
  const groups = new Map();

  const vectors = table.columns.map((_c, idx) => table.getColumnVector(idx));

  for (let rowIndex = 0; rowIndex < table.rowCount; rowIndex++) {
    const keyValues = groupIdx.map((idx) => {
      return vectors[idx].get(rowIndex);
    });
    const key = JSON.stringify(keyValues.map((v) => valueKey(normalizeMissing(v))));

    let entry = groups.get(key);
    if (!entry) {
      entry = {
        keyValues,
        states: aggSpecs.map((agg) => {
          switch (agg.op) {
            case "sum":
            case "average":
              return { sum: 0, count: 0 };
            case "count":
              return { count: 0 };
            case "min":
              return { value: null, has: false };
            case "max":
              return { value: null, has: false };
            case "countDistinct":
              return { set: new Set() };
            default: {
              /** @type {never} */
              const exhausted = agg.op;
              throw new Error(`Unsupported aggregation '${exhausted}'`);
            }
          }
        }),
      };
      groups.set(key, entry);
    }

    entry.states.forEach((state, idx) => {
      const agg = aggSpecs[idx];
      const value = normalizeMissing(vectors[agg.idx].get(rowIndex));

      switch (agg.op) {
        case "sum": {
          const num = toNumberOrNull(value);
          if (num != null) state.sum += num;
          break;
        }
        case "average": {
          const num = toNumberOrNull(value);
          if (num != null) {
            state.sum += num;
            state.count += 1;
          }
          break;
        }
        case "count":
          state.count += 1;
          break;
        case "min":
          if (value == null) return;
          if (!state.has || compareNonNull(value, state.value) < 0) {
            state.value = value;
            state.has = true;
          }
          break;
        case "max":
          if (value == null) return;
          if (!state.has || compareNonNull(value, state.value) > 0) {
            state.value = value;
            state.has = true;
          }
          break;
        case "countDistinct":
          state.set.add(valueKey(value));
          break;
        default: {
          /** @type {never} */
          const exhausted = agg.op;
          throw new Error(`Unsupported aggregation '${exhausted}'`);
        }
      }
    });
  }

  const resultColumns = [
    ...groupIdx.map((idx) => table.columns[idx]),
    ...aggSpecs.map((agg) => ({
      name: agg.as,
      type:
        agg.op === "sum" || agg.op === "average" || agg.op === "count" || agg.op === "countDistinct"
          ? "number"
          : table.columns[agg.idx].type,
    })),
  ];

  const resultRows = [];
  for (const entry of groups.values()) {
    const row = [...entry.keyValues];
    entry.states.forEach((state, idx) => {
      const agg = aggSpecs[idx];
      switch (agg.op) {
        case "sum":
          row.push(state.sum);
          break;
        case "average":
          row.push(state.count === 0 ? null : state.sum / state.count);
          break;
        case "count":
          row.push(state.count);
          break;
        case "min":
        case "max":
          row.push(state.has ? state.value : null);
          break;
        case "countDistinct":
          row.push(state.set.size);
          break;
        default: {
          /** @type {never} */
          const exhausted = agg.op;
          throw new Error(`Unsupported aggregation '${exhausted}'`);
        }
      }
    });
    resultRows.push(row);
  }

  if (table instanceof ArrowTableAdapter) {
    if (!arrowTableFromColumns) {
      return new DataTable(resultColumns, resultRows);
    }
    const out = Object.fromEntries(resultColumns.map((col, idx) => [col.name, resultRows.map((r) => r[idx])]));
    return new ArrowTableAdapter(arrowTableFromColumns(out), resultColumns);
  }

  return new DataTable(resultColumns, resultRows);
}

/**
 * @param {DataType} type
 * @param {unknown} value
 * @returns {unknown}
 */
function coerceType(type, value) {
  if (value == null) return null;
  switch (type) {
    case "any":
      return value;
    case "string":
      if (value instanceof Uint8Array) return bytesToBase64(value);
      return String(value);
    case "number": {
      const num = toNumberOrNull(value);
      return num == null ? null : num;
    }
    case "boolean":
      if (typeof value === "boolean") return value;
      if (typeof value === "number") return value !== 0;
      if (typeof value === "string") {
        const lower = value.trim().toLowerCase();
        if (lower === "true") return true;
        if (lower === "false") return false;
      }
      return Boolean(value);
    case "date":
      if (isDate(value)) {
        // Represent dates as midnight UTC to avoid leaking time components into a date-typed column.
        return new Date(Date.UTC(value.getUTCFullYear(), value.getUTCMonth(), value.getUTCDate()));
      }
      if (value instanceof PqDateTimeZone) {
        const d = value.toDate();
        return new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()));
      }
      if (typeof value === "number") {
        // Engine convention: numbers are treated as epoch milliseconds.
        // (Excel serial conversion is handled by the host layer.)
        const d = new Date(value);
        if (Number.isNaN(d.getTime())) return null;
        return new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()));
      }
      if (typeof value === "string") {
        const d = parseIsoLikeToUtcDate(value);
        if (!d) return null;
        return new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()));
      }
      return null;
    case "datetime":
      if (isDate(value)) return value;
      if (value instanceof PqDateTimeZone) return value.toDate();
      if (typeof value === "number") {
        // Engine convention: numbers are treated as epoch milliseconds.
        // (Excel serial conversion is handled by the host layer.)
        const d = new Date(value);
        return Number.isNaN(d.getTime()) ? null : d;
      }
      if (typeof value === "string") {
        const d = parseIsoLikeToUtcDate(value);
        return d ?? null;
      }
      return null;
    case "datetimezone":
      if (value instanceof PqDateTimeZone) return value;
      if (isDate(value)) return new PqDateTimeZone(value, 0);
      if (typeof value === "number") {
        // Engine convention: numbers are treated as epoch milliseconds.
        // (Excel serial conversion is handled by the host layer.)
        const d = new Date(value);
        return Number.isNaN(d.getTime()) ? null : new PqDateTimeZone(d, 0);
      }
      if (typeof value === "string") {
        return PqDateTimeZone.from(value) ?? null;
      }
      return null;
    case "time":
      if (value instanceof PqTime) return value;
      if (isDate(value)) {
        return new PqTime(
          value.getUTCHours() * 3_600_000 + value.getUTCMinutes() * 60_000 + value.getUTCSeconds() * 1000 + value.getUTCMilliseconds(),
        );
      }
      if (typeof value === "number" && Number.isFinite(value)) {
        return new PqTime(value * MS_PER_DAY);
      }
      if (typeof value === "string") {
        return PqTime.from(value) ?? null;
      }
      return null;
    case "duration":
      if (value instanceof PqDuration) return value;
      if (typeof value === "number" && Number.isFinite(value)) {
        return new PqDuration(value * MS_PER_DAY);
      }
      if (typeof value === "string") {
        return PqDuration.from(value) ?? null;
      }
      return null;
    case "decimal":
      if (value instanceof PqDecimal) return value;
      if (typeof value === "number" && Number.isFinite(value)) return new PqDecimal(String(value));
      if (typeof value === "bigint") return new PqDecimal(value.toString());
      if (typeof value === "boolean") return new PqDecimal(value ? "1" : "0");
      if (typeof value === "string") {
        const trimmed = value.trim();
        if (trimmed === "") return null;
        if (!NUMBER_TEXT_RE.test(trimmed)) return null;
        return new PqDecimal(trimmed);
      }
      return null;
    case "binary":
      if (value instanceof Uint8Array) return value;
      if (value instanceof ArrayBuffer) return new Uint8Array(value);
      if (typeof value === "string") {
        try {
          return base64ToBytes(value);
        } catch {
          return null;
        }
      }
      return null;
    default: {
      /** @type {never} */
      const exhausted = type;
      throw new Error(`Unsupported type '${exhausted}'`);
    }
  }
}

/**
 * @param {ITable} table
 * @param {string} column
 * @param {DataType} newType
 * @returns {ITable}
 */
function changeType(table, column, newType) {
  const idx = table.getColumnIndex(column);
  const columns = table.columns.map((col, i) => (i === idx ? { ...col, type: newType } : col));

  if (table instanceof ArrowTableAdapter) {
    const requiresRowTable =
      newType === "time" || newType === "duration" || newType === "decimal" || newType === "datetimezone";
    if (!arrowTableFromColumns || requiresRowTable) {
      const vectors = table.columns.map((_c, i) => table.getColumnVector(i));
      const outRows = new Array(table.rowCount);
      for (let rowIndex = 0; rowIndex < table.rowCount; rowIndex++) {
        const row = new Array(columns.length);
        for (let colIndex = 0; colIndex < vectors.length; colIndex++) {
          row[colIndex] = colIndex === idx ? coerceType(newType, vectors[idx].get(rowIndex)) : vectors[colIndex].get(rowIndex);
        }
        outRows[rowIndex] = row;
      }
      return new DataTable(columns, outRows);
    }

    const vectors = table.columns.map((_c, i) => table.getColumnVector(i));
    const outColumns = vectors.map((_v, i) => (i === idx ? new Array(table.rowCount) : null));

    for (let rowIndex = 0; rowIndex < table.rowCount; rowIndex++) {
      outColumns[idx][rowIndex] = coerceType(newType, vectors[idx].get(rowIndex));
    }

    // For unchanged columns, reuse values by reading them out lazily; this avoids materializing
    // row arrays but still constructs new column arrays for Arrow ingestion.
    for (let colIndex = 0; colIndex < vectors.length; colIndex++) {
      if (colIndex === idx) continue;
      const colValues = new Array(table.rowCount);
      const vec = vectors[colIndex];
      for (let rowIndex = 0; rowIndex < table.rowCount; rowIndex++) {
        colValues[rowIndex] = vec.get(rowIndex);
      }
      outColumns[colIndex] = colValues;
    }

    const out = Object.fromEntries(columns.map((col, i) => [col.name, outColumns[i]]));
    return new ArrowTableAdapter(arrowTableFromColumns(out), columns);
  }

  const rows = /** @type {DataTable} */ (table).rows.map((row) => {
    const next = row.slice();
    next[idx] = coerceType(newType, row[idx]);
    return next;
  });
  return new DataTable(columns, rows);
}

/**
 * Compile an `addColumn` formula into a row function.
 *
 * The supported surface area is intentionally small and is parsed/evaluated by
 * the sandboxed expression engine in `src/expr/*`. Columns can be referenced
 * via `[Column Name]`.
 *
 * @param {DataTable} table
 * @param {string} formula
 * @returns {(values: unknown[]) => unknown}
 */
export function compileRowFormula(table, formula) {
  const expr = parseFormula(formula);
  const bound = bindExprColumns(expr, (name) => table.getColumnIndex(name));
  return (values) => evaluateExpr(bound, values);
}

/**
 * Compile an `addColumn` formula into a row function for streaming execution.
 *
 * This variant binds column references against a plain column list (rather than
 * requiring a `DataTable` instance).
 *
 * @param {Array<{ name: string }>} columns
 * @param {string} formula
 * @returns {(values: unknown[]) => unknown}
 */
export function compileRowFormulaForColumns(columns, formula) {
  /** @type {Map<string, number>} */
  const index = new Map();
  for (let i = 0; i < columns.length; i++) {
    const name = columns[i]?.name;
    if (typeof name === "string") index.set(name, i);
  }

  const expr = parseFormula(formula);
  const bound = bindExprColumns(expr, (name) => {
    const idx = index.get(name);
    if (idx == null) {
      throw new Error(`Unknown column '${name}'. Available: ${columns.map((c) => c.name).join(", ")}`);
    }
    return idx;
  });
  return (values) => evaluateExpr(bound, values);
}

/**
 * Compile a `transformColumns` formula into a value function.
 *
 * @param {string} formula
 * @returns {(value: unknown) => unknown}
 */
export function compileValueFormula(formula) {
  const expr = parseFormula(formula);
  const refs = collectExprColumnRefs(expr);
  if (refs.size > 0) {
    throw new Error(`Value formulas cannot reference columns: ${Array.from(refs).join(", ")}`);
  }

  return (value) => evaluateExpr(expr, [], null, value);
}

/**
 * @param {DataTable} table
 * @param {string} name
 * @param {string} formula
 * @returns {DataTable}
 */
function addColumn(table, name, formula) {
  if (table.columnIndex.has(name)) {
    throw new Error(`Column '${name}' already exists`);
  }

  const compute = compileRowFormula(table, formula);
  const rows = table.rows.map((row) => [...row, compute(row)]);
  const type = inferColumnType(rows.map((row) => row[row.length - 1]));
  const columns = [...table.columns, { name, type }];
  return new DataTable(columns, rows);
}

/**
 * @param {DataTable} table
 * @param {import("./model.js").TransformColumnsOp} op
 * @returns {DataTable}
 */
function transformColumns(table, op) {
  const transforms = op.transforms.map((t) => ({
    ...t,
    idx: table.getColumnIndex(t.column),
    fn: compileValueFormula(t.formula),
  }));

  const rows = table.rows.map((row) => row.slice());
  for (const t of transforms) {
    for (const row of rows) {
      const next = t.fn(row[t.idx]);
      row[t.idx] = t.newType ? coerceType(t.newType, next) : next;
    }
  }

  const columns = table.columns.map((col, idx) => {
    const t = transforms.find((x) => x.idx === idx);
    if (!t) return col;
    const type = t.newType ?? inferColumnType(rows.map((r) => r[idx]));
    return { ...col, type };
  });

  return new DataTable(columns, rows);
}

/**
 * @param {ITable} table
 * @param {string} oldName
 * @param {string} newName
 * @returns {ITable}
 */
function renameColumn(table, oldName, newName) {
  const idx = table.getColumnIndex(oldName);
  if (newName !== oldName && table.columns.some((col, i) => i !== idx && col.name === newName)) {
    throw new Error(`Column '${newName}' already exists`);
  }

  const columns = table.columns.map((col, i) => (i === idx ? { ...col, name: newName } : col));

  if (table instanceof ArrowTableAdapter) {
    return new ArrowTableAdapter(table.table, columns);
  }

  if (table instanceof DataTable) {
    return new DataTable(columns, table.rows);
  }

  const materialized = ensureDataTable(table);
  return new DataTable(columns, materialized.rows);
}

/**
 * @param {DataTable} table
 * @param {import("./model.js").PivotOp} op
 * @returns {DataTable}
 */
function pivot(table, op) {
  const pivotIdx = table.getColumnIndex(op.rowColumn);
  const valueIdx = table.getColumnIndex(op.valueColumn);

  const keyIndices = table.columns
    .map((_c, idx) => idx)
    .filter((idx) => idx !== pivotIdx && idx !== valueIdx);

  /**
   * @param {unknown} value
   * @returns {string}
   */
  const pivotKey = (value) => {
    return valueKey(value);
  };

  /**
   * @param {unknown} value
   * @returns {string}
   */
  const pivotDisplayName = (value) => {
    if (value == null) return "(null)";
    const text = valueToString(value);
    return text === "" ? "(blank)" : text;
  };

  /** @type {string[]} */
  const pivotKeys = [];
  /** @type {string[]} */
  const pivotDisplayNames = [];
  /** @type {Set<string>} */
  const pivotSeen = new Set();
  for (const row of table.rows) {
    const raw = row[pivotIdx];
    const key = pivotKey(raw);
    if (pivotSeen.has(key)) continue;
    pivotSeen.add(key);
    pivotKeys.push(key);
    pivotDisplayNames.push(pivotDisplayName(raw));
  }

  const pivotColumnNames = makeUniqueColumnNames(pivotDisplayNames);
  const pivotKeyToPos = new Map(pivotKeys.map((key, idx) => [key, idx]));

  /** @type {Map<string, { keyValues: unknown[], states: any[] }>} */
  const groups = new Map();

  for (const row of table.rows) {
    const keyValues = keyIndices.map((idx) => row[idx]);
    const key = JSON.stringify(keyValues.map((v) => valueKey(normalizeMissing(v))));

    let entry = groups.get(key);
    if (!entry) {
      entry = {
        keyValues,
        states: pivotKeys.map(() => ({ sum: 0, count: 0, min: null, max: null, has: false })),
      };
      groups.set(key, entry);
    }

    const pivotPos = pivotKeyToPos.get(pivotKey(row[pivotIdx]));
    if (pivotPos == null) continue;

    const state = entry.states[pivotPos];
    const value = normalizeMissing(row[valueIdx]);

    switch (op.aggregation) {
      case "sum": {
        const num = toNumberOrNull(value);
        if (num != null) state.sum += num;
        break;
      }
      case "count":
        if (value != null) state.count += 1;
        break;
      case "average": {
        const num = toNumberOrNull(value);
        if (num != null) {
          state.sum += num;
          state.count += 1;
        }
        break;
      }
      case "min":
        if (value == null) break;
        if (!state.has || compareNonNull(value, state.min) < 0) {
          state.min = value;
          state.has = true;
        }
        break;
      case "max":
        if (value == null) break;
        if (!state.has || compareNonNull(value, state.max) > 0) {
          state.max = value;
          state.has = true;
        }
        break;
      default:
        throw new Error(`Unsupported pivot aggregation '${op.aggregation}'`);
    }
  }

  const keyColumns = keyIndices.map((idx) => table.columns[idx]);
  const valueColumns = pivotColumnNames.map((name) => ({
    name,
    type:
      op.aggregation === "sum" || op.aggregation === "average" || op.aggregation === "count"
        ? "number"
        : "any",
  }));
  const columns = [...keyColumns, ...valueColumns];

  const rows = [];
  for (const entry of groups.values()) {
    const row = [...entry.keyValues];
    for (const state of entry.states) {
      switch (op.aggregation) {
        case "sum":
          row.push(state.sum);
          break;
        case "count":
          row.push(state.count);
          break;
        case "average":
          row.push(state.count === 0 ? null : state.sum / state.count);
          break;
        case "min":
          row.push(state.has ? state.min : null);
          break;
        case "max":
          row.push(state.has ? state.max : null);
          break;
        default:
          throw new Error(`Unsupported pivot aggregation '${op.aggregation}'`);
      }
    }
    rows.push(row);
  }

  return new DataTable(columns, rows);
}

/**
 * @param {DataTable} table
 * @param {import("./model.js").UnpivotOp} op
 * @returns {DataTable}
 */
function unpivot(table, op) {
  const unpivotIdx = indicesForColumns(table, op.columns);
  const unpivotSet = new Set(unpivotIdx);
  const keepIdx = table.columns.map((_, idx) => idx).filter((idx) => !unpivotSet.has(idx));

  const keepColumns = keepIdx.map((idx) => table.columns[idx]);
  const columns = [
    ...keepColumns,
    { name: op.nameColumn, type: "string" },
    { name: op.valueColumn, type: "any" },
  ];

  const rows = [];
  for (const row of table.rows) {
    const prefix = keepIdx.map((idx) => row[idx]);
    for (const idx of unpivotIdx) {
      rows.push([...prefix, table.columns[idx].name, row[idx]]);
    }
  }

  return new DataTable(columns, rows);
}

/**
 * @param {DataTable} table
 * @param {import("./model.js").FillDownOp} op
 * @returns {DataTable}
 */
function fillDown(table, op) {
  const indices = indicesForColumns(table, op.columns);
  const rows = table.rows.map((row) => row.slice());

  const last = new Map(indices.map((idx) => [idx, null]));
  for (const row of rows) {
    for (const idx of indices) {
      const value = row[idx];
      if (value == null) {
        row[idx] = last.get(idx);
      } else {
        last.set(idx, value);
      }
    }
  }

  return new DataTable(table.columns, rows);
}

/**
 * @param {DataTable} table
 * @param {import("./model.js").ReplaceValuesOp} op
 * @returns {DataTable}
 */
function replaceValues(table, op) {
  const idx = table.getColumnIndex(op.column);
  const findKey = valueKey(op.find);
  const rows = table.rows.map((row) => {
    const next = row.slice();
    if (valueKey(next[idx]) === findKey) next[idx] = op.replace;
    return next;
  });
  return new DataTable(table.columns, rows);
}

/**
 * @param {DataTable} table
 * @param {import("./model.js").SplitColumnOp} op
 * @returns {DataTable}
 */
function splitColumn(table, op) {
  const idx = table.getColumnIndex(op.column);
  const partsByRow = table.rows.map((row) => valueToString(row[idx]).split(op.delimiter));
  // Avoid `Math.max(...partsByRow.map(...))` because spreading large arrays can exceed the
  // VM argument limit / call stack on big datasets.
  let maxParts = 0;
  for (const parts of partsByRow) {
    if (parts.length > maxParts) maxParts = parts.length;
  }

  const requestedNames = Array.isArray(op.newColumns) && op.newColumns.length > 0 ? op.newColumns.slice() : null;
  if (requestedNames && maxParts > requestedNames.length) {
    throw new Error(`Split produced ${maxParts} columns but only ${requestedNames.length} names were provided`);
  }

  const baseNames = requestedNames ?? Array.from({ length: maxParts }, (_, i) => (i === 0 ? op.column : `${op.column}.${i + 1}`));
  const unique = requestedNames ? baseNames : makeUniqueColumnNames(baseNames);

  const columns = table.columns
    .map((col, i) => (i === idx ? { ...col, name: unique[0], type: "string" } : col))
    .concat(unique.slice(1).map((name) => ({ name, type: "string" })));

  const rows = table.rows.map((row, rIdx) => {
    const next = row.slice();
    const parts = partsByRow[rIdx];
    next[idx] = parts[0] ?? "";
    const expectedParts = requestedNames ? requestedNames.length : maxParts;
    for (let i = 1; i < expectedParts; i++) {
      next.push(parts[i] ?? null);
    }
    return next;
  });

  return new DataTable(columns, rows);
}

/**
 * @param {DataTable} table
 * @returns {DataTable}
 */
function promoteHeaders(table) {
  if (table.rows.length === 0) return table;

  const width = table.columns.length;
  const header = table.rows[0] ?? [];
  const names = makeUniqueColumnNames(Array.from({ length: width }, (_, i) => header[i]));
  const rows = table.rows.slice(1);
  const columns = names.map((name, idx) => ({ name, type: inferColumnType(rows.map((row) => row[idx])) }));
  return new DataTable(columns, rows);
}

/**
 * @param {DataTable} table
 * @returns {DataTable}
 */
function demoteHeaders(table) {
  const width = table.columns.length;
  const headerRow = table.columns.map((c) => c.name);
  const columns = Array.from({ length: width }, (_v, i) => ({ name: `Column${i + 1}`, type: "any" }));
  const rows = [headerRow, ...table.rows];
  return new DataTable(columns, rows);
}

/**
 * @param {ITable} table
 * @param {import("./model.js").ReorderColumnsOp} op
 * @returns {ITable}
 */
function reorderColumns(table, op) {
  const missingField = op.missingField ?? "error";
  const specified = op.columns;
  const seen = new Set();
  /** @type {(number | null)[]} */
  const order = [];
  /** @type {string[]} */
  const missingNames = [];

  for (const name of specified) {
    if (seen.has(name)) {
      throw new Error(`Duplicate column name '${name}' in reorder list`);
    }
    seen.add(name);
    const idx = table.columns.findIndex((c) => c.name === name);
    if (idx >= 0) {
      order.push(idx);
      continue;
    }
    if (missingField === "ignore") continue;
    if (missingField === "useNull") {
      order.push(null);
      missingNames.push(name);
      continue;
    }
    throw new Error(`Unknown column '${name}'`);
  }

  for (let idx = 0; idx < table.columns.length; idx++) {
    const name = table.columns[idx].name;
    if (seen.has(name)) continue;
    order.push(idx);
  }

  // Fast path: Arrow-backed reorder without adding new null columns.
  if (table instanceof ArrowTableAdapter && missingNames.length === 0) {
    const indices = /** @type {number[]} */ (order);
    const columns = indices.map((idx) => table.columns[idx]);
    return new ArrowTableAdapter(table.table.selectAt(indices), columns);
  }

  const materialized = ensureDataTable(table);
  const rows = materialized.rows.map((row) =>
    order.map((idxOrNull) => (idxOrNull == null ? null : row[idxOrNull])),
  );
  const columns = [];
  let missingIndex = 0;
  for (const idxOrNull of order) {
    if (idxOrNull == null) {
      columns.push({ name: missingNames[missingIndex++] ?? "", type: "any" });
    } else {
      columns.push(materialized.columns[idxOrNull]);
    }
  }
  return new DataTable(columns, rows);
}

/**
 * @param {DataTable} table
 * @param {import("./model.js").AddIndexColumnOp} op
 * @returns {DataTable}
 */
function addIndexColumn(table, op) {
  if (table.columnIndex.has(op.name)) {
    throw new Error(`Column '${op.name}' already exists`);
  }
  const rows = table.rows.map((row, i) => [...row, op.initialValue + i * op.increment]);
  const columns = [...table.columns, { name: op.name, type: "number" }];
  return new DataTable(columns, rows);
}

/**
 * @param {DataTable} table
 * @param {import("./model.js").CombineColumnsOp} op
 * @returns {DataTable}
 */
function combineColumns(table, op) {
  const indices = indicesForColumns(table, op.columns);
  const remove = new Set(indices);
  if (indices.length === 0) {
    throw new Error("combineColumns requires at least one column");
  }
  // Avoid `Math.min(...indices)` because spreading large arrays can exceed the
  // VM argument limit / call stack on very wide datasets.
  let insertAt = indices[0];
  for (const idx of indices) {
    if (idx < insertAt) insertAt = idx;
  }

  if (table.columns.some((col, idx) => !remove.has(idx) && col.name === op.newColumnName)) {
    throw new Error(`Column '${op.newColumnName}' already exists`);
  }

  const columns = [];
  for (let i = 0; i < table.columns.length; i++) {
    if (i === insertAt) {
      columns.push({ name: op.newColumnName, type: "string" });
    }
    if (remove.has(i)) continue;
    columns.push(table.columns[i]);
  }

  const rows = table.rows.map((row) => {
    const combined = indices.map((idx) => valueToString(row[idx])).join(op.delimiter);
    const next = [];
    for (let i = 0; i < row.length; i++) {
      if (i === insertAt) {
        next.push(combined);
      }
      if (remove.has(i)) continue;
      next.push(row[i]);
    }
    return next;
  });

  return new DataTable(columns, rows);
}

/**
 * @param {ITable} table
 * @param {import("./model.js").TransformColumnNamesOp} op
 * @returns {ITable}
 */
function transformColumnNames(table, op) {
  const rawNames = table.columns.map((col) => {
    switch (op.transform) {
      case "upper":
        return col.name.toUpperCase();
      case "lower":
        return col.name.toLowerCase();
      case "trim":
        return col.name.trim();
      default:
        return col.name;
    }
  });
  const names = makeUniqueColumnNames(rawNames);
  const columns = table.columns.map((col, idx) => ({ ...col, name: names[idx] }));

  if (table instanceof ArrowTableAdapter) {
    return new ArrowTableAdapter(table.table, columns);
  }
  if (table instanceof DataTable) {
    return new DataTable(columns, table.rows);
  }
  const materialized = ensureDataTable(table);
  return new DataTable(columns, materialized.rows);
}

/**
 * @param {DataTable} table
 * @param {import("./model.js").ReplaceErrorValuesOp} op
 * @returns {DataTable}
 */
function replaceErrorValues(table, op) {
  const replacements = new Map(op.replacements.map((r) => [table.getColumnIndex(r.column), r.value]));
  const rows = table.rows.map((row) => {
    const next = row.slice();
    for (const [idx, value] of replacements.entries()) {
      if (next[idx] instanceof Error) next[idx] = value;
    }
    return next;
  });
  return new DataTable(table.columns, rows);
}

/**
 * @param {ITable} table
 * @param {number} count
 * @returns {ITable}
 */
function take(table, count) {
  if (!Number.isFinite(count) || count < 0) {
    throw new Error(`Invalid take count '${count}'`);
  }
  return table.head(count);
}

/**
 * Operation types that can be executed incrementally (per batch) without
 * materializing the full table in memory.
 *
 * Note: this list intentionally excludes order-sensitive or stateful operations
 * like `sortRows` / `groupBy` / `pivot` / `merge` / `append`.
 */
export const STREAMABLE_OPERATION_TYPES = new Set([
  "selectColumns",
  "removeColumns",
  "filterRows",
  "addColumn",
  "renameColumn",
  "changeType",
  "transformColumns",
  "take",
  "skip",
  "removeRows",
  "fillDown",
  "replaceValues",
  "removeRowsWithErrors",
  "distinctRows",
  "reorderColumns",
  "addIndexColumn",
  "combineColumns",
  "transformColumnNames",
  "replaceErrorValues",
  "splitColumn",
  "demoteHeaders",
]);

/**
 * @param {QueryOperation} operation
 */
export function isStreamableOperation(operation) {
  if (operation.type === "splitColumn") {
    return Array.isArray(operation.newColumns) && operation.newColumns.length > 0;
  }
  return STREAMABLE_OPERATION_TYPES.has(operation.type);
}

/**
 * Determine whether a sequence of operations can be executed incrementally without materializing.
 *
 * This includes special-casing `promoteHeaders`, which is streamable only once per sequence.
 *
 * The engine consumes the first data row at the `promoteHeaders` point to derive the new header
 * names, then continues streaming the remaining operations.
 *
 * @param {QueryOperation[]} operations
 * @returns {boolean}
 */
export function isStreamableOperationSequence(operations) {
  let promoteHeadersCount = 0;
  for (const op of operations) {
    if (op.type === "promoteHeaders") {
      promoteHeadersCount += 1;
      if (promoteHeadersCount > 1) return false;
      continue;
    }
    if (!isStreamableOperation(op)) return false;
  }
  return true;
}

/**
 * Compile a list of streamable operations into a batch transformer.
 *
 * This is used by `QueryEngine.executeQueryStreaming(..., { materialize: false })`
 * to keep memory bounded while applying common transformations.
 *
 * @param {QueryOperation[]} operations
 * @param {import("./table.js").Column[]} inputColumns
 * @returns {{
 *   columns: import("./table.js").Column[];
 *   transformBatch: (rows: unknown[][]) => { rows: unknown[][]; done: boolean };
 * }}
 */
export function compileStreamingPipeline(operations, inputColumns) {
  const normalizeCell = normalizeMissing;

  /** @type {import("./table.js").Column[]} */
  let columns = inputColumns.map((c) => ({ name: c.name, type: c.type ?? "any" }));

  /**
   * @param {import("./table.js").Column[]} cols
   */
  const buildIndex = (cols) => new Map(cols.map((c, idx) => [c.name, idx]));

  /** @type {Map<string, number>} */
  let columnIndex = buildIndex(columns);

  /**
   * @param {string} name
   */
  const getColumnIndex = (name) => {
    const idx = columnIndex.get(name);
    if (idx == null) {
      throw new Error(`Unknown column '${name}'. Available: ${columns.map((c) => c.name).join(", ")}`);
    }
    return idx;
  };

  /**
   * @typedef {(rows: unknown[][]) => { rows: unknown[][]; done: boolean }} BatchTransform
   */

  /** @type {BatchTransform[]} */
  const transforms = [];

  for (const op of operations) {
    switch (op.type) {
      case "selectColumns": {
        const indices = op.columns.map((name) => getColumnIndex(name));
        const outColumns = indices.map((idx) => columns[idx]);
        columns = outColumns;
        columnIndex = buildIndex(columns);

        transforms.push((rows) => ({
          rows: rows.map((row) => indices.map((idx) => normalizeCell(row?.[idx]))),
          done: false,
        }));
        break;
      }
      case "removeColumns": {
        const remove = new Set(op.columns.map((name) => getColumnIndex(name)));
        const keepIndices = columns
          .map((_c, idx) => idx)
          .filter((idx) => !remove.has(idx));

        const outColumns = keepIndices.map((idx) => columns[idx]);
        columns = outColumns;
        columnIndex = buildIndex(columns);

        transforms.push((rows) => ({
          rows: rows.map((row) => keepIndices.map((idx) => normalizeCell(row?.[idx]))),
          done: false,
        }));
        break;
      }
      case "filterRows": {
        const predicate = compileRowPredicate(columns, op.predicate);
        transforms.push((rows) => ({ rows: rows.filter((row) => predicate(row)), done: false }));
        break;
      }
      case "addColumn": {
        if (columnIndex.has(op.name)) {
          throw new Error(`Column '${op.name}' already exists`);
        }
        const compute = compileRowFormulaForColumns(columns, op.formula);
        columns = [...columns, { name: op.name, type: "any" }];
        columnIndex = buildIndex(columns);

        transforms.push((rows) => ({
          rows: rows.map((row) => {
            const computed = normalizeCell(compute(row));
            return [...row, computed];
          }),
          done: false,
        }));
        break;
      }
      case "renameColumn": {
        const idx = getColumnIndex(op.oldName);
        if (op.newName !== op.oldName && columns.some((col, i) => i !== idx && col.name === op.newName)) {
          throw new Error(`Column '${op.newName}' already exists`);
        }
        columns = columns.map((col, i) => (i === idx ? { ...col, name: op.newName } : col));
        columnIndex = buildIndex(columns);
        transforms.push((rows) => ({ rows, done: false }));
        break;
      }
      case "changeType": {
        const idx = getColumnIndex(op.column);
        columns = columns.map((col, i) => (i === idx ? { ...col, type: op.newType } : col));
        columnIndex = buildIndex(columns);

        transforms.push((rows) => ({
          rows: rows.map((row) => {
            const next = row.slice();
            next[idx] = coerceType(op.newType, row?.[idx]);
            return next;
          }),
          done: false,
        }));
        break;
      }
      case "transformColumns": {
        const compiled = op.transforms.map((t) => ({
          idx: getColumnIndex(t.column),
          newType: t.newType ?? null,
          fn: compileValueFormula(t.formula),
        }));

        columns = columns.map((col, idx) => {
          const t = compiled.find((x) => x.idx === idx);
          if (!t) return col;
          return { ...col, type: t.newType ?? "any" };
        });
        columnIndex = buildIndex(columns);

        transforms.push((rows) => {
          const out = rows.map((row) => row.slice());
          for (const t of compiled) {
            for (const row of out) {
              const next = t.fn(row[t.idx]);
              row[t.idx] = t.newType ? coerceType(t.newType, next) : normalizeCell(next);
            }
          }
          return { rows: out, done: false };
        });
        break;
      }
      case "take": {
        if (!Number.isFinite(op.count) || op.count < 0) {
          throw new Error(`Invalid take count '${op.count}'`);
        }
        let seen = 0;
        const limit = op.count;
        transforms.push((rows) => {
          const remaining = limit - seen;
          if (remaining <= 0) return { rows: [], done: true };
          const slice = rows.slice(0, remaining);
          seen += slice.length;
          return { rows: slice, done: seen >= limit };
        });
        break;
      }
      case "skip": {
        if (!Number.isFinite(op.count) || op.count < 0) {
          throw new Error(`Invalid skip count '${op.count}'`);
        }
        let skipped = 0;
        const skipCount = op.count;
        transforms.push((rows) => {
          const remaining = skipCount - skipped;
          if (remaining <= 0) return { rows, done: false };
          if (rows.length === 0) return { rows: [], done: false };
          if (remaining >= rows.length) {
            skipped += rows.length;
            return { rows: [], done: false };
          }
          skipped = skipCount;
          return { rows: rows.slice(remaining), done: false };
        });
        break;
      }
      case "removeRows": {
        if (!Number.isFinite(op.offset) || op.offset < 0 || !Number.isFinite(op.count) || op.count < 0) {
          throw new Error(`Invalid removeRows range (${op.offset}, ${op.count})`);
        }
        const start = Math.floor(op.offset);
        const end = start + Math.floor(op.count);
        let seen = 0;
        transforms.push((rows) => {
          if (rows.length === 0) return { rows: [], done: false };
          const batchStart = seen;
          const batchEnd = seen + rows.length;
          seen = batchEnd;

          // No overlap with the removal window.
          if (batchEnd <= start || batchStart >= end) return { rows, done: false };

          const prefixLen = Math.max(0, Math.min(rows.length, start - batchStart));
          const suffixStart = Math.max(prefixLen, Math.min(rows.length, end - batchStart));
          const prefix = prefixLen > 0 ? rows.slice(0, prefixLen) : [];
          const suffix = suffixStart < rows.length ? rows.slice(suffixStart) : [];
          return { rows: prefix.length === 0 ? suffix : prefix.concat(suffix), done: false };
        });
        break;
      }
      case "fillDown": {
        const indices = op.columns.map((name) => getColumnIndex(name));
        /** @type {Map<number, unknown>} */
        const last = new Map(indices.map((idx) => [idx, null]));
        transforms.push((rows) => ({
          rows: rows.map((row) => {
            const next = row.slice();
            for (const idx of indices) {
              const value = next[idx];
              if (value == null) {
                next[idx] = last.get(idx);
              } else {
                last.set(idx, value);
              }
            }
            return next;
          }),
          done: false,
        }));
        break;
      }
      case "replaceValues": {
        const idx = getColumnIndex(op.column);
        const findKey = valueKey(op.find);
        transforms.push((rows) => ({
          rows: rows.map((row) => {
            const next = row.slice();
            if (valueKey(next[idx]) === findKey) next[idx] = op.replace;
            return next;
          }),
          done: false,
        }));
        break;
      }
      case "removeRowsWithErrors": {
        const indices =
          op.columns && op.columns.length > 0
            ? op.columns.map((name) => getColumnIndex(name))
            : columns.map((_c, idx) => idx);
        transforms.push((rows) => ({
          rows: rows.filter((row) => !indices.some((idx) => row?.[idx] instanceof Error)),
          done: false,
        }));
        break;
      }
      case "distinctRows": {
        const indices =
          op.columns && op.columns.length > 0
            ? op.columns.map((name) => getColumnIndex(name))
            : columns.map((_c, idx) => idx);
        const seen = new Set();
        transforms.push((rows) => ({
          rows: rows.filter((row) => {
            const keyValues = indices.map((idx) => distinctKey(row?.[idx]));
            const key = JSON.stringify(keyValues);
            if (seen.has(key)) return false;
            seen.add(key);
            return true;
          }),
          done: false,
        }));
        break;
      }
      case "reorderColumns": {
        const missingField = op.missingField ?? "error";
        const specified = op.columns;
        const seen = new Set();
        /** @type {(number | null)[]} */
        const order = [];
        /** @type {string[]} */
        const missingNames = [];

        for (const name of specified) {
          if (seen.has(name)) {
            throw new Error(`Duplicate column name '${name}' in reorder list`);
          }
          seen.add(name);
          const idx = columns.findIndex((c) => c.name === name);
          if (idx >= 0) {
            order.push(idx);
            continue;
          }
          if (missingField === "ignore") continue;
          if (missingField === "useNull") {
            order.push(null);
            missingNames.push(name);
            continue;
          }
          throw new Error(`Unknown column '${name}'`);
        }

        for (let idx = 0; idx < columns.length; idx++) {
          const name = columns[idx].name;
          if (seen.has(name)) continue;
          order.push(idx);
        }

        const outColumns = [];
        let missingIndex = 0;
        for (const idxOrNull of order) {
          if (idxOrNull == null) {
            outColumns.push({ name: missingNames[missingIndex++] ?? "", type: "any" });
          } else {
            outColumns.push(columns[idxOrNull]);
          }
        }

        columns = outColumns;
        columnIndex = buildIndex(columns);

        transforms.push((rows) => ({
          rows: rows.map((row) => order.map((idxOrNull) => (idxOrNull == null ? null : normalizeCell(row?.[idxOrNull])))),
          done: false,
        }));
        break;
      }
      case "addIndexColumn": {
        if (columnIndex.has(op.name)) {
          throw new Error(`Column '${op.name}' already exists`);
        }
        let index = 0;
        const initialValue = op.initialValue;
        const increment = op.increment;
        columns = [...columns, { name: op.name, type: "number" }];
        columnIndex = buildIndex(columns);

        transforms.push((rows) => ({
          rows: rows.map((row) => [...row, initialValue + index++ * increment]),
          done: false,
        }));
        break;
      }
      case "combineColumns": {
        const indices = op.columns.map((name) => getColumnIndex(name));
        const remove = new Set(indices);
        if (indices.length === 0) {
          throw new Error("combineColumns requires at least one column");
        }
        // Avoid `Math.min(...indices)` because spreading large arrays can exceed the
        // VM argument limit / call stack on very wide datasets.
        let insertAt = indices[0];
        for (const idx of indices) {
          if (idx < insertAt) insertAt = idx;
        }

        if (columns.some((col, idx) => !remove.has(idx) && col.name === op.newColumnName)) {
          throw new Error(`Column '${op.newColumnName}' already exists`);
        }

        const nextColumns = [];
        for (let i = 0; i < columns.length; i++) {
          if (i === insertAt) {
            nextColumns.push({ name: op.newColumnName, type: "string" });
          }
          if (remove.has(i)) continue;
          nextColumns.push(columns[i]);
        }
        columns = nextColumns;
        columnIndex = buildIndex(columns);

        transforms.push((rows) => ({
          rows: rows.map((row) => {
            const combined = indices.map((idx) => valueToString(row?.[idx])).join(op.delimiter);
            const next = [];
            for (let i = 0; i < row.length; i++) {
              if (i === insertAt) {
                next.push(combined);
              }
              if (remove.has(i)) continue;
              next.push(normalizeCell(row?.[i]));
            }
            return next;
          }),
          done: false,
        }));
        break;
      }
      case "transformColumnNames": {
        const rawNames = columns.map((col) => {
          switch (op.transform) {
            case "upper":
              return col.name.toUpperCase();
            case "lower":
              return col.name.toLowerCase();
            case "trim":
              return col.name.trim();
            default:
              return col.name;
          }
        });
        const names = makeUniqueColumnNames(rawNames);
        columns = columns.map((col, idx) => ({ ...col, name: names[idx] }));
        columnIndex = buildIndex(columns);
        transforms.push((rows) => ({ rows, done: false }));
        break;
      }
      case "replaceErrorValues": {
        const replacements = op.replacements.map((r) => ({ idx: getColumnIndex(r.column), value: r.value }));
        transforms.push((rows) => ({
          rows: rows.map((row) => {
            const next = row.slice();
            for (const entry of replacements) {
              if (next[entry.idx] instanceof Error) next[entry.idx] = entry.value;
            }
            return next;
          }),
          done: false,
        }));
        break;
      }
      case "splitColumn": {
        if (!Array.isArray(op.newColumns) || op.newColumns.length === 0) {
          throw new Error("Streaming splitColumn requires an explicit newColumns list");
        }

        const idx = getColumnIndex(op.column);
        const names = op.newColumns.slice();

        columns = columns
          .map((col, i) => (i === idx ? { ...col, name: names[0], type: "string" } : col))
          .concat(names.slice(1).map((name) => ({ name, type: "string" })));

        const seenNames = new Set();
        for (const col of columns) {
          if (seenNames.has(col.name)) {
            throw new Error(`Duplicate column name '${col.name}'`);
          }
          seenNames.add(col.name);
        }
        columnIndex = buildIndex(columns);

        const expectedParts = names.length;
        transforms.push((rows) => ({
          rows: rows.map((row) => {
            const parts = valueToString(row?.[idx]).split(op.delimiter);
            if (parts.length > expectedParts) {
              throw new Error(`Split produced ${parts.length} columns but only ${expectedParts} names were provided`);
            }
            const next = row.slice();
            next[idx] = parts[0] ?? "";
            for (let i = 1; i < expectedParts; i++) {
              next.push(parts[i] ?? null);
            }
            return next;
          }),
          done: false,
        }));
        break;
      }
      case "demoteHeaders": {
        const width = columns.length;
        const headerRow = columns.map((c) => c.name);
        columns = Array.from({ length: width }, (_v, i) => ({ name: `Column${i + 1}`, type: "any" }));
        columnIndex = buildIndex(columns);

        let inserted = false;
        transforms.push((rows) => {
          if (inserted) return { rows, done: false };
          inserted = true;
          return { rows: [headerRow, ...rows], done: false };
        });
        break;
      }
      default: {
        /** @type {never} */
        const exhausted = op;
        throw new Error(`Unsupported operation '${exhausted.type}'`);
      }
    }
  }

  let done = false;
  /**
   * @param {unknown[][]} rows
   */
  const transformBatch = (rows) => {
    if (done) return { rows: [], done: true };
    /** @type {unknown[][]} */
    let current = rows;
    let stop = false;
    for (const fn of transforms) {
      const result = fn(current);
      current = result.rows;
      if (result.done) stop = true;
    }
    if (stop) done = true;
    return { rows: current, done };
  };

  return { columns, transformBatch };
}

/**
 * @param {ITable} table
 * @param {number} count
 * @returns {ITable}
 */
function skip(table, count) {
  if (!Number.isFinite(count) || count < 0) {
    throw new Error(`Invalid skip count '${count}'`);
  }

  if (table instanceof ArrowTableAdapter) {
    return new ArrowTableAdapter(table.table.slice(count), table.columns);
  }

  if (table instanceof DataTable) {
    return new DataTable(table.columns, table.rows.slice(count));
  }

  const materialized = ensureDataTable(table);
  return new DataTable(materialized.columns, materialized.rows.slice(count));
}

/**
 * @param {ITable} table
 * @param {import("./model.js").RemoveRowsOp} op
 * @returns {ITable}
 */
function removeRows(table, op) {
  if (!Number.isFinite(op.offset) || op.offset < 0 || !Number.isFinite(op.count) || op.count < 0) {
    throw new Error(`Invalid removeRows range (${op.offset}, ${op.count})`);
  }
  const materialized = ensureDataTable(table);
  const start = Math.floor(op.offset);
  const end = start + Math.floor(op.count);
  const rows = materialized.rows.filter((_row, idx) => idx < start || idx >= end);
  return new DataTable(materialized.columns, rows);
}

/**
 * @param {ITable} table
 * @returns {DataTable}
 */
function ensureDataTable(table) {
  if (table instanceof DataTable) return table;

  const rows = [];
  for (const row of table.iterRows()) {
    rows.push(row);
  }
  return new DataTable(table.columns, rows);
}

/**
 * @param {unknown} value
 * @returns {value is ITable}
 */
function isITable(value) {
  if (!value || typeof value !== "object" || Array.isArray(value)) return false;
  // @ts-ignore - runtime duck-typing
  if (!("columns" in value) || !Array.isArray(value.columns)) return false;
  // @ts-ignore - runtime duck-typing
  return typeof value.getColumnIndex === "function" && typeof value.getCell === "function" && typeof value.rowCount === "number";
}

/**
 * @param {DataTable} table
 * @param {import("./model.js").ExpandTableColumnOp} op
 * @returns {DataTable}
 */
function expandTableColumn(table, op) {
  const columnIdx = table.getColumnIndex(op.column);

  let columns = op.columns ?? null;
  if (columns == null) {
    // Best-effort: infer columns from the first non-null nested table.
    for (const row of table.rows) {
      const nested = row[columnIdx];
      if (isITable(nested)) {
        columns = nested.columns.map((c) => c.name);
        break;
      }
    }
    if (columns == null) columns = [];
  }

  const newColumnNames = op.newColumnNames ?? null;
  if (newColumnNames && newColumnNames.length !== columns.length) {
    throw new Error(
      `expandTableColumn expected newColumnNames to have the same length as columns (${columns.length}), got ${newColumnNames.length}`,
    );
  }

  // Power Query expands the nested table column by inserting new columns at the position
  // of the nested column and repeating the outer row for every nested row. When the
  // nested table is empty or null, the result keeps the outer row and fills expanded
  // columns with nulls.
  const prefixColumns = table.columns.slice(0, columnIdx);
  const suffixColumns = table.columns.slice(columnIdx + 1);

  // Expand uses either the provided new column names or the nested column names.
  const rawExpandedNames = newColumnNames ?? columns;

  // Ensure we never rename existing columns; only the newly expanded columns should be uniqued.
  const reserved = new Set([...prefixColumns.map((c) => c.name), ...suffixColumns.map((c) => c.name)]);
  const baseExpanded = makeUniqueColumnNames(rawExpandedNames);
  const expandedNames = baseExpanded.map((base) => {
    if (!reserved.has(base)) {
      reserved.add(base);
      return base;
    }
    let i = 1;
    while (reserved.has(`${base}.${i}`)) i += 1;
    const unique = `${base}.${i}`;
    reserved.add(unique);
    return unique;
  });

  /** @type {ITable | null} */
  let sampleNested = null;
  for (const row of table.rows) {
    const nested = row[columnIdx];
    if (isITable(nested)) {
      sampleNested = nested;
      break;
    }
  }

  const expandedTypes = columns.map((name) => {
    if (!sampleNested) return "any";
    const idx = sampleNested.columns.findIndex((c) => c.name === name);
    return idx === -1 ? "any" : sampleNested.columns[idx]?.type ?? "any";
  });

  const outColumns = [
    ...prefixColumns,
    ...expandedNames.map((name, idx) => ({ name, type: expandedTypes[idx] ?? "any" })),
    ...suffixColumns,
  ];

  /** @type {unknown[][]} */
  const outRows = [];

  for (const row of table.rows) {
    const nestedValue = row[columnIdx];
    const prefix = row.slice(0, columnIdx);
    const suffix = row.slice(columnIdx + 1);

    if (nestedValue == null) {
      outRows.push([...prefix, ...columns.map(() => null), ...suffix]);
      continue;
    }

    if (!isITable(nestedValue)) {
      throw new Error(`expandTableColumn expected '${op.column}' to contain nested tables or null`);
    }

    const nested = nestedValue;
    const nestedIndices = columns.map((c) => nested.getColumnIndex(c));

    if (nested.rowCount === 0) {
      outRows.push([...prefix, ...columns.map(() => null), ...suffix]);
      continue;
    }

    for (let nestedRow = 0; nestedRow < nested.rowCount; nestedRow++) {
      const expanded = nestedIndices.map((idx) => nested.getCell(nestedRow, idx));
      outRows.push([...prefix, ...expanded, ...suffix]);
    }
  }

  return new DataTable(outColumns, outRows);
}

/**
 * Apply a query operation locally.
 *
 * Operations that require external state (`merge`, `append`) are handled by the
 * `QueryEngine` because they need access to other queries.
 *
 * @param {ITable} table
 * @param {QueryOperation} operation
 * @returns {ITable}
 */
export function applyOperation(table, operation) {
  switch (operation.type) {
    case "selectColumns":
      return selectColumns(table, operation.columns);
    case "removeColumns":
      return removeColumns(table, operation.columns);
    case "filterRows":
      return filterRows(table, operation.predicate);
    case "sortRows":
      return sortRows(table, operation.sortBy);
    case "groupBy":
      return groupBy(table, operation.groupColumns, operation.aggregations);
    case "addColumn":
      return addColumn(ensureDataTable(table), operation.name, operation.formula);
    case "renameColumn":
      return renameColumn(table, operation.oldName, operation.newName);
    case "changeType":
      return changeType(table, operation.column, operation.newType);
    case "distinctRows":
      return distinctRows(table, operation.columns);
    case "removeRowsWithErrors":
      return removeRowsWithErrors(table, operation.columns);
    case "transformColumns":
      return transformColumns(ensureDataTable(table), operation);
    case "take":
      return take(table, operation.count);
    case "skip":
      return skip(table, operation.count);
    case "removeRows":
      return removeRows(table, operation);
    case "pivot":
      return pivot(ensureDataTable(table), operation);
    case "unpivot":
      return unpivot(ensureDataTable(table), operation);
    case "fillDown":
      return fillDown(ensureDataTable(table), operation);
    case "replaceValues":
      return replaceValues(ensureDataTable(table), operation);
    case "splitColumn":
      return splitColumn(ensureDataTable(table), operation);
    case "expandTableColumn":
      return expandTableColumn(ensureDataTable(table), operation);
    case "promoteHeaders":
      return promoteHeaders(ensureDataTable(table));
    case "demoteHeaders":
      return demoteHeaders(ensureDataTable(table));
    case "reorderColumns":
      return reorderColumns(table, operation);
    case "addIndexColumn":
      return addIndexColumn(ensureDataTable(table), operation);
    case "combineColumns":
      return combineColumns(ensureDataTable(table), operation);
    case "transformColumnNames":
      return transformColumnNames(table, operation);
    case "replaceErrorValues":
      return replaceErrorValues(ensureDataTable(table), operation);
    case "merge":
    case "append":
      throw new Error(`Operation '${operation.type}' requires QueryEngine context`);
    default: {
      /** @type {never} */
      const exhausted = operation;
      throw new Error(`Unsupported operation '${exhausted.type}'`);
    }
  }
}
