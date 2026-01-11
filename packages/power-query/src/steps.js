import { arrowTableFromColumns } from "../../data-io/src/index.js";
import { ArrowTableAdapter } from "./arrowTable.js";
import { DataTable, inferColumnType, makeUniqueColumnNames } from "./table.js";
import { compilePredicate } from "./predicate.js";

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
 * @param {import("./model.js").FilterPredicate} predicate
 * @returns {ITable}
 */
function filterRows(table, predicate) {
  const fn = compilePredicate(table, predicate);

  if (table instanceof ArrowTableAdapter) {
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
 * @param {unknown} value
 * @returns {unknown}
 */
function distinctKey(value) {
  if (value == null) return null;
  if (isDate(value)) return `__date__:${value.toISOString()}`;
  if (typeof value === "object") return `__json__:${JSON.stringify(value)}`;
  return value;
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
    const key = JSON.stringify(
      keyValues.map((v) => (isDate(v) ? `__date__:${v.toISOString()}` : v ?? null)),
    );

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
          state.set.add(distinctKey(value));
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
      if (isDate(value)) return value;
      if (typeof value === "number") {
        const d = new Date(value);
        return Number.isNaN(d.getTime()) ? null : d;
      }
      if (typeof value === "string") {
        const d = new Date(value);
        return Number.isNaN(d.getTime()) ? null : d;
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
 * This is intentionally limited to simple JS expressions that reference columns
 * via `[Column Name]`.
 *
 * @param {DataTable} table
 * @param {string} formula
 * @returns {(values: unknown[]) => unknown}
 */
export function compileRowFormula(table, formula) {
  let expr = formula.trim();
  if (expr.startsWith("=")) expr = expr.slice(1).trim();

  expr = expr.replaceAll(/\[([^\]]+)\]/g, (_match, rawName) => {
    const name = String(rawName).trim();
    const idx = table.getColumnIndex(name);
    return `values[${idx}]`;
  });

  // Very defensive sanitization: allow only a small subset of JS.
  if (/[{};]/.test(expr)) {
    throw new Error("Formula contains unsupported characters");
  }

  if (
    /\b(?:while|for|function|class|return|new|this|globalThis|process|require|import|eval|Function|constructor|prototype)\b/.test(
      expr,
    )
  ) {
    throw new Error("Formula contains unsupported identifiers");
  }

  if (!/^[\d\s+\-*/%().,<>=!&|?:'"[\]A-Za-z_]+$/.test(expr)) {
    throw new Error("Formula contains unsupported tokens");
  }

  // eslint-disable-next-line no-new-func
  return new Function("values", `"use strict"; return (${expr});`);
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
    if (value == null) return "__null__";
    if (isDate(value)) return `__date__:${value.toISOString()}`;
    return `__${typeof value}__:${String(value)}`;
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
    const key = JSON.stringify(
      keyValues.map((v) => (isDate(v) ? `__date__:${v.toISOString()}` : v ?? null)),
    );

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
  const rows = table.rows.map((row) => {
    const next = row.slice();
    if (Object.is(next[idx], op.find)) next[idx] = op.replace;
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
  const maxParts = Math.max(...partsByRow.map((parts) => parts.length));

  const baseNames = Array.from({ length: maxParts }, (_, i) => (i === 0 ? op.column : `${op.column}.${i + 1}`));
  const unique = makeUniqueColumnNames(baseNames);

  const columns = table.columns
    .map((col, i) => (i === idx ? { ...col, name: unique[0], type: "string" } : col))
    .concat(unique.slice(1).map((name) => ({ name, type: "string" })));

  const rows = table.rows.map((row, rIdx) => {
    const next = row.slice();
    const parts = partsByRow[rIdx];
    next[idx] = parts[0] ?? "";
    for (let i = 1; i < maxParts; i++) {
      next.push(parts[i] ?? null);
    }
    return next;
  });

  return new DataTable(columns, rows);
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
    case "take":
      return take(table, operation.count);
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
