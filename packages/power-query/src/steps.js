import { ArrowTableAdapter } from "./arrowTable.js";
import { DataTable, inferColumnType, makeUniqueColumnNames } from "./table.js";
import { compilePredicate, compileRowPredicate } from "./predicate.js";
import { valueKey } from "./valueKey.js";
import { bindExprColumns, collectExprColumnRefs, evaluateExpr, parseFormula } from "./expr/index.js";

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
      const keyValues = indices.map((idx) => valueKey(normalizeMissing(vectors[idx].get(rowIndex))));
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
    const keyValues = indices.map((idx) => valueKey(normalizeMissing(row[idx])));
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
    if (!arrowTableFromColumns) {
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
]);

/**
 * @param {QueryOperation} operation
 */
export function isStreamableOperation(operation) {
  return STREAMABLE_OPERATION_TYPES.has(operation.type);
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
    for (const fn of transforms) {
      if (done) break;
      const result = fn(current);
      current = result.rows;
      if (result.done) done = true;
      if (current.length === 0 && done) break;
    }
    return { rows: current, done };
  };

  return { columns, transformBatch };
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
    case "distinctRows":
      return distinctRows(table, operation.columns);
    case "removeRowsWithErrors":
      return removeRowsWithErrors(table, operation.columns);
    case "transformColumns":
      return transformColumns(ensureDataTable(table), operation);
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
