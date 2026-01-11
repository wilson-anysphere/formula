/**
 * Parquet column projection planning.
 *
 * When reading Parquet into Arrow we can ask parquet-wasm to only decode a subset of columns.
 * This can be a major win for memory and latency on wide tables, but it is only safe when the
 * query pipeline has an explicit projection (e.g. `selectColumns` or `groupBy`) so the final
 * output does not implicitly require "all source columns".
 */

/**
 * @typedef {import("./model.js").QueryStep} QueryStep
 * @typedef {import("./model.js").QueryOperation} QueryOperation
 * @typedef {import("./model.js").FilterPredicate} FilterPredicate
 */

const SUPPORTED_OPS = new Set([
  "selectColumns",
  "removeColumns",
  "filterRows",
  "sortRows",
  "groupBy",
  "changeType",
  "addColumn",
  "renameColumn",
  "take",
  "fillDown",
  "replaceValues",
]);

/**
 * Extract bracketed column references (`[Column Name]`) from an `addColumn` formula.
 *
 * This intentionally mirrors `compileRowFormula`'s very small supported surface area.
 *
 * @param {string} formula
 * @returns {Set<string>}
 */
function parseFormulaColumnRefs(formula) {
  const refs = new Set();
  let expr = formula.trim();
  if (expr.startsWith("=")) expr = expr.slice(1).trim();
  for (const match of expr.matchAll(/\[([^\]]+)\]/g)) {
    const name = String(match[1]).trim();
    if (name) refs.add(name);
  }
  return refs;
}

/**
 * @param {FilterPredicate} predicate
 * @param {Set<string>} out
 */
function collectPredicateColumns(predicate, out) {
  switch (predicate.type) {
    case "comparison":
      out.add(predicate.column);
      return;
    case "and":
    case "or":
      predicate.predicates.forEach((p) => collectPredicateColumns(p, out));
      return;
    case "not":
      collectPredicateColumns(predicate.predicate, out);
      return;
    default: {
      /** @type {never} */
      const exhausted = predicate;
      throw new Error(`Unsupported predicate type '${exhausted.type}'`);
    }
  }
}

/**
 * Compute a set of Parquet source columns to request via `parquet-wasm`'s `ReaderOptions.columns`.
 *
 * Returns `null` when the pipeline could still require all source columns (e.g. a query that only
 * filters/sorts but never projects).
 *
 * @param {QueryStep[]} steps
 * @returns {string[] | null}
 */
export function computeParquetProjectionColumns(steps) {
  const hasExplicitProjection = steps.some(
    (step) => step.operation.type === "selectColumns" || step.operation.type === "groupBy",
  );
  if (!hasExplicitProjection) return null;

  for (const step of steps) {
    if (!SUPPORTED_OPS.has(step.operation.type)) {
      return null;
    }
  }

  /** @type {Map<string, string | null>} current column name -> parquet column name (null means derived) */
  const mapping = new Map();
  /** @type {Set<string>} */
  const required = new Set();

  /**
   * @param {string} name
   */
  const getSourceName = (name) => {
    if (!mapping.has(name)) return name;
    return mapping.get(name) ?? null;
  };

  /**
   * @param {string} name
   */
  const requireColumn = (name) => {
    const source = getSourceName(name);
    if (source != null) required.add(source);
  };

  for (const step of steps) {
    /** @type {QueryOperation} */
    const op = step.operation;

    switch (op.type) {
      case "selectColumns":
        op.columns.forEach(requireColumn);
        break;
      case "removeColumns":
        // `removeColumns` validates column existence, so we must still read these columns.
        op.columns.forEach(requireColumn);
        break;
      case "filterRows": {
        const cols = new Set();
        collectPredicateColumns(op.predicate, cols);
        cols.forEach(requireColumn);
        break;
      }
      case "sortRows":
        op.sortBy.forEach((spec) => requireColumn(spec.column));
        break;
      case "groupBy":
        op.groupColumns.forEach(requireColumn);
        op.aggregations.forEach((agg) => requireColumn(agg.column));
        break;
      case "changeType":
        requireColumn(op.column);
        break;
      case "addColumn":
        parseFormulaColumnRefs(op.formula).forEach(requireColumn);
        break;
      case "renameColumn":
        requireColumn(op.oldName);
        break;
      case "fillDown":
        op.columns.forEach(requireColumn);
        break;
      case "replaceValues":
        requireColumn(op.column);
        break;
      case "take":
        break;
      default: {
        /** @type {never} */
        const exhausted = op;
        throw new Error(`Unsupported operation '${exhausted.type}'`);
      }
    }

    // Update the name->source mapping as we walk forward.
    switch (op.type) {
      case "renameColumn": {
        const sourceName = getSourceName(op.oldName);
        mapping.delete(op.oldName);
        mapping.set(op.newName, sourceName);
        break;
      }
      case "selectColumns": {
        const next = new Map();
        for (const name of op.columns) {
          next.set(name, getSourceName(name));
        }
        mapping.clear();
        for (const [k, v] of next) mapping.set(k, v);
        break;
      }
      case "removeColumns": {
        for (const name of op.columns) {
          mapping.delete(name);
        }
        break;
      }
      case "groupBy": {
        const next = new Map();
        for (const name of op.groupColumns) {
          next.set(name, getSourceName(name));
        }
        for (const agg of op.aggregations) {
          const outName = agg.as ?? `${agg.op} of ${agg.column}`;
          next.set(outName, getSourceName(agg.column));
        }
        mapping.clear();
        for (const [k, v] of next) mapping.set(k, v);
        break;
      }
      case "addColumn":
        mapping.set(op.name, null);
        break;
      default:
        break;
    }
  }

  return Array.from(required);
}

const LIMIT_UNSAFE_OPS = new Set([
  "filterRows",
  "sortRows",
  "distinctRows",
  "removeRowsWithErrors",
  "groupBy",
  "pivot",
  "unpivot",
  "splitColumn",
]);

/**
 * Compute a safe Parquet reader `limit` value to push down.
 *
 * This is safe only when the pipeline preserves row order and does not require inspecting rows
 * beyond the first N to compute the first N output rows.
 *
 * @param {QueryStep[]} steps
 * @param {number | undefined} limit
 * @returns {number | null}
 */
export function computeParquetRowLimit(steps, limit) {
  if (limit == null) return null;
  if (!Number.isFinite(limit) || limit <= 0) return null;

  let effective = limit;
  for (const step of steps) {
    const op = step.operation;
    if (LIMIT_UNSAFE_OPS.has(op.type)) return null;
    if (op.type === "take") {
      effective = Math.min(effective, op.count);
    }
  }

  return effective;
}
