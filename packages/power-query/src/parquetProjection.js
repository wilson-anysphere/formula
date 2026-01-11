/**
 * Parquet column projection planning.
 *
 * When reading Parquet into Arrow we can ask parquet-wasm to only decode a subset of columns.
 * This can be a major win for memory and latency on wide tables, but it is only safe when the
 * query pipeline has an explicit projection (e.g. `selectColumns` or `groupBy`) so the final
 * output does not implicitly require "all source columns".
 */

import { collectExprColumnRefs, parseFormula } from "./expr/index.js";

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
  "distinctRows",
  "removeRowsWithErrors",
  "groupBy",
  "changeType",
  "addColumn",
  "transformColumns",
  "renameColumn",
  "take",
  "skip",
  "removeRows",
  "reorderColumns",
  "addIndexColumn",
  "combineColumns",
  "replaceErrorValues",
  "fillDown",
  "replaceValues",
]);

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
  /** @type {Set<string> | null} */
  let schema = null;

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
      case "reorderColumns": {
        const missingField = op.missingField ?? "error";
        if (missingField === "error") {
          // `reorderColumns` validates column existence when missingField === "error". We must
          // read these columns even if they are later dropped by a downstream projection.
          op.columns.forEach(requireColumn);
          break;
        }

        // `missingField` modes that tolerate missing columns are only safe to project when we
        // already have a known schema from an earlier explicit projection. Otherwise we'd risk
        // requesting a column that doesn't exist in the Parquet file.
        if (!schema) return null;

        for (const name of op.columns) {
          if (schema.has(name)) requireColumn(name);
        }
        break;
      }
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
      case "distinctRows":
      case "removeRowsWithErrors": {
        if (op.columns && op.columns.length > 0) {
          op.columns.forEach(requireColumn);
        } else {
          // All-columns distinct/error checks are only safe to project when we
          // already have a known schema from an earlier explicit projection.
          if (!schema) return null;
          schema.forEach(requireColumn);
        }
        break;
      }
      case "addColumn":
        try {
          collectExprColumnRefs(parseFormula(op.formula)).forEach(requireColumn);
        } catch {
          return null;
        }
        break;
      case "transformColumns":
        op.transforms.forEach((t) => requireColumn(t.column));
        break;
      case "renameColumn":
        requireColumn(op.oldName);
        break;
      case "addIndexColumn":
        break;
      case "combineColumns":
        op.columns.forEach(requireColumn);
        break;
      case "replaceErrorValues":
        op.replacements.forEach((r) => requireColumn(r.column));
        break;
      case "skip":
      case "removeRows":
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
        if (schema) {
          schema.delete(op.oldName);
          schema.add(op.newName);
        }
        break;
      }
      case "selectColumns": {
        const next = new Map();
        for (const name of op.columns) {
          next.set(name, getSourceName(name));
        }
        mapping.clear();
        for (const [k, v] of next) mapping.set(k, v);
        schema = new Set(op.columns);
        break;
      }
      case "removeColumns": {
        for (const name of op.columns) {
          mapping.delete(name);
          schema?.delete(name);
        }
        break;
      }
      case "reorderColumns": {
        if (!schema) break;
        const missingField = op.missingField ?? "error";
        if (missingField !== "useNull") break;

        for (const name of op.columns) {
          if (schema.has(name)) continue;
          mapping.set(name, null);
          schema.add(name);
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
        schema = new Set(next.keys());
        break;
      }
      case "addColumn":
        mapping.set(op.name, null);
        schema?.add(op.name);
        break;
      case "addIndexColumn":
        mapping.set(op.name, null);
        schema?.add(op.name);
        break;
      case "combineColumns": {
        for (const name of op.columns) {
          mapping.delete(name);
          schema?.delete(name);
        }
        mapping.set(op.newColumnName, null);
        schema?.add(op.newColumnName);
        break;
      }
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
  "merge",
  "append",
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
  const base = Number.isFinite(limit) ? Math.max(0, Math.trunc(limit)) : 0;
  if (base <= 0) return null;

  // Refuse to push down a reader limit when a step could require reading beyond the first
  // N source rows to produce the first N output rows (e.g. filter/sort/group/pivot).
  for (const step of steps) {
    const op = step.operation;
    if (LIMIT_UNSAFE_OPS.has(op.type)) return null;
  }

  // Walk backwards from the final output limit to compute how many source rows are required
  // to produce the first N output rows.
  //
  // This allows safe limit pushdown through deterministic row-window operations like
  // `take`, `skip`, `removeRows`, and `promoteHeaders`.
  let required = base;
  for (let i = steps.length - 1; i >= 0; i--) {
    /** @type {QueryOperation} */
    const op = steps[i].operation;
    switch (op.type) {
      case "take": {
        if (!Number.isFinite(op.count) || op.count < 0) return null;
        required = Math.min(required, Math.max(0, Math.trunc(op.count)));
        break;
      }
      case "skip": {
        if (!Number.isFinite(op.count) || op.count < 0) return null;
        if (required > 0) {
          required += Math.max(0, Math.trunc(op.count));
        }
        break;
      }
      case "removeRows": {
        if (!Number.isFinite(op.offset) || op.offset < 0 || !Number.isFinite(op.count) || op.count < 0) return null;
        const offset = Math.floor(op.offset);
        const count = Math.floor(op.count);
        // `removeRows` drops a contiguous window starting at `offset`.
        // If the downstream limit is entirely before that window, no adjustment is needed.
        if (required > offset) required += count;
        break;
      }
      case "promoteHeaders": {
        // `promoteHeaders` consumes the first data row as a header row (dropping it from the output).
        // To preserve output limit semantics we need one extra source row when at least one row is requested.
        if (required > 0) required += 1;
        break;
      }
      default:
        break;
    }
  }

  return required;
}
