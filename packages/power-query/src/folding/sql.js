import { predicateToSql, quoteIdentifier } from "../predicate.js";

/**
 * @typedef {import("../model.js").Query} Query
 * @typedef {import("../model.js").QueryStep} QueryStep
 * @typedef {import("../model.js").QueryOperation} QueryOperation
 * @typedef {import("../model.js").SortSpec} SortSpec
 * @typedef {import("../model.js").Aggregation} Aggregation
 */

/**
 * @typedef {{
 *   type: "local";
 *   steps: QueryStep[];
 * }} LocalPlan
 *
 * @typedef {{
 *   type: "sql";
 *   sql: string;
 * }} SqlPlan
 *
 * @typedef {{
 *   type: "hybrid";
 *   sql: string;
 *   localSteps: QueryStep[];
 * }} HybridPlan
 *
 * @typedef {LocalPlan | SqlPlan | HybridPlan} CompiledQueryPlan
 */

/**
 * A conservative SQL query folding engine. It folds a prefix of operations to
 * SQL for `database` sources and returns a hybrid plan when folding breaks.
 *
 * The folding strategy is intentionally "dumb but correct": it wraps the
 * previous query in a subquery at every step instead of trying to flatten.
 */
export class QueryFoldingEngine {
  constructor() {
    /** @type {Set<string>} */
    this.foldable = new Set(["selectColumns", "filterRows", "sortRows", "groupBy"]);
  }

  /**
   * @param {Query} query
   * @returns {CompiledQueryPlan}
   */
  compile(query) {
    if (query.source.type !== "database") {
      return { type: "local", steps: query.steps };
    }

    let currentSql = query.source.query;
    /** @type {QueryStep[]} */
    const localSteps = [];
    let foldingBroken = false;

    for (const step of query.steps) {
      if (foldingBroken) {
        localSteps.push(step);
        continue;
      }

      if (!this.foldable.has(step.operation.type)) {
        foldingBroken = true;
        localSteps.push(step);
        continue;
      }

      const next = this.applySqlStep(currentSql, step.operation);
      if (!next) {
        foldingBroken = true;
        localSteps.push(step);
        continue;
      }
      currentSql = next;
    }

    if (localSteps.length === 0) {
      return { type: "sql", sql: currentSql };
    }
    return { type: "hybrid", sql: currentSql, localSteps };
  }

  /**
   * @param {string} sql
   * @param {QueryOperation} operation
   * @returns {string | null}
   */
  applySqlStep(sql, operation) {
    const from = `(${sql}) AS t`;
    switch (operation.type) {
      case "selectColumns": {
        const cols = operation.columns.map((c) => `t.${quoteIdentifier(c)}`).join(", ");
        return `SELECT ${cols} FROM ${from}`;
      }
      case "filterRows": {
        const where = predicateToSql(operation.predicate, { alias: "t" });
        return `SELECT * FROM ${from} WHERE ${where}`;
      }
      case "sortRows": {
        const orderBy = sortSpecsToSql(operation.sortBy);
        return `SELECT * FROM ${from} ORDER BY ${orderBy}`;
      }
      case "groupBy": {
        const groupCols = operation.groupColumns.map((c) => `t.${quoteIdentifier(c)}`).join(", ");
        const aggSql = operation.aggregations.map((agg) => aggregationToSql(agg)).join(", ");
        const selectList = [groupCols, aggSql].filter(Boolean).join(", ");
        const groupByClause = groupCols ? ` GROUP BY ${groupCols}` : "";
        return `SELECT ${selectList} FROM ${from}${groupByClause}`;
      }
      default:
        return null;
    }
  }
}

/**
 * @param {SortSpec[]} specs
 * @returns {string}
 */
function sortSpecsToSql(specs) {
  return specs
    .map((spec) => {
      const dir = spec.direction === "descending" ? "DESC" : "ASC";
      const nulls = spec.nulls ? ` NULLS ${spec.nulls.toUpperCase()}` : "";
      return `t.${quoteIdentifier(spec.column)} ${dir}${nulls}`;
    })
    .join(", ");
}

/**
 * @param {Aggregation} agg
 * @returns {string}
 */
function aggregationToSql(agg) {
  const alias = quoteIdentifier(agg.as ?? `${agg.op} of ${agg.column}`);
  const col = `t.${quoteIdentifier(agg.column)}`;
  switch (agg.op) {
    case "sum":
      return `SUM(${col}) AS ${alias}`;
    case "count":
      return `COUNT(*) AS ${alias}`;
    case "average":
      return `AVG(${col}) AS ${alias}`;
    case "min":
      return `MIN(${col}) AS ${alias}`;
    case "max":
      return `MAX(${col}) AS ${alias}`;
    case "countDistinct":
      return `COUNT(DISTINCT ${col}) AS ${alias}`;
    default: {
      /** @type {never} */
      const exhausted = agg.op;
      throw new Error(`Unsupported aggregation '${exhausted}'`);
    }
  }
}

