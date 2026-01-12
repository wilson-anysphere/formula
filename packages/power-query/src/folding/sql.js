import { predicateToSql } from "../predicate.js";
import { hashValue } from "../cache/key.js";
import { makeUniqueColumnNames } from "../table.js";
import { POSTGRES_DIALECT, getSqlDialect } from "./dialect.js";
import { compileExprToSql, parseFormula } from "../expr/index.js";
import { collectSourcePrivacy, distinctPrivacyLevels } from "../privacy/firewall.js";
import { getPrivacyLevel } from "../privacy/levels.js";
import { getSqlSourceId } from "../privacy/sourceId.js";
import { PqDateTimeZone, PqDecimal, PqDuration, PqTime } from "../values.js";

/**
 * @typedef {import("../model.js").Query} Query
 * @typedef {import("../model.js").QueryStep} QueryStep
 * @typedef {import("../model.js").QueryOperation} QueryOperation
 * @typedef {import("../model.js").Aggregation} Aggregation
 * @typedef {import("../model.js").DataType} DataType
 * @typedef {import("./dialect.js").SqlDialect} SqlDialect
 * @typedef {import("./dialect.js").SqlDialectName} SqlDialectName
 */

/**
 * @typedef {{
 *   type: "local";
 *   steps: QueryStep[];
 *   diagnostics?: FoldingFirewallDiagnostic[];
 * }} LocalPlan
 *
 * @typedef {{
 *   type: "sql";
 *   sql: string;
 *   params: unknown[];
 *   diagnostics?: FoldingFirewallDiagnostic[];
 * }} SqlPlan
 *
 * @typedef {{
 *   type: "hybrid";
 *   sql: string;
 *   params: unknown[];
 *   localSteps: QueryStep[];
 *   diagnostics?: FoldingFirewallDiagnostic[];
 * }} HybridPlan
 *
 * @typedef {LocalPlan | SqlPlan | HybridPlan} CompiledQueryPlan
 */

/**
 * @typedef {{
 *   stepId: string;
 *   opType: QueryOperation["type"];
 *   status: "folded" | "local";
 *   sqlFragment?: string;
 *   reason?: string;
 * }} FoldingExplainStep
 */

/**
 * @typedef {{
 *   plan: CompiledQueryPlan;
 *   steps: FoldingExplainStep[];
 * }} FoldingExplainResult
 */

/**
 * @typedef {{
 *   sql: string;
 *   params: unknown[];
 * }} SqlFragment
 */

/**
 * @typedef {{
 *   fragment: SqlFragment;
 *   // When known, output column ordering + names. Used for conservative folding
 *   // of operations that need an explicit projection (rename/changeType/merge).
  *   columns: string[] | null;
  *   // SQL Server does not allow `ORDER BY` in derived tables unless paired with
  *   // `TOP`/`OFFSET`. Because this folding engine wraps each step in a subquery,
  *   // we track the current sort spec separately and only emit the final `ORDER BY`
  *   // at the outermost level (or alongside `TOP`).
  *   sortBy?: import("../model.js").SortSpec[] | null;
  *   // Whether `fragment.sql` currently includes a top-level `ORDER BY` that matches
  *   // `sortBy` (only relevant for SQL Server, where `TOP` queries embed ordering).
  *   sortInFragment?: boolean;
  *   // Connection identity used to ensure we only fold `merge`/`append` when both
  *   // sides originate from the same database connection.
  *   connectionId: string | null;
  *   // Original connection descriptor used for backwards-compatible fallback
  *   // when no stable identity is available.
  *   connection: unknown;
  * }} SqlState
  */

/**
 * @typedef {{
 *   dialect?: SqlDialect | SqlDialectName;
 *   // Queries are required to fold operations like `merge`, `append`, and
 *   // sources of type `query`.
 *   queries?: Record<string, Query>;
 *   getConnectionIdentity?: (connection: unknown) => unknown;
 *   privacyMode?: "ignore" | "enforce" | "warn";
 *   privacyLevelsBySourceId?: Record<string, import("../privacy/levels.js").PrivacyLevel>;
 * }} CompileOptions
 */

/**
 * @typedef {{
 *   kind: "privacy:firewall";
 *   phase: "folding";
 *   operation: "merge" | "append";
 *   sources: Array<{ sourceId: string; level: import("../privacy/levels.js").PrivacyLevel }>;
 *   message: string;
 * }} FoldingFirewallDiagnostic
 */

/**
 * A conservative SQL query folding engine. It folds a prefix of operations to
 * SQL for `database` sources and returns a hybrid plan when folding breaks.
 *
 * The folding strategy is intentionally "dumb but correct": it wraps the
 * previous query in a subquery at every step instead of trying to flatten.
 */
export class QueryFoldingEngine {
  /**
   * @param {{
   *   dialect?: SqlDialect | SqlDialectName;
   *   queries?: Record<string, Query>;
   * }} [options]
   */
  constructor(options = {}) {
    /** @type {SqlDialect} */
    this.dialect = resolveDialect(options.dialect);
    /** @type {boolean} */
    this.dialectExplicit = options.dialect != null;
    /** @type {Record<string, Query> | null} */
    this.queries = options.queries ?? null;

    /** @type {Set<string>} */
    this.foldable = new Set([
      "selectColumns",
      "removeColumns",
      "filterRows",
      "sortRows",
      "distinctRows",
      "groupBy",
      "renameColumn",
      "changeType",
      "transformColumns",
      "addColumn",
      "merge",
      "append",
      "take",
      "skip",
    ]);
  }

  /**
   * @param {Query} query
   * @param {CompileOptions} [options]
   * @returns {CompiledQueryPlan}
   */
  compile(query, options = {}) {
    const dialect = resolveDialect(options.dialect ?? this.dialect);
    const queries = options.queries ?? this.queries;
    const getConnectionIdentity = options.getConnectionIdentity ?? null;
    /** @type {FoldingFirewallDiagnostic[]} */
    const diagnostics = [];
    const ctx = {
      dialect,
      queries,
      getConnectionIdentity,
      privacyMode: options.privacyMode ?? "ignore",
      privacyLevelsBySourceId: options.privacyLevelsBySourceId,
      diagnostics,
    };

    const callStack = new Set([query.id]);
    const initial = this.compileSourceToSqlState(query.source, ctx, callStack);
    if (!initial) {
      return diagnostics.length > 0 ? { type: "local", steps: query.steps, diagnostics } : { type: "local", steps: query.steps };
    }

    /** @type {SqlState} */
    let current = initial;
    /** @type {QueryStep[]} */
    const localSteps = [];
    let foldingBroken = false;

    for (let idx = 0; idx < query.steps.length; idx++) {
      const step = query.steps[idx];

      if (foldingBroken) {
        localSteps.push(step);
        continue;
      }

      // Special case: `Table.NestedJoin` followed immediately by `Table.ExpandTableColumn`
      // is equivalent to a flattened join, so we can fold the pair into SQL even
      // though nested joins are not foldable on their own.
      if (step.operation.type === "merge" && (step.operation.joinMode ?? "flat") === "nested") {
        const expandStep = query.steps[idx + 1] ?? null;
        if (expandStep?.operation.type === "expandTableColumn" && expandStep.operation.column === step.operation.newColumnName) {
          const folded = this.applySqlNestedJoinExpand(current, step.operation, expandStep.operation, ctx, callStack);
          if (folded) {
            current = folded;
            idx += 1; // consume the expand step too
            continue;
          }
        }
      }

      if (!this.foldable.has(step.operation.type)) {
        foldingBroken = true;
        localSteps.push(step);
        continue;
      }

      const next = this.applySqlStep(current, step.operation, ctx, callStack);
      if (!next) {
        foldingBroken = true;
        localSteps.push(step);
        continue;
      }
      current = next;
    }

    if (localSteps.length === 0) {
      const sql = finalizeSqlForDialect(current, dialect);
      return diagnostics.length > 0
        ? { type: "sql", sql, params: current.fragment.params, diagnostics }
        : { type: "sql", sql, params: current.fragment.params };
    }
    const sql = finalizeSqlForDialect(current, dialect);
    return diagnostics.length > 0
      ? { type: "hybrid", sql, params: current.fragment.params, localSteps, diagnostics }
      : { type: "hybrid", sql, params: current.fragment.params, localSteps };
  }

  /**
   * Explain folding decisions for a query.
   *
   * The folding engine only attempts to generate SQL when a dialect is known.
   * `compile()` defaults to Postgres for backwards compatibility, but `explain()`
   * is intentionally conservative and returns `missing_dialect` when the dialect
   * isn't explicitly provided.
   *
   * @param {Query} query
   * @param {CompileOptions} [options]
   * @returns {FoldingExplainResult}
   */
  explain(query, options = {}) {
    const queries = options.queries ?? this.queries;
    const getConnectionIdentity = options.getConnectionIdentity ?? null;

    const dialectInput =
      options.dialect ??
      (query.source.type === "database" ? query.source.dialect : undefined) ??
      (this.dialectExplicit ? this.dialect : undefined);

    if (!dialectInput) {
      return {
        plan: { type: "local", steps: query.steps },
        steps: query.steps.map((step) => ({
          stepId: step.id,
          opType: step.operation.type,
          status: "local",
          reason: "missing_dialect",
        })),
      };
    }

    const dialect = resolveDialect(dialectInput);
    /** @type {FoldingFirewallDiagnostic[]} */
    const diagnostics = [];
    const ctx = {
      dialect,
      queries,
      getConnectionIdentity,
      privacyMode: options.privacyMode ?? "ignore",
      privacyLevelsBySourceId: options.privacyLevelsBySourceId,
      diagnostics,
    };

    const callStack = new Set([query.id]);
    const initial = this.compileSourceToSqlState(query.source, ctx, callStack);
    if (!initial) {
      const reason = explainSourceFailure(query.source, { queries, callStack, dialect });
      return {
        plan: { type: "local", steps: query.steps },
        steps: query.steps.map((step) => ({
          stepId: step.id,
          opType: step.operation.type,
          status: "local",
          reason,
        })),
      };
    }

    /** @type {SqlState} */
    let current = initial;
    /** @type {QueryStep[]} */
    const localSteps = [];
    let foldingBroken = false;
    /** @type {string | undefined} */
    let stopReason;
    /** @type {FoldingExplainStep[]} */
    const steps = [];

    for (let idx = 0; idx < query.steps.length; idx++) {
      const step = query.steps[idx];

      if (foldingBroken) {
        localSteps.push(step);
        steps.push({
          stepId: step.id,
          opType: step.operation.type,
          status: "local",
          reason: "folding_stopped",
        });
        continue;
      }

      // Same special-case as `compile()`: fold nested join + expand as a single SQL join.
      if (step.operation.type === "merge" && (step.operation.joinMode ?? "flat") === "nested") {
        const expandStep = query.steps[idx + 1] ?? null;
        if (expandStep?.operation.type === "expandTableColumn" && expandStep.operation.column === step.operation.newColumnName) {
          const folded = this.applySqlNestedJoinExpand(current, step.operation, expandStep.operation, ctx, callStack);
          if (folded) {
            current = folded;
            const sqlFragment = finalizeSqlForDialect(current, dialect);
            steps.push({
              stepId: step.id,
              opType: step.operation.type,
              status: "folded",
              sqlFragment,
            });
            steps.push({
              stepId: expandStep.id,
              opType: expandStep.operation.type,
              status: "folded",
              sqlFragment,
            });
            idx += 1;
            continue;
          }
        }
      }

      if (!this.foldable.has(step.operation.type)) {
        foldingBroken = true;
        stopReason = "unsupported_op";
        localSteps.push(step);
        steps.push({
          stepId: step.id,
          opType: step.operation.type,
          status: "local",
          reason: stopReason,
        });
        continue;
      }

      const next = this.applySqlStep(current, step.operation, ctx, callStack);
      if (!next) {
        foldingBroken = true;
        stopReason = this.explainSqlStepFailure(current, step.operation, ctx, callStack);
        localSteps.push(step);
        steps.push({
          stepId: step.id,
          opType: step.operation.type,
          status: "local",
          reason: stopReason,
        });
        continue;
      }

      current = next;
      steps.push({
        stepId: step.id,
        opType: step.operation.type,
        status: "folded",
        sqlFragment: finalizeSqlForDialect(current, dialect),
      });
    }

    if (localSteps.length === 0) {
      const sql = finalizeSqlForDialect(current, dialect);
      return {
        plan:
          diagnostics.length > 0
            ? { type: "sql", sql, params: current.fragment.params, diagnostics }
            : { type: "sql", sql, params: current.fragment.params },
        steps,
      };
    }
    const sql = finalizeSqlForDialect(current, dialect);
    return {
      plan:
        diagnostics.length > 0
          ? { type: "hybrid", sql, params: current.fragment.params, localSteps, diagnostics }
          : { type: "hybrid", sql, params: current.fragment.params, localSteps },
      steps,
    };
  }

  /**
   * Try to compile an entire query to SQL. Used when folding operations depend
   * on other queries (e.g. `merge` + `append`).
   *
   * @param {Query} query
   * @param {{
   *   dialect: SqlDialect;
   *   queries?: Record<string, Query> | null;
   *   getConnectionIdentity?: ((connection: unknown) => unknown) | null;
   *   privacyMode?: "ignore" | "enforce" | "warn";
   *   privacyLevelsBySourceId?: Record<string, import("../privacy/levels.js").PrivacyLevel>;
   *   diagnostics?: FoldingFirewallDiagnostic[];
   * }} ctx
   * @param {Set<string>} callStack
   * @returns {SqlState | null}
   */
  compileQueryToSqlState(query, ctx, callStack) {
    if (callStack.has(query.id)) return null;
    const nextStack = new Set(callStack);
    nextStack.add(query.id);

    const initial = this.compileSourceToSqlState(query.source, ctx, nextStack);
    if (!initial) return null;

    /** @type {SqlState} */
    let current = initial;
    for (let idx = 0; idx < query.steps.length; idx++) {
      const step = query.steps[idx];

      if (step.operation.type === "merge" && (step.operation.joinMode ?? "flat") === "nested") {
        const expandStep = query.steps[idx + 1] ?? null;
        if (expandStep?.operation.type === "expandTableColumn" && expandStep.operation.column === step.operation.newColumnName) {
          const folded = this.applySqlNestedJoinExpand(current, step.operation, expandStep.operation, ctx, nextStack);
          if (folded) {
            current = folded;
            idx += 1;
            continue;
          }
        }
      }

      if (!this.foldable.has(step.operation.type)) return null;
      const next = this.applySqlStep(current, step.operation, ctx, nextStack);
      if (!next) return null;
      current = next;
    }
    return current;
  }

  /**
   * Fold `Table.NestedJoin` + `Table.ExpandTableColumn` into a single SQL join.
   *
   * Nested joins themselves do not translate to SQL because they yield nested
   * tables as cell values. When the next step immediately expands the nested
   * column, the combined operation is equivalent to a flattened join.
   *
   * @private
   * @param {SqlState} state
   * @param {import("../model.js").MergeOp} merge
   * @param {import("../model.js").ExpandTableColumnOp} expand
   * @param {{
   *   dialect: SqlDialect;
   *   queries?: Record<string, Query> | null;
   *   getConnectionIdentity?: ((connection: unknown) => unknown) | null;
   *   privacyMode?: "ignore" | "enforce" | "warn";
   *   privacyLevelsBySourceId?: Record<string, import("../privacy/levels.js").PrivacyLevel>;
   *   diagnostics?: FoldingFirewallDiagnostic[];
   * }} ctx
   * @param {Set<string>} callStack
   * @returns {SqlState | null}
   */
  applySqlNestedJoinExpand(state, merge, expand, ctx, callStack) {
    const dialect = ctx.dialect;
    const quoteIdentifier = dialect.quoteIdentifier;
    if (!state.columns) return null;

    const rightQuery = ctx.queries?.[merge.rightQuery];
    if (!rightQuery) return null;
    const rightState = this.compileQueryToSqlState(rightQuery, ctx, callStack);
    if (!rightState?.columns) return null;

    const leftSourceId = sqlSourceIdForState(state);
    const rightSourceId = sqlSourceIdForState(rightState);
    if (!connectionsMatch(state, rightState)) {
      recordFoldingPrivacyDiagnostic(ctx, "merge", [leftSourceId, rightSourceId]);
      return null;
    }

    if (ctx.privacyMode && ctx.privacyMode !== "ignore") {
      const levelsBySourceId = ctx.privacyLevelsBySourceId;
      const leftLevel = getPrivacyLevel(levelsBySourceId, leftSourceId);
      const rightLevel = getPrivacyLevel(levelsBySourceId, rightSourceId);
      if (leftLevel !== rightLevel) {
        recordFoldingPrivacyDiagnostic(ctx, "merge", [leftSourceId, rightSourceId]);
        return null;
      }
    }

    if (!isSqlFoldableJoinComparer(effectiveJoinComparer(merge))) return null;

    const join = joinTypeToSql(dialect, merge.joinType);
    if (!join) return null;

    const leftKeys =
      Array.isArray(merge.leftKeys) && merge.leftKeys.length > 0
        ? merge.leftKeys
        : typeof merge.leftKey === "string" && merge.leftKey
          ? [merge.leftKey]
          : [];
    const rightKeys =
      Array.isArray(merge.rightKeys) && merge.rightKeys.length > 0
        ? merge.rightKeys
        : typeof merge.rightKey === "string" && merge.rightKey
          ? [merge.rightKey]
          : [];

    if (leftKeys.length === 0 || rightKeys.length === 0) return null;
    if (leftKeys.length !== rightKeys.length) return null;
    for (const key of leftKeys) {
      if (!state.columns.includes(key)) return null;
    }
    for (const key of rightKeys) {
      if (!rightState.columns.includes(key)) return null;
    }

    // Nested table schema is either explicitly projected by the merge op or the full right schema.
    const nestedSchema = Array.isArray(merge.rightColumns) ? merge.rightColumns : rightState.columns;
    if (!nestedSchema) return null;

    /** @type {string[] | null} */
    let expandedColumns = Array.isArray(expand.columns) ? expand.columns : null;
    if (!expandedColumns) {
      // Match local semantics: `columns: null` expands all columns from the first
      // non-null nested table, which for nested joins is the nested schema.
      expandedColumns = nestedSchema.slice();
    }

    const newColumnNames = Array.isArray(expand.newColumnNames) ? expand.newColumnNames : null;
    if (newColumnNames && newColumnNames.length !== expandedColumns.length) return null;

    for (const col of expandedColumns) {
      if (!nestedSchema.includes(col)) return null;
      if (!rightState.columns.includes(col)) return null;
    }

    const rawExpandedNames = newColumnNames ?? expandedColumns;
    const expandedOutNames = computeExpandedColumnNames(state.columns, rawExpandedNames);

    const selectList = [
      ...state.columns.map((name) => `l.${quoteIdentifier(name)} AS ${quoteIdentifier(name)}`),
      ...expandedColumns.map((name, idx) => `r.${quoteIdentifier(name)} AS ${quoteIdentifier(expandedOutNames[idx])}`),
    ].join(", ");

    const on = leftKeys
      .map((leftKey, idx) =>
        nullSafeEqualsSql(dialect, `l.${quoteIdentifier(leftKey)}`, `r.${quoteIdentifier(rightKeys[idx])}`),
      )
      .join(" AND ");
    const sql = `SELECT ${selectList} FROM (${state.fragment.sql}) AS l ${join} (${rightState.fragment.sql}) AS r ON ${on}`;

    return {
      fragment: { sql, params: [...state.fragment.params, ...rightState.fragment.params] },
      columns: [...state.columns, ...expandedOutNames],
      sortBy: null,
      sortInFragment: false,
      connectionId: state.connectionId,
      connection: state.connection,
    };
  }

  /**
   * @param {import("../model.js").QuerySource} source
   * @param {{
   *   dialect: SqlDialect;
   *   queries?: Record<string, Query> | null;
   *   getConnectionIdentity?: ((connection: unknown) => unknown) | null;
   *   privacyMode?: "ignore" | "enforce" | "warn";
   *   privacyLevelsBySourceId?: Record<string, import("../privacy/levels.js").PrivacyLevel>;
   *   diagnostics?: FoldingFirewallDiagnostic[];
   * }} ctx
   * @param {Set<string>} callStack
   * @returns {SqlState | null}
   */
  compileSourceToSqlState(source, ctx, callStack) {
    switch (source.type) {
      case "database": {
        const columns = source.columns ? source.columns.slice() : null;
        const connectionId = resolveConnectionId(source, ctx.getConnectionIdentity);
        if (ctx.dialect.name === "sqlserver" && !isSqlServerDerivedTableSafe(source.query)) {
          return null;
        }
        return {
          fragment: { sql: source.query, params: [] },
          columns,
          sortBy: null,
          sortInFragment: false,
          connectionId,
          connection: source.connection,
        };
      }
      case "query": {
        const target = ctx.queries?.[source.queryId];
        if (!target) return null;
        if (callStack.has(target.id)) return null;
        return this.compileQueryToSqlState(target, ctx, callStack);
      }
      default:
        return null;
    }
  }

  /**
   * @param {SqlState} state
   * @param {QueryOperation} operation
   * @param {{
   *   dialect: SqlDialect;
   *   queries?: Record<string, Query> | null;
   *   getConnectionIdentity?: ((connection: unknown) => unknown) | null;
   *   privacyMode?: "ignore" | "enforce" | "warn";
   *   privacyLevelsBySourceId?: Record<string, import("../privacy/levels.js").PrivacyLevel>;
   *   diagnostics?: FoldingFirewallDiagnostic[];
   * }} ctx
   * @param {Set<string>} callStack
   * @returns {SqlState | null}
   */
  applySqlStep(state, operation, ctx, callStack) {
    const dialect = ctx.dialect;
    const quoteIdentifier = dialect.quoteIdentifier;
    const params = state.fragment.params.slice();
    const param = makeParam(dialect, params);
    const sortBy = state.sortBy ?? null;

    const from = `(${state.fragment.sql}) AS t`;
    switch (operation.type) {
      case "selectColumns": {
        if (operation.columns.length === 0) return null;
        if (hasDuplicateStrings(operation.columns)) return null;
        if (dialect.name === "sqlserver" && sortBy && sortBy.some((spec) => !operation.columns.includes(spec.column))) {
          return null;
        }
        const cols = operation.columns.map((c) => `t.${quoteIdentifier(c)}`).join(", ");
        return {
          fragment: { sql: `SELECT ${cols} FROM ${from}`, params },
          columns: operation.columns.slice(),
          sortBy,
          sortInFragment: false,
          connectionId: state.connectionId,
          connection: state.connection,
        };
      }
      case "removeColumns": {
        if (!state.columns) return null;
        const remove = new Set(operation.columns);
        for (const name of operation.columns) {
          if (!state.columns.includes(name)) return null;
        }
        if (dialect.name === "sqlserver" && sortBy && sortBy.some((spec) => remove.has(spec.column))) {
          return null;
        }

        const remaining = state.columns.filter((name) => !remove.has(name));
        if (remaining.length === 0) return null;
        const cols = remaining.map((c) => `t.${quoteIdentifier(c)}`).join(", ");
        return {
          fragment: { sql: `SELECT ${cols} FROM ${from}`, params },
          columns: remaining,
          sortBy,
          sortInFragment: false,
          connectionId: state.connectionId,
          connection: state.connection,
        };
      }
      case "filterRows": {
        if (!predicateHasOnlySqlScalarValues(operation.predicate)) return null;
        const where = predicateToSql(operation.predicate, {
          alias: "t",
          quoteIdentifier,
          // Local filter semantics stringify null/undefined to "" before applying
          // contains/startsWith/endsWith. Use COALESCE to preserve that behavior
          // for LIKE-based predicates.
          castText: (expr) => `COALESCE(${dialect.castText(expr)}, '')`,
          param,
        });
        return {
          fragment: { sql: `SELECT * FROM ${from} WHERE ${where}`, params },
          columns: state.columns,
          sortBy,
          sortInFragment: false,
          connectionId: state.connectionId,
          connection: state.connection,
        };
      }
      case "sortRows": {
        if (operation.sortBy.length === 0) {
          return {
            fragment: { sql: state.fragment.sql, params },
            columns: state.columns,
            sortBy,
            sortInFragment: state.sortInFragment ?? false,
            connectionId: state.connectionId,
            connection: state.connection,
          };
        }
        if (dialect.name === "sqlserver") {
          return {
            fragment: { sql: state.fragment.sql, params },
            columns: state.columns,
            sortBy: operation.sortBy.slice(),
            sortInFragment: false,
            connectionId: state.connectionId,
            connection: state.connection,
          };
        }
        const orderBy = sortSpecsToSql(dialect, operation.sortBy);
        return {
          fragment: { sql: `SELECT * FROM ${from} ORDER BY ${orderBy}`, params },
          columns: state.columns,
          sortBy: null,
          sortInFragment: false,
          connectionId: state.connectionId,
          connection: state.connection,
        };
      }
      case "distinctRows": {
        if (!state.columns) return null;
        // Folding `Table.Distinct` safely requires a stable, explicit projection.
        // We only fold the full-row distinct case; distinct-by-columns semantics
        // require "first row wins" behavior that SQL cannot guarantee without
        // additional ordering constraints.
        if (operation.columns && operation.columns.length > 0) return null;
        const cols = state.columns.map((c) => `t.${quoteIdentifier(c)}`).join(", ");
        return {
          fragment: { sql: `SELECT DISTINCT ${cols} FROM ${from}`, params },
          columns: state.columns.slice(),
          sortBy,
          sortInFragment: false,
          connectionId: state.connectionId,
          connection: state.connection,
        };
      }
      case "groupBy": {
        if (operation.groupColumns.length === 0 && operation.aggregations.length === 0) return null;
        if (hasDuplicateStrings(operation.groupColumns)) return null;
        if (
          hasDuplicateStrings([
            ...operation.groupColumns,
            ...operation.aggregations.map((agg) => agg.as ?? `${agg.op} of ${agg.column}`),
          ])
        ) {
          return null;
        }
        const groupCols = operation.groupColumns.map((c) => `t.${quoteIdentifier(c)}`).join(", ");
        const aggSql = operation.aggregations.map((agg) => aggregationToSql(agg, dialect)).join(", ");
        const selectList = [groupCols, aggSql].filter(Boolean).join(", ");

        // We intentionally introduce a constant grouping key when groupColumns is
        // empty. This matches the local engine behavior: empty input produces no
        // rows (instead of a single row of aggregates over the empty set).
        const needsSyntheticGroup = operation.groupColumns.length === 0;
        const synthetic = "__group";
        const groupFrom = needsSyntheticGroup
          ? `(SELECT 1 AS ${quoteIdentifier(synthetic)}, s.* FROM (${state.fragment.sql}) AS s) AS t`
          : from;
        const groupByClause = needsSyntheticGroup ? ` GROUP BY t.${quoteIdentifier(synthetic)}` : groupCols ? ` GROUP BY ${groupCols}` : "";
        const columns = [
          ...operation.groupColumns,
          ...operation.aggregations.map((agg) => agg.as ?? `${agg.op} of ${agg.column}`),
        ];
        return {
          fragment: { sql: `SELECT ${selectList} FROM ${groupFrom}${groupByClause}`, params },
          columns,
          sortBy: null,
          sortInFragment: false,
          connectionId: state.connectionId,
          connection: state.connection,
        };
      }
      case "renameColumn": {
        if (!state.columns) return null;
        const idx = state.columns.indexOf(operation.oldName);
        if (idx === -1) return null;
        if (state.columns.includes(operation.newName) && operation.newName !== operation.oldName) {
          return null;
        }
        if (operation.oldName === operation.newName) {
          return {
            fragment: { sql: state.fragment.sql, params },
            columns: state.columns.slice(),
            sortBy,
            sortInFragment: state.sortInFragment ?? false,
            connectionId: state.connectionId,
            connection: state.connection,
          };
        }

        const cols = state.columns.map((name) => {
          if (name === operation.oldName) {
            return `t.${quoteIdentifier(operation.oldName)} AS ${quoteIdentifier(operation.newName)}`;
          }
          return `t.${quoteIdentifier(name)}`;
        });

        const nextColumns = state.columns.slice();
        nextColumns[idx] = operation.newName;

        const nextSortBy =
          dialect.name === "sqlserver" && sortBy
            ? sortBy.map((spec) => (spec.column === operation.oldName ? { ...spec, column: operation.newName } : spec))
            : sortBy;

        return {
          fragment: { sql: `SELECT ${cols.join(", ")} FROM ${from}`, params },
          columns: nextColumns,
          sortBy: nextSortBy,
          sortInFragment: false,
          connectionId: state.connectionId,
          connection: state.connection,
        };
      }
      case "changeType": {
        if (operation.newType === "any") {
          return {
            fragment: { sql: state.fragment.sql, params },
            columns: state.columns ? state.columns.slice() : null,
            sortBy,
            sortInFragment: state.sortInFragment ?? false,
            connectionId: state.connectionId,
            connection: state.connection,
          };
        }
        if (!state.columns) return null;
        if (!state.columns.includes(operation.column)) return null;
        if (dialect.name === "sqlserver" && sortBy && sortBy.some((spec) => spec.column === operation.column)) {
          return null;
        }

        const expr = changeTypeToSqlExpr(dialect, `t.${quoteIdentifier(operation.column)}`, operation.newType);
        if (!expr) return null;

        const cols = state.columns.map((name) => {
          if (name !== operation.column) return `t.${quoteIdentifier(name)}`;
          return `${expr} AS ${quoteIdentifier(name)}`;
        });

        return {
          fragment: { sql: `SELECT ${cols.join(", ")} FROM ${from}`, params },
          columns: state.columns.slice(),
          sortBy,
          sortInFragment: false,
          connectionId: state.connectionId,
          connection: state.connection,
        };
      }
      case "transformColumns": {
        if (!state.columns) return null;
        const byName = new Map(operation.transforms.map((t) => [t.column, t]));
        if (
          dialect.name === "sqlserver" &&
          sortBy &&
          operation.transforms.some((t) => sortBy.some((spec) => spec.column === t.column))
        ) {
          return null;
        }
        const projections = [];
        for (const name of state.columns) {
          const t = byName.get(name);
          if (!t) {
            projections.push(`t.${quoteIdentifier(name)}`);
            continue;
          }

          // Only fold pure type-casts where the transformation is identity.
          const rawFormula = typeof t.formula === "string" ? t.formula : "";
          try {
            const parsed = parseFormula(rawFormula);
            if (parsed.type !== "value") return null;
          } catch {
            return null;
          }
          if (!t.newType || t.newType === "any") {
            projections.push(`t.${quoteIdentifier(name)}`);
            continue;
          }

          const expr = changeTypeToSqlExpr(dialect, `t.${quoteIdentifier(name)}`, t.newType);
          if (!expr) return null;
          projections.push(`${expr} AS ${quoteIdentifier(name)}`);
        }

        return {
          fragment: { sql: `SELECT ${projections.join(", ")} FROM ${from}`, params },
          columns: state.columns.slice(),
          sortBy,
          sortInFragment: false,
          connectionId: state.connectionId,
          connection: state.connection,
        };
      }
      case "addColumn": {
        if (state.columns && state.columns.includes(operation.name)) return null;

        let compiled;
        try {
          const expr = parseFormula(operation.formula);
          compiled = compileExprToSql(expr, {
            alias: "t",
            quoteIdentifier,
            dialect,
            knownColumns: state.columns,
          });
        } catch {
          return null;
        }

        // Important: the expression appears in the SELECT list before the wrapped
        // subquery, so its placeholders occur *before* the placeholders in
        // `state.fragment.sql`. Preserve placeholder ordering by prepending the
        // formula params.
        const nextParams = [...compiled.params, ...params];
        const nextColumns = state.columns ? [...state.columns, operation.name] : null;
        return {
          fragment: {
            sql: `SELECT t.*, ${compiled.sql} AS ${quoteIdentifier(operation.name)} FROM ${from}`,
            params: nextParams,
          },
          columns: nextColumns,
          sortBy,
          sortInFragment: false,
          connectionId: state.connectionId,
          connection: state.connection,
        };
      }
      case "take": {
        if (!Number.isFinite(operation.count) || operation.count < 0) return null;
        if (dialect.name === "sqlserver") {
          // `TOP (?)` appears before the wrapped subquery, so its placeholder must
          // come before any placeholders in `state.fragment.sql`.
          const nextParams = [operation.count, ...params];
          const orderBy = sortBy && sortBy.length > 0 ? sortSpecsToSql(dialect, sortBy) : null;
          return {
            fragment: { sql: orderBy ? `SELECT TOP (?) * FROM ${from} ORDER BY ${orderBy}` : `SELECT TOP (?) * FROM ${from}`, params: nextParams },
            columns: state.columns,
            sortBy,
            sortInFragment: Boolean(orderBy),
            connectionId: state.connectionId,
            connection: state.connection,
          };
        }
        params.push(operation.count);
        return {
          fragment: { sql: `SELECT * FROM ${from} LIMIT ?`, params },
          columns: state.columns,
          sortBy: null,
          sortInFragment: false,
          connectionId: state.connectionId,
          connection: state.connection,
        };
      }
      case "skip": {
        if (!Number.isFinite(operation.count) || operation.count < 0) return null;
        if (dialect.name === "sqlserver") {
          // SQL Server requires an ORDER BY for OFFSET. When no ordering has been
          // specified we use a constant ordering expression to preserve the same
          // nondeterministic semantics as `OFFSET` without `ORDER BY` in other
          // dialects.
          const orderBy = sortBy && sortBy.length > 0 ? sortSpecsToSql(dialect, sortBy) : "(SELECT NULL)";
          params.push(operation.count);
          return {
            fragment: { sql: `SELECT * FROM ${from} ORDER BY ${orderBy} OFFSET ? ROWS`, params },
            columns: state.columns,
            sortBy,
            sortInFragment: Boolean(sortBy && sortBy.length > 0),
            connectionId: state.connectionId,
            connection: state.connection,
          };
        }

        params.push(operation.count);
        const offsetSql =
          ctx.dialect.name === "postgres"
            ? "OFFSET ?"
            : ctx.dialect.name === "mysql"
              ? "LIMIT 18446744073709551615 OFFSET ?"
              : "LIMIT -1 OFFSET ?";
        return {
          fragment: { sql: `SELECT * FROM ${from} ${offsetSql}`, params },
          columns: state.columns,
          sortBy: null,
          sortInFragment: false,
          connectionId: state.connectionId,
          connection: state.connection,
        };
      }
      case "merge": {
        if (!state.columns) return null;
        const rightQuery = ctx.queries?.[operation.rightQuery];
        if (!rightQuery) return null;
        const rightState = this.compileQueryToSqlState(rightQuery, ctx, callStack);
        if (!rightState?.columns) return null;
        const leftSourceId = sqlSourceIdForState(state);
        const rightSourceId = sqlSourceIdForState(rightState);
        if (!connectionsMatch(state, rightState)) {
          recordFoldingPrivacyDiagnostic(ctx, "merge", [leftSourceId, rightSourceId]);
          return null;
        }

        if (ctx.privacyMode && ctx.privacyMode !== "ignore") {
          const levelsBySourceId = ctx.privacyLevelsBySourceId;
          const leftLevel = getPrivacyLevel(levelsBySourceId, leftSourceId);
          const rightLevel = getPrivacyLevel(levelsBySourceId, rightSourceId);
          if (leftLevel !== rightLevel) {
            recordFoldingPrivacyDiagnostic(ctx, "merge", [leftSourceId, rightSourceId]);
            return null;
          }
        }

        const joinMode = operation.joinMode ?? "flat";
        if (joinMode !== "flat") return null;

        if (!isSqlFoldableJoinComparer(effectiveJoinComparer(operation))) return null;

        const leftKeys =
          Array.isArray(operation.leftKeys) && operation.leftKeys.length > 0
            ? operation.leftKeys
            : typeof operation.leftKey === "string" && operation.leftKey
              ? [operation.leftKey]
              : [];
        const rightKeys =
          Array.isArray(operation.rightKeys) && operation.rightKeys.length > 0
            ? operation.rightKeys
            : typeof operation.rightKey === "string" && operation.rightKey
              ? [operation.rightKey]
              : [];

        if (leftKeys.length === 0 || rightKeys.length === 0) return null;
        if (leftKeys.length !== rightKeys.length) return null;

        for (const key of leftKeys) {
          if (!state.columns.includes(key)) return null;
        }
        for (const key of rightKeys) {
          if (!rightState.columns.includes(key)) return null;
        }

        const join = joinTypeToSql(dialect, operation.joinType);
        if (!join) return null;

        const leftCols = state.columns;
        // Match local `Table.Join` semantics: exclude right-side key columns from
        // the output projection even when key names differ.
        const excludeRightKeys = new Set(rightKeys);

        const rightColsToInclude = rightState.columns.filter((c) => !excludeRightKeys.has(c));

        const outNames = makeUniqueColumnNames([...leftCols, ...rightColsToInclude]);
        const leftOutNames = outNames.slice(0, leftCols.length);
        const rightOutNames = outNames.slice(leftCols.length);

        const selectList = [
          ...leftCols.map((name, idx) => `l.${quoteIdentifier(name)} AS ${quoteIdentifier(leftOutNames[idx])}`),
          ...rightColsToInclude.map((name, idx) => `r.${quoteIdentifier(name)} AS ${quoteIdentifier(rightOutNames[idx])}`),
        ].join(", ");

        const on = leftKeys
          .map((leftKey, idx) =>
            nullSafeEqualsSql(dialect, `l.${quoteIdentifier(leftKey)}`, `r.${quoteIdentifier(rightKeys[idx])}`),
          )
          .join(" AND ");
        const mergedSql = `SELECT ${selectList} FROM (${state.fragment.sql}) AS l ${join} (${rightState.fragment.sql}) AS r ON ${on}`;
        return {
          fragment: { sql: mergedSql, params: [...state.fragment.params, ...rightState.fragment.params] },
          columns: [...leftOutNames, ...rightOutNames],
          sortBy: null,
          sortInFragment: false,
          connectionId: state.connectionId,
          connection: state.connection,
        };
      }
      case "append": {
        if (!state.columns) return null;
        const queries = ctx.queries;
        if (!queries) return null;

        const baseColumns = state.columns;
        /** @type {SqlState[]} */
        const branches = [state];
        const baseSourceId = sqlSourceIdForState(state);

        for (const id of operation.queries) {
          const q = queries[id];
          if (!q) return null;
          const compiled = this.compileQueryToSqlState(q, ctx, callStack);
          if (!compiled?.columns) return null;
          const branchSourceId = sqlSourceIdForState(compiled);
          if (!connectionsMatch(state, compiled)) {
            recordFoldingPrivacyDiagnostic(ctx, "append", [baseSourceId, branchSourceId]);
            return null;
          }
          if (ctx.privacyMode && ctx.privacyMode !== "ignore") {
            const levelsBySourceId = ctx.privacyLevelsBySourceId;
            const leftLevel = getPrivacyLevel(levelsBySourceId, baseSourceId);
            const rightLevel = getPrivacyLevel(levelsBySourceId, branchSourceId);
            if (leftLevel !== rightLevel) {
              recordFoldingPrivacyDiagnostic(ctx, "append", [baseSourceId, branchSourceId]);
              return null;
            }
          }
          if (!columnsCompatible(baseColumns, compiled.columns)) return null;
          branches.push(compiled);
        }

        const selectCols = baseColumns.map((name) => `t.${quoteIdentifier(name)}`).join(", ");
        const unionSql = branches
          .map((branch) => `(SELECT ${selectCols} FROM (${branch.fragment.sql}) AS t)`)
          .join(" UNION ALL ");
        const unionParams = branches.flatMap((branch) => branch.fragment.params);
        return {
          fragment: { sql: unionSql, params: unionParams },
          columns: baseColumns.slice(),
          sortBy: null,
          sortInFragment: false,
          connectionId: state.connectionId,
          connection: state.connection,
        };
      }
      default:
        return null;
    }
  }

  /**
   * @private
   * @param {SqlState} state
   * @param {QueryOperation} operation
   * @param {{
   *   dialect: SqlDialect;
   *   queries?: Record<string, Query> | null;
   *   getConnectionIdentity?: ((connection: unknown) => unknown) | null;
   *   privacyMode?: "ignore" | "enforce" | "warn";
   *   privacyLevelsBySourceId?: Record<string, import("../privacy/levels.js").PrivacyLevel>;
   *   diagnostics?: FoldingFirewallDiagnostic[];
   * }} ctx
   * @param {Set<string>} callStack
   * @returns {string}
   */
  explainSqlStepFailure(state, operation, ctx, callStack) {
    switch (operation.type) {
      case "selectColumns": {
        if (operation.columns.length === 0) return "invalid_projection";
        if (hasDuplicateStrings(operation.columns)) return "invalid_projection";
        return "unsupported_op";
      }
      case "removeColumns": {
        if (!state.columns) return "unknown_projection";
        for (const name of operation.columns) {
          if (!state.columns.includes(name)) return "unknown_projection";
        }
        const remaining = state.columns.filter((name) => !operation.columns.includes(name));
        if (remaining.length === 0) return "invalid_projection";
        return "unsupported_op";
      }
      case "filterRows":
      case "sortRows":
        return "unsupported_op";
      case "groupBy": {
        if (operation.groupColumns.length === 0 && operation.aggregations.length === 0) return "unsupported_op";
        if (hasDuplicateStrings(operation.groupColumns)) return "invalid_projection";
        if (
          hasDuplicateStrings([
            ...operation.groupColumns,
            ...operation.aggregations.map((agg) => agg.as ?? `${agg.op} of ${agg.column}`),
          ])
        ) {
          return "invalid_projection";
        }
        return "unsupported_op";
      }
      case "renameColumn": {
        if (!state.columns) return "unknown_projection";
        const idx = state.columns.indexOf(operation.oldName);
        if (idx === -1) return "unknown_projection";
        if (state.columns.includes(operation.newName) && operation.newName !== operation.oldName) return "invalid_projection";
        return "unsupported_op";
      }
      case "changeType": {
        if (operation.newType === "any") return "unsupported_op";
        if (!state.columns) return "unknown_projection";
        if (!state.columns.includes(operation.column)) return "unknown_projection";
        const expr = changeTypeToSqlExpr(ctx.dialect, `t.${ctx.dialect.quoteIdentifier(operation.column)}`, operation.newType);
        if (!expr) return "unsupported_type";
        return "unsupported_op";
      }
      case "addColumn": {
        if (state.columns && state.columns.includes(operation.name)) return "invalid_projection";
        try {
          const expr = parseFormula(operation.formula);
          // Only used to validate that the expression is foldable; the actual SQL
          // compilation happens in `applySqlStep`.
          compileExprToSql(expr, {
            alias: "t",
            quoteIdentifier: ctx.dialect.quoteIdentifier,
            dialect: ctx.dialect,
            knownColumns: state.columns,
          });
        } catch {
          return "unsafe_formula";
        }
        return "unsupported_op";
      }
      case "take": {
        if (!Number.isFinite(operation.count) || operation.count < 0) return "invalid_argument";
        return "unsupported_op";
      }
      case "skip": {
        if (!Number.isFinite(operation.count) || operation.count < 0) return "invalid_argument";
        return "unsupported_op";
      }
      case "merge": {
        if (!state.columns) return "unknown_projection";
        const rightQuery = ctx.queries?.[operation.rightQuery];
        if (!rightQuery) return "missing_query";
        const rightState = this.compileQueryToSqlState(rightQuery, ctx, callStack);
        if (!rightState?.columns) return "unsupported_query";
        const leftSourceId = sqlSourceIdForState(state);
        const rightSourceId = sqlSourceIdForState(rightState);
        const privacyMode = ctx.privacyMode;
        const levelsBySourceId = ctx.privacyLevelsBySourceId;
        const leftLevel = privacyMode && privacyMode !== "ignore" ? getPrivacyLevel(levelsBySourceId, leftSourceId) : null;
        const rightLevel = privacyMode && privacyMode !== "ignore" ? getPrivacyLevel(levelsBySourceId, rightSourceId) : null;
        if (!connectionsMatch(state, rightState)) {
          if (leftLevel && rightLevel && leftLevel !== rightLevel) return "privacy_firewall";
          return "different_connection";
        }
        if (ctx.privacyMode && ctx.privacyMode !== "ignore") {
          if (leftLevel && rightLevel && leftLevel !== rightLevel) return "privacy_firewall";
        }

        const joinMode = operation.joinMode ?? "flat";
        if (joinMode !== "flat") return "unsupported_join_mode";

        if (!isSqlFoldableJoinComparer(effectiveJoinComparer(operation))) return "unsupported_comparer";

        const leftKeys =
          Array.isArray(operation.leftKeys) && operation.leftKeys.length > 0
            ? operation.leftKeys
            : typeof operation.leftKey === "string" && operation.leftKey
              ? [operation.leftKey]
              : [];
        const rightKeys =
          Array.isArray(operation.rightKeys) && operation.rightKeys.length > 0
            ? operation.rightKeys
            : typeof operation.rightKey === "string" && operation.rightKey
              ? [operation.rightKey]
              : [];

        if (leftKeys.length === 0 || rightKeys.length === 0) return "invalid_argument";
        if (leftKeys.length !== rightKeys.length) return "invalid_argument";
        for (const key of leftKeys) {
          if (!state.columns.includes(key)) return "unknown_projection";
        }
        for (const key of rightKeys) {
          if (!rightState.columns.includes(key)) return "unknown_projection";
        }

        const join = joinTypeToSql(ctx.dialect, operation.joinType);
        if (!join) return "unsupported_join_type";
        return "unsupported_op";
      }
      case "append": {
        if (!state.columns) return "unknown_projection";
        if (!ctx.queries) return "missing_queries";
        for (const id of operation.queries) {
          const q = ctx.queries[id];
          if (!q) return "missing_query";
          const compiled = this.compileQueryToSqlState(q, ctx, callStack);
          if (!compiled?.columns) return "unsupported_query";
          if (!connectionsMatch(state, compiled)) {
            if (ctx.privacyMode && ctx.privacyMode !== "ignore") {
              const baseSourceId = sqlSourceIdForState(state);
              const branchSourceId = sqlSourceIdForState(compiled);
              const levelsBySourceId = ctx.privacyLevelsBySourceId;
              const leftLevel = getPrivacyLevel(levelsBySourceId, baseSourceId);
              const rightLevel = getPrivacyLevel(levelsBySourceId, branchSourceId);
              if (leftLevel !== rightLevel) return "privacy_firewall";
            }
            return "different_connection";
          }
          if (ctx.privacyMode && ctx.privacyMode !== "ignore") {
            const baseSourceId = sqlSourceIdForState(state);
            const branchSourceId = sqlSourceIdForState(compiled);
            const levelsBySourceId = ctx.privacyLevelsBySourceId;
            const leftLevel = getPrivacyLevel(levelsBySourceId, baseSourceId);
            const rightLevel = getPrivacyLevel(levelsBySourceId, branchSourceId);
            if (leftLevel !== rightLevel) return "privacy_firewall";
          }
          if (!columnsCompatible(state.columns, compiled.columns)) return "incompatible_schema";
        }
        return "unsupported_op";
      }
      default:
        return "unsupported_op";
    }
  }
}

/**
 * @param {import("../model.js").QuerySource} source
 * @param {{ queries: Record<string, Query> | null | undefined, callStack: Set<string>, dialect?: SqlDialect }} ctx
 * @returns {string}
 */
function explainSourceFailure(source, ctx) {
  switch (source.type) {
    case "query": {
      if (!ctx.queries) return "missing_queries";
      const target = ctx.queries[source.queryId];
      if (!target) return "missing_query";
      if (ctx.callStack.has(target.id)) return "query_cycle";
      return "unsupported_query";
    }
    case "database":
      if (ctx.dialect?.name === "sqlserver" && !isSqlServerDerivedTableSafe(source.query)) {
        return "sqlserver_order_by_in_source";
      }
      return "unsupported_source";
    default:
      return "unsupported_source";
  }
}

/**
 * @param {{ privacyMode?: "ignore" | "enforce" | "warn", privacyLevelsBySourceId?: Record<string, import("../privacy/levels.js").PrivacyLevel>, diagnostics?: FoldingFirewallDiagnostic[] }} ctx
 * @param {"merge" | "append"} operation
 * @param {string[]} sourceIds
 */
function recordFoldingPrivacyDiagnostic(ctx, operation, sourceIds) {
  const diagnostics = ctx.diagnostics;
  if (!diagnostics) return;
  if (!ctx.privacyMode || ctx.privacyMode === "ignore") return;

  const infos = collectSourcePrivacy(sourceIds, ctx.privacyLevelsBySourceId);
  const levels = distinctPrivacyLevels(infos);
  if (levels.size <= 1) return;

  diagnostics.push({
    kind: "privacy:firewall",
    phase: "folding",
    operation,
    sources: infos,
    message: `Formula firewall prevented folding of ${operation} across privacy levels (${Array.from(levels).join(", ")})`,
  });
}

/**
 * Compute a stable privacy `sourceId` for a SQL folding branch.
 *
 * @param {SqlState} state
 * @returns {string}
 */
function sqlSourceIdForState(state) {
  if (state.connectionId) return getSqlSourceId(state.connectionId);
  return getSqlSourceId(state.connection);
}

/**
 * @param {import("../model.js").DatabaseQuerySource} source
 * @param {((connection: unknown) => unknown) | null | undefined} getConnectionIdentity
 * @returns {string | null}
 */
function resolveConnectionId(source, getConnectionIdentity) {
  if (typeof source.connectionId === "string" && source.connectionId) {
    return source.connectionId;
  }

  const connection = source.connection;

  // Prefer a host-provided identity hook when available.
  if (getConnectionIdentity) {
    try {
      const identity = getConnectionIdentity(connection);
      if (identity != null) {
        if (typeof identity === "string") return identity;
        return hashValue(identity);
      }
    } catch {
      // Fall through to conservative fallback below.
    }
  }

  // Conservative fallback: treat primitives as identities, and treat `{ id: string }` as a stable identity.
  if (connection == null) return null;
  const type = typeof connection;
  if (type === "string" && connection) return connection;
  if (type === "number" || type === "boolean") return hashValue(connection);

  if (type === "object" && !Array.isArray(connection)) {
    // @ts-ignore - runtime inspection
    if (typeof connection.id === "string" && connection.id) return connection.id;
  }

  return null;
}

/**
 * @param {SqlState} left
 * @param {SqlState} right
 * @returns {boolean}
 */
function connectionsMatch(left, right) {
  if (left.connectionId && right.connectionId) {
    return left.connectionId === right.connectionId;
  }
  if (!left.connectionId && !right.connectionId) {
    return left.connection === right.connection;
  }
  // If only one side has an identity, be conservative and only allow folding
  // when both sides share the same (referentially equal) connection handle.
  return left.connection === right.connection;
}

/**
 * Prefer per-key comparers when present; otherwise fall back to the scalar comparer.
 *
 * @param {{ comparer?: unknown; comparers?: unknown }} op
 * @returns {unknown}
 */
function effectiveJoinComparer(op) {
  const list = /** @type {any} */ (op).comparers;
  if (Array.isArray(list) && list.length > 0) return list;
  return /** @type {any} */ (op).comparer;
}

/**
 * `Table.Join` / `Table.NestedJoin` support passing a comparer to override how
 * join keys are compared. SQL folding is conservative and only supports the
 * default comparer semantics.
 *
 * @param {unknown} comparer
 * @returns {boolean}
 */
function isSqlFoldableJoinComparer(comparer) {
  if (comparer == null) return true;
  if (Array.isArray(comparer)) {
    return comparer.every((entry) => isSqlFoldableJoinComparer(entry));
  }
  if (!comparer || typeof comparer !== "object" || Array.isArray(comparer)) return false;
  // @ts-ignore - runtime inspection
  const name = typeof comparer.comparer === "string" ? comparer.comparer.toLowerCase() : "";
  // @ts-ignore - runtime inspection
  const caseSensitive = comparer.caseSensitive;
  if (name !== "ordinal") return false;
  if (caseSensitive === false) return false;
  return true;
}

/**
 * @param {SqlDialect} dialect
 * @param {import("../model.js").SortSpec[]} specs
 * @returns {string}
 */
function sortSpecsToSql(dialect, specs) {
  return specs.flatMap((spec) => dialect.sortSpecToSql("t", spec)).join(", ");
}

/**
 * SQL Server does not allow `ORDER BY` in derived tables unless paired with
 * `TOP`/`OFFSET`. Because the folding engine wraps each step in a subquery, we
 * defer emitting the final `ORDER BY` until the outermost query.
 *
 * @param {SqlState} state
 * @param {SqlDialect} dialect
 * @returns {string}
 */
function finalizeSqlForDialect(state, dialect) {
  if (dialect.name !== "sqlserver") return state.fragment.sql;
  const sortBy = state.sortBy;
  if (!sortBy || sortBy.length === 0) return state.fragment.sql;
  if (state.sortInFragment) return state.fragment.sql;
  return `SELECT * FROM (${state.fragment.sql}) AS t ORDER BY ${sortSpecsToSql(dialect, sortBy)}`;
}

/**
 * SQL Server forbids `ORDER BY` in derived tables unless paired with `TOP` or
 * `OFFSET`. Because the folding engine wraps each step in a derived table, a
 * source query containing a top-level `ORDER BY` (without `TOP`/`OFFSET`) would
 * produce invalid SQL once folding begins.
 *
 * @param {string} sql
 * @returns {boolean}
 */
function isSqlServerDerivedTableSafe(sql) {
  let inSingle = false;
  let inDouble = false;
  let inBracket = false;
  let inLineComment = false;
  let inBlockComment = false;
  let parenDepth = 0;

  /** @type {string[]} */
  const prefixTokens = [];
  /** @type {string | null} */
  let prevToken = null;
  let hasOrderBy = false;
  let hasOffset = false;

  for (let i = 0; i < sql.length; i++) {
    const ch = sql[i];
    const next = sql[i + 1] ?? "";

    if (inLineComment) {
      if (ch === "\n") inLineComment = false;
      continue;
    }

    if (inBlockComment) {
      if (ch === "*" && next === "/") {
        i += 1;
        inBlockComment = false;
      }
      continue;
    }

    if (inSingle) {
      if (ch === "'") {
        if (next === "'") i += 1;
        else inSingle = false;
      }
      continue;
    }

    if (inDouble) {
      if (ch === '"') {
        if (next === '"') i += 1;
        else inDouble = false;
      }
      continue;
    }

    if (inBracket) {
      if (ch === "]") {
        if (next === "]") i += 1;
        else inBracket = false;
      }
      continue;
    }

    if (ch === "-" && next === "-") {
      i += 1;
      inLineComment = true;
      continue;
    }

    if (ch === "/" && next === "*") {
      i += 1;
      inBlockComment = true;
      continue;
    }

    if (ch === "'") {
      inSingle = true;
      continue;
    }

    if (ch === '"') {
      inDouble = true;
      continue;
    }

    if (ch === "[") {
      inBracket = true;
      continue;
    }

    if (ch === "(") {
      parenDepth += 1;
      continue;
    }

    if (ch === ")") {
      if (parenDepth > 0) parenDepth -= 1;
      continue;
    }

    if (parenDepth !== 0) continue;

    if (/[A-Za-z_]/.test(ch)) {
      let end = i + 1;
      while (end < sql.length && /[A-Za-z0-9_]/.test(sql[end])) end += 1;
      const token = sql.slice(i, end).toUpperCase();

      if (prefixTokens.length < 3) prefixTokens.push(token);

      if (prevToken === "ORDER" && token === "BY") hasOrderBy = true;
      if (token === "OFFSET") hasOffset = true;

      prevToken = token;
      i = end - 1;
    }
  }

  if (!hasOrderBy) return true;
  if (hasOffset) return true;

  const startsWithSelectTop =
    prefixTokens[0] === "SELECT" &&
    (prefixTokens[1] === "TOP" || ((prefixTokens[1] === "DISTINCT" || prefixTokens[1] === "ALL") && prefixTokens[2] === "TOP"));
  return startsWithSelectTop;
}

/**
 * @param {Aggregation} agg
 * @param {SqlDialect} dialect
 * @returns {string}
 */
function aggregationToSql(agg, dialect) {
  const quoteIdentifier = dialect.quoteIdentifier;
  const alias = quoteIdentifier(agg.as ?? `${agg.op} of ${agg.column}`);
  const col = `t.${quoteIdentifier(agg.column)}`;
  const numeric = safeCastNumberToSql(dialect, col) ?? col;
  switch (agg.op) {
    case "sum":
      return `COALESCE(SUM(${numeric}), 0) AS ${alias}`;
    case "count":
      return `COUNT(*) AS ${alias}`;
    case "average":
      return `AVG(${numeric}) AS ${alias}`;
    case "min":
      return `MIN(${col}) AS ${alias}`;
    case "max":
      return `MAX(${col}) AS ${alias}`;
    case "countDistinct":
      return `(COUNT(DISTINCT ${col}) + COALESCE(MAX(CASE WHEN ${col} IS NULL THEN 1 ELSE 0 END), 0)) AS ${alias}`;
    default: {
      /** @type {never} */
      const exhausted = agg.op;
      throw new Error(`Unsupported aggregation '${exhausted}'`);
    }
  }
}

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
function isSqlScalarValue(value) {
  if (value == null) return true;
  if (typeof value === "string") return true;
  if (typeof value === "boolean") return true;
  if (typeof value === "number") return Number.isFinite(value);
  if (isDate(value)) return true;
  if (value instanceof PqDateTimeZone) return true;
  if (value instanceof PqTime) return true;
  if (value instanceof PqDuration) return true;
  if (value instanceof PqDecimal) return true;
  return false;
}

/**
 * @param {import("../model.js").FilterPredicate} predicate
 * @returns {boolean}
 */
function predicateHasOnlySqlScalarValues(predicate) {
  if (!predicate || typeof predicate !== "object") return false;
  switch (predicate.type) {
    case "and":
    case "or":
      return predicate.predicates.every(predicateHasOnlySqlScalarValues);
    case "not":
      return predicateHasOnlySqlScalarValues(predicate.predicate);
    case "comparison": {
      // LIKE-based predicates stringify values before passing them through the
      // SQL parameterizer, so they never require non-scalar params.
      if (predicate.operator === "contains" || predicate.operator === "startsWith" || predicate.operator === "endsWith") {
        return true;
      }
      return isSqlScalarValue(predicate.value);
    }
    default: {
      /** @type {never} */
      const exhausted = predicate;
      throw new Error(`Unsupported predicate type '${exhausted.type}'`);
    }
  }
}

/**
 * @param {SqlDialect} dialect
 * @param {unknown[]} params
 * @returns {(value: unknown) => string}
 */
function makeParam(dialect, params) {
  return (value) => {
    if (isDate(value)) {
      params.push(dialect.formatDateParam(value));
      return "?";
    }
    if (value instanceof PqDateTimeZone) {
      params.push(dialect.formatDateParam(value.toDate()));
      return "?";
    }
    if (value instanceof PqTime) {
      params.push(value.toString());
      return "?";
    }
    if (value instanceof PqDuration) {
      params.push(value.toString());
      return "?";
    }
    if (value instanceof PqDecimal) {
      params.push(value.value);
      return "?";
    }
    if (typeof value === "number" && !Number.isFinite(value)) {
      params.push(null);
      return "?";
    }
    params.push(value === undefined ? null : value);
    return "?";
  };
}

/**
 * @param {SqlDialect | SqlDialectName | undefined} dialect
 * @returns {SqlDialect}
 */
function resolveDialect(dialect) {
  if (!dialect) return POSTGRES_DIALECT;
  if (typeof dialect === "string") return getSqlDialect(dialect);
  return dialect;
}

/**
 * @param {SqlDialect} dialect
 * @param {"inner" | "left" | "right" | "full" | "leftAnti" | "rightAnti" | "leftSemi" | "rightSemi"} joinType
 * @returns {string | null}
 */
function joinTypeToSql(dialect, joinType) {
  switch (joinType) {
    case "inner":
      return "INNER JOIN";
    case "left":
      return "LEFT JOIN";
    case "right":
      return dialect.name === "postgres" || dialect.name === "sqlserver" || dialect.name === "mysql" ? "RIGHT JOIN" : null;
    case "full":
      return dialect.name === "postgres" || dialect.name === "sqlserver" ? "FULL OUTER JOIN" : null;
    case "leftAnti":
    case "rightAnti":
    case "leftSemi":
    case "rightSemi":
      return null;
    default: {
      /** @type {never} */
      const exhausted = joinType;
      throw new Error(`Unsupported joinType '${exhausted}'`);
    }
  }
}

/**
 * Join key comparison that matches the local engine semantics: `null` values
 * compare equal when joining/merging.
 *
 * @param {SqlDialect} dialect
 * @param {string} leftExpr
 * @param {string} rightExpr
 * @returns {string}
 */
function nullSafeEqualsSql(dialect, leftExpr, rightExpr) {
  switch (dialect.name) {
    case "postgres":
      return `${leftExpr} IS NOT DISTINCT FROM ${rightExpr}`;
    case "mysql":
      return `${leftExpr} <=> ${rightExpr}`;
    case "sqlite":
      return `${leftExpr} IS ${rightExpr}`;
    case "sqlserver":
      return `(${leftExpr} = ${rightExpr} OR (${leftExpr} IS NULL AND ${rightExpr} IS NULL))`;
    default: {
      /** @type {never} */
      const exhausted = dialect.name;
      throw new Error(`Unsupported dialect '${exhausted}'`);
    }
  }
}

/**
 * Compute output column names for `Table.ExpandTableColumn` while preserving
 * Power Query-style uniquing semantics: existing columns are never renamed,
 * and expanded columns use the `A`, `A.1`, `A.2` pattern.
 *
 * This mirrors the local execution logic in `src/steps.js`.
 *
 * @param {string[]} reservedNames
 * @param {string[]} rawExpandedNames
 * @returns {string[]}
 */
function computeExpandedColumnNames(reservedNames, rawExpandedNames) {
  const reserved = new Set(reservedNames);
  const baseExpanded = makeUniqueColumnNames(rawExpandedNames);
  return baseExpanded.map((base) => {
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
}

/**
 * @param {string[]} base
 * @param {string[]} candidate
 * @returns {boolean}
 */
function columnsCompatible(base, candidate) {
  if (base.length !== candidate.length) return false;
  const baseSet = new Set(base);
  if (baseSet.size !== base.length) return false;
  const candidateSet = new Set(candidate);
  if (candidateSet.size !== candidate.length) return false;
  if (baseSet.size !== candidateSet.size) return false;
  for (const name of baseSet) {
    if (!candidateSet.has(name)) return false;
  }
  return true;
}

/**
 * @param {string[]} values
 * @returns {boolean}
 */
function hasDuplicateStrings(values) {
  const seen = new Set();
  for (const value of values) {
    if (seen.has(value)) return true;
    seen.add(value);
  }
  return false;
}

/**
 * @param {SqlDialect} dialect
 * @param {DataType} type
 * @returns {string | null}
 */
function dataTypeToSqlType(dialect, type) {
  if (type === "any") return null;
  switch (dialect.name) {
    case "postgres":
      switch (type) {
        case "string":
          return "TEXT";
        case "number":
          return "DOUBLE PRECISION";
        case "boolean":
          return "BOOLEAN";
        case "date":
          return "TIMESTAMPTZ";
        case "datetime":
          return "TIMESTAMP";
        case "datetimezone":
          return "TIMESTAMPTZ";
        case "time":
          return "TIME";
        case "duration":
          return "INTERVAL";
        case "decimal":
          return "NUMERIC";
        case "binary":
          return "BYTEA";
        default: {
          /** @type {never} */
          const exhausted = type;
          throw new Error(`Unsupported type '${exhausted}'`);
        }
      }
    case "mysql":
      switch (type) {
        case "string":
          return "TEXT";
        case "number":
          return "DOUBLE";
        case "boolean":
          return "BOOLEAN";
        case "date":
          return "DATETIME";
        case "datetime":
          return "DATETIME";
        case "datetimezone":
          return "DATETIME";
        case "time":
          return "TIME";
        case "duration":
          return null;
        case "decimal":
          return "DECIMAL";
        case "binary":
          return "BLOB";
        default: {
          /** @type {never} */
          const exhausted = type;
          throw new Error(`Unsupported type '${exhausted}'`);
        }
      }
    case "sqlite":
      switch (type) {
        case "string":
          return "TEXT";
        case "number":
          return "REAL";
        case "boolean":
          return "INTEGER";
        case "date":
          return "TEXT";
        case "datetime":
        case "datetimezone":
        case "time":
        case "duration":
        case "decimal":
          return "TEXT";
        case "binary":
          return "BLOB";
        default: {
          /** @type {never} */
          const exhausted = type;
          throw new Error(`Unsupported type '${exhausted}'`);
        }
      }
    case "sqlserver":
      switch (type) {
        case "string":
          return "NVARCHAR(MAX)";
        case "number":
          return "FLOAT";
        case "boolean":
          return "BIT";
        case "date":
          return "DATETIME2";
        case "datetime":
          return "DATETIME2";
        case "datetimezone":
          return "DATETIMEOFFSET";
        case "time":
          return "TIME";
        case "duration":
          return null;
        case "decimal":
          return "DECIMAL";
        case "binary":
          return "VARBINARY(MAX)";
        default: {
          /** @type {never} */
          const exhausted = type;
          throw new Error(`Unsupported type '${exhausted}'`);
        }
      }
    default: {
      /** @type {never} */
      const exhausted = dialect.name;
      throw new Error(`Unsupported dialect '${exhausted}'`);
    }
  }
}

/**
 * @param {SqlDialect} dialect
 * @param {string} colRef
 * @param {DataType} newType
 * @returns {string | null}
 */
function changeTypeToSqlExpr(dialect, colRef, newType) {
  switch (newType) {
    case "string":
      return dialect.castText(colRef);
    case "number":
      return safeCastNumberToSql(dialect, colRef);
    case "datetime":
    case "datetimezone":
      return safeCastDateTimeToSql(dialect, colRef, newType);
    default:
      return null;
  }
}

/**
 * @param {SqlDialect} dialect
 * @param {string} colRef
 * @param {"datetime" | "datetimezone"} newType
 * @returns {string | null}
 */
function safeCastDateTimeToSql(dialect, colRef, newType) {
  const sqlType = dataTypeToSqlType(dialect, newType);
  if (!sqlType) return null;

  // Local `changeType` semantics map unparseable inputs to `null` instead of
  // throwing. We approximate that behavior by regex-gating the cast. Dialects
  // without a regex operator (e.g. SQLite) fall back to local execution.
  const trimmed =
    dialect.name === "sqlserver" ? `LTRIM(RTRIM(${dialect.castText(colRef)}))` : `TRIM(${dialect.castText(colRef)})`;
  const pattern =
    "'^\\\\d{4}-\\\\d{2}-\\\\d{2}([ T]\\\\d{2}:\\\\d{2}(:\\\\d{2}(\\\\.\\\\d{1,9})?)?([zZ]|[+-]\\\\d{2}(:?\\\\d{2})?)?)?$'";
  const casted = `CAST(${trimmed} AS ${sqlType})`;

  switch (dialect.name) {
    case "postgres":
      return `CASE WHEN ${trimmed} = '' THEN NULL WHEN ${trimmed} ~ ${pattern} THEN ${casted} ELSE NULL END`;
    case "mysql":
      return `CASE WHEN ${trimmed} = '' THEN NULL WHEN ${trimmed} REGEXP ${pattern} THEN ${casted} ELSE NULL END`;
    case "sqlite":
      return null;
    case "sqlserver":
      return `TRY_CAST(NULLIF(${trimmed}, '') AS ${sqlType})`;
    default: {
      /** @type {never} */
      const exhausted = dialect.name;
      throw new Error(`Unsupported dialect '${exhausted}'`);
    }
  }
}

/**
 * @param {SqlDialect} dialect
 * @param {string} colRef
 * @returns {string | null}
 */
function safeCastNumberToSql(dialect, colRef) {
  const sqlType = dataTypeToSqlType(dialect, "number");
  if (!sqlType) return null;

  // Local `changeType` semantics map non-numeric inputs to `null` instead of
  // throwing (and they accept leading/trailing whitespace). We emulate that
  // behavior with a conservative numeric regex gate. Dialects without a regex
  // operator (e.g. SQLite) fall back to local execution.
  const trimmed =
    dialect.name === "sqlserver" ? `LTRIM(RTRIM(${dialect.castText(colRef)}))` : `TRIM(${dialect.castText(colRef)})`;
  const pattern = "'^[+-]?([0-9]+([.][0-9]*)?|[.][0-9]+)([eE][+-]?[0-9]+)?$'";
  const casted = `CAST(${trimmed} AS ${sqlType})`;

  switch (dialect.name) {
    case "postgres":
      return `CASE WHEN ${trimmed} = '' THEN NULL WHEN ${trimmed} ~ ${pattern} THEN (CASE WHEN isfinite(${casted}) THEN ${casted} ELSE NULL END) ELSE NULL END`;
    case "mysql":
      return `CASE WHEN ${trimmed} = '' THEN NULL WHEN ${trimmed} REGEXP ${pattern} THEN (CASE WHEN ABS(${casted}) <= 1.7976931348623157e308 THEN ${casted} ELSE NULL END) ELSE NULL END`;
    case "sqlite":
      return null;
    case "sqlserver":
      return `TRY_CAST(NULLIF(${trimmed}, '') AS ${sqlType})`;
    default: {
      /** @type {never} */
      const exhausted = dialect.name;
      throw new Error(`Unsupported dialect '${exhausted}'`);
    }
  }
}
