import { predicateToSql } from "../predicate.js";
import { POSTGRES_DIALECT, getSqlDialect } from "./dialect.js";

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
 * }} LocalPlan
 *
 * @typedef {{
 *   type: "sql";
 *   sql: string;
 *   params: unknown[];
 * }} SqlPlan
 *
 * @typedef {{
 *   type: "hybrid";
 *   sql: string;
 *   params: unknown[];
 *   localSteps: QueryStep[];
 * }} HybridPlan
 *
 * @typedef {LocalPlan | SqlPlan | HybridPlan} CompiledQueryPlan
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
 *   // Connection identity used to ensure we only fold `merge`/`append` when both
 *   // sides originate from the same database connection.
 *   connection: unknown;
 * }} SqlState
 */

/**
 * @typedef {{
 *   dialect?: SqlDialect | SqlDialectName;
 *   // Queries are required to fold operations like `merge`, `append`, and
 *   // sources of type `query`.
 *   queries?: Record<string, Query>;
 * }} CompileOptions
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
    /** @type {Record<string, Query> | null} */
    this.queries = options.queries ?? null;

    /** @type {Set<string>} */
    this.foldable = new Set([
      "selectColumns",
      "removeColumns",
      "filterRows",
      "sortRows",
      "groupBy",
      "renameColumn",
      "changeType",
      "addColumn",
      "merge",
      "append",
      "take",
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
    const initial = this.compileSourceToSqlState(query.source, { dialect, queries }, new Set([query.id]));
    if (!initial) {
      return { type: "local", steps: query.steps };
    }

    /** @type {SqlState} */
    let current = initial;
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

      const next = this.applySqlStep(current, step.operation, { dialect, queries }, new Set([query.id]));
      if (!next) {
        foldingBroken = true;
        localSteps.push(step);
        continue;
      }
      current = next;
    }

    if (localSteps.length === 0) {
      return { type: "sql", sql: current.fragment.sql, params: current.fragment.params };
    }
    return { type: "hybrid", sql: current.fragment.sql, params: current.fragment.params, localSteps };
  }

  /**
   * Try to compile an entire query to SQL. Used when folding operations depend
   * on other queries (e.g. `merge` + `append`).
   *
   * @param {Query} query
   * @param {{ dialect: SqlDialect, queries?: Record<string, Query> | null }} ctx
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
    for (const step of query.steps) {
      if (!this.foldable.has(step.operation.type)) return null;
      const next = this.applySqlStep(current, step.operation, ctx, nextStack);
      if (!next) return null;
      current = next;
    }
    return current;
  }

  /**
   * @param {import("../model.js").QuerySource} source
   * @param {{ dialect: SqlDialect, queries?: Record<string, Query> | null }} ctx
   * @param {Set<string>} callStack
   * @returns {SqlState | null}
   */
  compileSourceToSqlState(source, ctx, callStack) {
    switch (source.type) {
      case "database": {
        const columns = source.columns ? source.columns.slice() : null;
        return { fragment: { sql: source.query, params: [] }, columns, connection: source.connection };
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
   * @param {{ dialect: SqlDialect, queries?: Record<string, Query> | null }} ctx
   * @param {Set<string>} callStack
   * @returns {SqlState | null}
   */
  applySqlStep(state, operation, ctx, callStack) {
    const dialect = ctx.dialect;
    const quoteIdentifier = dialect.quoteIdentifier;
    const params = state.fragment.params.slice();
    const param = makeParam(dialect, params);

    const from = `(${state.fragment.sql}) AS t`;
    switch (operation.type) {
      case "selectColumns": {
        const cols = operation.columns.map((c) => `t.${quoteIdentifier(c)}`).join(", ");
        return {
          fragment: { sql: `SELECT ${cols} FROM ${from}`, params },
          columns: operation.columns.slice(),
          connection: state.connection,
        };
      }
      case "removeColumns": {
        if (!state.columns) return null;
        const remove = new Set(operation.columns);
        for (const name of operation.columns) {
          if (!state.columns.includes(name)) return null;
        }

        const remaining = state.columns.filter((name) => !remove.has(name));
        if (remaining.length === 0) return null;
        const cols = remaining.map((c) => `t.${quoteIdentifier(c)}`).join(", ");
        return {
          fragment: { sql: `SELECT ${cols} FROM ${from}`, params },
          columns: remaining,
          connection: state.connection,
        };
      }
      case "filterRows": {
        const where = predicateToSql(operation.predicate, {
          alias: "t",
          quoteIdentifier,
          castText: dialect.castText,
          param,
        });
        return {
          fragment: { sql: `SELECT * FROM ${from} WHERE ${where}`, params },
          columns: state.columns,
          connection: state.connection,
        };
      }
      case "sortRows": {
        if (operation.sortBy.length === 0) {
          return { fragment: { sql: state.fragment.sql, params }, columns: state.columns, connection: state.connection };
        }
        const orderBy = sortSpecsToSql(dialect, operation.sortBy);
        return {
          fragment: { sql: `SELECT * FROM ${from} ORDER BY ${orderBy}`, params },
          columns: state.columns,
          connection: state.connection,
        };
      }
      case "groupBy": {
        const groupCols = operation.groupColumns.map((c) => `t.${quoteIdentifier(c)}`).join(", ");
        const aggSql = operation.aggregations.map((agg) => aggregationToSql(agg, quoteIdentifier)).join(", ");
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

        return {
          fragment: { sql: `SELECT ${cols.join(", ")} FROM ${from}`, params },
          columns: nextColumns,
          connection: state.connection,
        };
      }
      case "changeType": {
        if (operation.newType === "any") {
          return {
            fragment: { sql: state.fragment.sql, params },
            columns: state.columns ? state.columns.slice() : null,
            connection: state.connection,
          };
        }
        if (!state.columns) return null;
        if (!state.columns.includes(operation.column)) return null;

        const expr = changeTypeToSqlExpr(dialect, `t.${quoteIdentifier(operation.column)}`, operation.newType);
        if (!expr) return null;

        const cols = state.columns.map((name) => {
          if (name !== operation.column) return `t.${quoteIdentifier(name)}`;
          return `${expr} AS ${quoteIdentifier(name)}`;
        });

        return {
          fragment: { sql: `SELECT ${cols.join(", ")} FROM ${from}`, params },
          columns: state.columns.slice(),
          connection: state.connection,
        };
      }
      case "addColumn": {
        if (state.columns && state.columns.includes(operation.name)) return null;

        const exprSql = compileFormulaToSql(operation.formula, {
          alias: "t",
          quoteIdentifier,
          params,
          dialect,
          knownColumns: state.columns,
        });
        if (!exprSql) return null;

        const nextColumns = state.columns ? [...state.columns, operation.name] : null;
        return {
          fragment: { sql: `SELECT t.*, ${exprSql} AS ${quoteIdentifier(operation.name)} FROM ${from}`, params },
          columns: nextColumns,
          connection: state.connection,
        };
      }
      case "take": {
        if (!Number.isFinite(operation.count) || operation.count < 0) return null;
        params.push(operation.count);
        return {
          fragment: { sql: `SELECT * FROM ${from} LIMIT ?`, params },
          columns: state.columns,
          connection: state.connection,
        };
      }
      case "merge": {
        if (!state.columns) return null;
        const rightQuery = ctx.queries?.[operation.rightQuery];
        if (!rightQuery) return null;
        const rightState = this.compileQueryToSqlState(rightQuery, ctx, callStack);
        if (!rightState?.columns) return null;
        if (state.connection !== rightState.connection) return null;

        if (!state.columns.includes(operation.leftKey)) return null;
        if (!rightState.columns.includes(operation.rightKey)) return null;

        const join = joinTypeToSql(dialect, operation.joinType);
        if (!join) return null;

        const leftCols = state.columns;
        const rightColsToInclude = rightState.columns.filter(
          (c) => c !== operation.rightKey || operation.rightKey !== operation.leftKey,
        );

        const leftNames = new Set(leftCols);
        const rightOut = rightColsToInclude.map((name) => ({
          source: name,
          out: leftNames.has(name) ? `${name}.right` : name,
        }));

        const selectList = [
          ...leftCols.map((name) => `l.${quoteIdentifier(name)} AS ${quoteIdentifier(name)}`),
          ...rightOut.map(({ source, out }) => `r.${quoteIdentifier(source)} AS ${quoteIdentifier(out)}`),
        ].join(", ");

        const mergedSql = `SELECT ${selectList} FROM (${state.fragment.sql}) AS l ${join} (${rightState.fragment.sql}) AS r ON l.${quoteIdentifier(operation.leftKey)} = r.${quoteIdentifier(operation.rightKey)}`;
        return {
          fragment: { sql: mergedSql, params: [...state.fragment.params, ...rightState.fragment.params] },
          columns: [...leftCols, ...rightOut.map((c) => c.out)],
          connection: state.connection,
        };
      }
      case "append": {
        if (!state.columns) return null;
        const queries = ctx.queries;
        if (!queries) return null;

        const baseColumns = state.columns;
        /** @type {SqlState[]} */
        const branches = [{ fragment: state.fragment, columns: baseColumns }];

        for (const id of operation.queries) {
          const q = queries[id];
          if (!q) return null;
          const compiled = this.compileQueryToSqlState(q, ctx, callStack);
          if (!compiled?.columns) return null;
          if (state.connection !== compiled.connection) return null;
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
          connection: state.connection,
        };
      }
      default:
        return null;
    }
  }
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
 * @param {Aggregation} agg
 * @param {(identifier: string) => string} quoteIdentifier
 * @returns {string}
 */
function aggregationToSql(agg, quoteIdentifier) {
  const alias = quoteIdentifier(agg.as ?? `${agg.op} of ${agg.column}`);
  const col = `t.${quoteIdentifier(agg.column)}`;
  switch (agg.op) {
    case "sum":
      return `COALESCE(SUM(${col}), 0) AS ${alias}`;
    case "count":
      return `COUNT(*) AS ${alias}`;
    case "average":
      return `AVG(${col}) AS ${alias}`;
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
 * @param {"inner" | "left" | "right" | "full"} joinType
 * @returns {string | null}
 */
function joinTypeToSql(dialect, joinType) {
  switch (joinType) {
    case "inner":
      return "INNER JOIN";
    case "left":
      return "LEFT JOIN";
    case "right":
      return dialect.name === "postgres" ? "RIGHT JOIN" : null;
    case "full":
      return dialect.name === "postgres" ? "FULL OUTER JOIN" : null;
    default: {
      /** @type {never} */
      const exhausted = joinType;
      throw new Error(`Unsupported joinType '${exhausted}'`);
    }
  }
}

/**
 * @param {string[]} base
 * @param {string[]} candidate
 * @returns {boolean}
 */
function columnsCompatible(base, candidate) {
  if (base.length !== candidate.length) return false;
  const set = new Set(base);
  return candidate.every((c) => set.has(c));
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
    default:
      return null;
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
  const trimmed = `TRIM(${dialect.castText(colRef)})`;
  const pattern = "'^[+-]?([0-9]+([.][0-9]*)?|[.][0-9]+)([eE][+-]?[0-9]+)?$'";

  switch (dialect.name) {
    case "postgres":
      return `CASE WHEN ${trimmed} = '' THEN NULL WHEN ${trimmed} ~ ${pattern} THEN CAST(${trimmed} AS ${sqlType}) ELSE NULL END`;
    case "mysql":
      return `CASE WHEN ${trimmed} = '' THEN NULL WHEN ${trimmed} REGEXP ${pattern} THEN CAST(${trimmed} AS ${sqlType}) ELSE NULL END`;
    case "sqlite":
      return null;
    default: {
      /** @type {never} */
      const exhausted = dialect.name;
      throw new Error(`Unsupported dialect '${exhausted}'`);
    }
  }
}

/**
 * Extremely small SQL expression compiler for `addColumn` folding.
 *
 * It supports a safe subset of the `compileRowFormula` syntax:
 * - column refs: `[Column Name]`
 * - numbers: `1`, `2.5`
 * - strings: `'text'` / `"text"` (no escapes)
 * - operators: `+ - * / % < <= > >= == != && || ! ? :`
 *
 * Unsupported tokens cause folding to break and the `addColumn` step executes
 * locally.
 *
 * @param {string} formula
 * @param {{
 *   alias: string;
 *   quoteIdentifier: (identifier: string) => string;
 *   params: unknown[];
 *   dialect: SqlDialect;
 *   knownColumns: string[] | null;
 * }} ctx
 * @returns {string | null}
 */
function compileFormulaToSql(formula, ctx) {
  let expr = formula.trim();
  if (expr.startsWith("=")) expr = expr.slice(1).trim();
  if (expr === "") return null;

  /** @type {Token[]} */
  let tokens;
  try {
    tokens = tokenizeFormula(expr);
  } catch {
    return null;
  }

  /** @type {Parser} */
  const parser = new Parser(tokens);
  let ast;
  try {
    ast = parser.parseExpression();
    parser.expect("eof");
  } catch {
    return null;
  }

  /** @param {unknown} value */
  const param = (value) => {
    if (isDate(value)) {
      ctx.params.push(ctx.dialect.formatDateParam(value));
      return "?";
    }
    ctx.params.push(value === undefined ? null : value);
    return "?";
  };

  /**
   * @param {ExprNode} node
   * @returns {string}
   */
  function toSql(node) {
    switch (node.type) {
      case "column": {
        if (ctx.knownColumns && !ctx.knownColumns.includes(node.name)) {
          throw new Error(`Unknown column '${node.name}'`);
        }
        return `${ctx.alias}.${ctx.quoteIdentifier(node.name)}`;
      }
      case "literal":
        if (node.value == null) return "NULL";
        return param(node.value);
      case "unary": {
        const rhs = toSql(node.arg);
        switch (node.op) {
          case "!":
            return `(NOT ${rhs})`;
          case "+":
            return `(+${rhs})`;
          case "-":
            return `(-${rhs})`;
          default:
            throw new Error(`Unsupported unary operator '${node.op}'`);
        }
      }
      case "binary": {
        const left = toSql(node.left);
        const right = toSql(node.right);
        const op = binaryOpToSql(node.op);
        return `(${left} ${op} ${right})`;
      }
      case "ternary": {
        const test = toSql(node.test);
        const cons = toSql(node.consequent);
        const alt = toSql(node.alternate);
        return `(CASE WHEN ${test} THEN ${cons} ELSE ${alt} END)`;
      }
      default: {
        /** @type {never} */
        const exhausted = node;
        throw new Error(`Unsupported node '${exhausted.type}'`);
      }
    }
  }

  try {
    return toSql(ast);
  } catch {
    return null;
  }
}

/**
 * @param {string} op
 * @returns {string}
 */
function binaryOpToSql(op) {
  switch (op) {
    case "&&":
      return "AND";
    case "||":
      return "OR";
    case "==":
    case "===":
      return "=";
    case "!=":
    case "!==":
      return "!=";
    default:
      return op;
  }
}

/**
 * @typedef {{
 *   type:
 *     | "number"
 *     | "string"
 *     | "column"
 *     | "identifier"
 *     | "operator"
 *     | "eof";
 *   value?: unknown;
 * }} Token
 */

/**
 * @param {string} input
 * @returns {Token[]}
 */
function tokenizeFormula(input) {
  /** @type {Token[]} */
  const tokens = [];
  let i = 0;
  while (i < input.length) {
    const ch = input[i];
    if (/\s/.test(ch)) {
      i += 1;
      continue;
    }

    if (ch === "[") {
      const end = input.indexOf("]", i + 1);
      if (end === -1) throw new Error("Unterminated column reference");
      const raw = input.slice(i + 1, end).trim();
      if (!raw) throw new Error("Empty column reference");
      tokens.push({ type: "column", value: raw });
      i = end + 1;
      continue;
    }

    if (ch === "'" || ch === '"') {
      const end = input.indexOf(ch, i + 1);
      if (end === -1) throw new Error("Unterminated string literal");
      const value = input.slice(i + 1, end);
      tokens.push({ type: "string", value });
      i = end + 1;
      continue;
    }

    const rest = input.slice(i);
    const numberMatch = rest.match(/^(?:\d+(?:\.\d+)?|\.\d+)/);
    if (numberMatch) {
      tokens.push({ type: "number", value: Number(numberMatch[0]) });
      i += numberMatch[0].length;
      continue;
    }

    const identMatch = rest.match(/^[A-Za-z_][A-Za-z0-9_]*/);
    if (identMatch) {
      tokens.push({ type: "identifier", value: identMatch[0] });
      i += identMatch[0].length;
      continue;
    }

    const opMatch = rest.match(/^(?:!==|===|!=|==|<=|>=|\|\||&&)/);
    if (opMatch) {
      tokens.push({ type: "operator", value: opMatch[0] });
      i += opMatch[0].length;
      continue;
    }

    if ("+-*/%()<>!?:".includes(ch)) {
      tokens.push({ type: "operator", value: ch });
      i += 1;
      continue;
    }

    throw new Error(`Unsupported character '${ch}'`);
  }
  tokens.push({ type: "eof" });
  return tokens;
}

/**
 * @typedef {{
 *   type: "literal";
 *   value: unknown;
 * } | {
 *   type: "column";
 *   name: string;
 * } | {
 *   type: "unary";
 *   op: string;
 *   arg: ExprNode;
 * } | {
 *   type: "binary";
 *   op: string;
 *   left: ExprNode;
 *   right: ExprNode;
 * } | {
 *   type: "ternary";
 *   test: ExprNode;
 *   consequent: ExprNode;
 *   alternate: ExprNode;
 * }} ExprNode
 */

class Parser {
  /**
   * @param {Token[]} tokens
   */
  constructor(tokens) {
    this.tokens = tokens;
    this.pos = 0;
  }

  /** @returns {Token} */
  peek() {
    return this.tokens[this.pos] ?? { type: "eof" };
  }

  /** @returns {Token} */
  next() {
    const tok = this.peek();
    this.pos += 1;
    return tok;
  }

  /**
   * @param {Token["type"]} type
   * @param {string} [value]
   * @returns {boolean}
   */
  match(type, value) {
    const tok = this.peek();
    if (tok.type !== type) return false;
    if (value !== undefined && tok.value !== value) return false;
    this.pos += 1;
    return true;
  }

  /**
   * @param {Token["type"]} type
   * @param {string} [value]
   */
  expect(type, value) {
    const tok = this.peek();
    if (tok.type !== type || (value !== undefined && tok.value !== value)) {
      throw new Error(`Expected ${value ?? type}`);
    }
    this.pos += 1;
  }

  /** @returns {ExprNode} */
  parseExpression() {
    return this.parseTernary();
  }

  /** @returns {ExprNode} */
  parseTernary() {
    let expr = this.parseOr();
    if (this.match("operator", "?")) {
      const consequent = this.parseTernary();
      this.expect("operator", ":");
      const alternate = this.parseTernary();
      expr = { type: "ternary", test: expr, consequent, alternate };
    }
    return expr;
  }

  /** @returns {ExprNode} */
  parseOr() {
    let expr = this.parseAnd();
    while (this.match("operator", "||")) {
      expr = { type: "binary", op: "||", left: expr, right: this.parseAnd() };
    }
    return expr;
  }

  /** @returns {ExprNode} */
  parseAnd() {
    let expr = this.parseEquality();
    while (this.match("operator", "&&")) {
      expr = { type: "binary", op: "&&", left: expr, right: this.parseEquality() };
    }
    return expr;
  }

  /** @returns {ExprNode} */
  parseEquality() {
    let expr = this.parseComparison();
    for (;;) {
      const tok = this.peek();
      if (tok.type !== "operator") break;
      if (!["==", "!=", "===", "!=="].includes(String(tok.value))) break;
      this.next();
      expr = { type: "binary", op: String(tok.value), left: expr, right: this.parseComparison() };
    }
    return expr;
  }

  /** @returns {ExprNode} */
  parseComparison() {
    let expr = this.parseAdditive();
    for (;;) {
      const tok = this.peek();
      if (tok.type !== "operator") break;
      if (!["<", "<=", ">", ">="].includes(String(tok.value))) break;
      this.next();
      expr = { type: "binary", op: String(tok.value), left: expr, right: this.parseAdditive() };
    }
    return expr;
  }

  /** @returns {ExprNode} */
  parseAdditive() {
    let expr = this.parseMultiplicative();
    for (;;) {
      const tok = this.peek();
      if (tok.type !== "operator") break;
      if (!["+", "-"].includes(String(tok.value))) break;
      this.next();
      expr = { type: "binary", op: String(tok.value), left: expr, right: this.parseMultiplicative() };
    }
    return expr;
  }

  /** @returns {ExprNode} */
  parseMultiplicative() {
    let expr = this.parseUnary();
    for (;;) {
      const tok = this.peek();
      if (tok.type !== "operator") break;
      if (!["*", "/", "%"].includes(String(tok.value))) break;
      this.next();
      expr = { type: "binary", op: String(tok.value), left: expr, right: this.parseUnary() };
    }
    return expr;
  }

  /** @returns {ExprNode} */
  parseUnary() {
    const tok = this.peek();
    if (tok.type === "operator" && ["!", "+", "-"].includes(String(tok.value))) {
      this.next();
      return { type: "unary", op: String(tok.value), arg: this.parseUnary() };
    }
    return this.parsePrimary();
  }

  /** @returns {ExprNode} */
  parsePrimary() {
    const tok = this.peek();
    switch (tok.type) {
      case "number":
        this.next();
        return { type: "literal", value: tok.value };
      case "string":
        this.next();
        return { type: "literal", value: tok.value };
      case "column":
        this.next();
        return { type: "column", name: String(tok.value) };
      case "identifier": {
        this.next();
        const ident = String(tok.value).toLowerCase();
        if (ident === "true") return { type: "literal", value: true };
        if (ident === "false") return { type: "literal", value: false };
        if (ident === "null") return { type: "literal", value: null };
        throw new Error(`Unsupported identifier '${tok.value}'`);
      }
      case "operator":
        if (tok.value === "(") {
          this.next();
          const expr = this.parseExpression();
          this.expect("operator", ")");
          return expr;
        }
        break;
      default:
        break;
    }
    throw new Error(`Unexpected token '${tok.type}'`);
  }
}
