import "./ast.js";

import { parseDateLiteral } from "./date.js";

/**
 * @typedef {import("./ast.js").ExprNode} ExprNode
 * @typedef {import("../folding/dialect.js").SqlDialect} SqlDialect
 */

/**
 * @typedef {{
 *   sql: string;
 *   params: unknown[];
 * }} SqlFragment
 */

/**
 * @param {unknown} value
 * @returns {value is Date}
 */
function isDate(value) {
  return value instanceof Date && !Number.isNaN(value.getTime());
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
 * @param {SqlDialect["name"]} dialect
 * @returns {{ boolean: string, number: string, date: string }}
 */
function sqlTypesForDialect(dialect) {
  switch (dialect) {
    case "postgres":
      return { boolean: "BOOLEAN", number: "DOUBLE PRECISION", date: "TIMESTAMPTZ" };
    case "mysql":
      // MySQL's BOOLEAN is an alias for TINYINT(1), but CAST(... AS BOOLEAN)
      // is not consistently supported across drivers/versions. Use a numeric
      // cast to keep the generated SQL conservative.
      return { boolean: "SIGNED", number: "DOUBLE", date: "DATETIME" };
    case "sqlite":
      return { boolean: "INTEGER", number: "REAL", date: "TEXT" };
    case "sqlserver":
      return { boolean: "BIT", number: "FLOAT", date: "DATETIME2" };
    default: {
      /** @type {never} */
      const exhausted = dialect;
      throw new Error(`Unsupported dialect '${exhausted}'`);
    }
  }
}

/**
 * Compile an expression AST to a SQL fragment.
 *
 * All non-null literals are parameterized (`?`) and returned via `params` to
 * avoid SQL injection via formulas.
 *
 * @param {ExprNode} expr
 * @param {{
 *   alias: string;
 *   quoteIdentifier: (identifier: string) => string;
 *   dialect: SqlDialect;
 *   knownColumns: string[] | null;
 * }} ctx
 * @returns {SqlFragment}
 */
export function compileExprToSql(expr, ctx) {
  /** @type {unknown[]} */
  const params = [];
  const types = sqlTypesForDialect(ctx.dialect.name);
  const isSqlServer = ctx.dialect.name === "sqlserver";

  /** @param {unknown} value */
  const param = (value) => {
    if (isDate(value)) {
      params.push(ctx.dialect.formatDateParam(value));
      return "?";
    }
    params.push(value === undefined ? null : value);
    return "?";
  };

  /**
   * @param {string} placeholder
   * @param {string} sqlType
   * @returns {string}
   */
  const castParam = (placeholder, sqlType) => {
    return `CAST(${placeholder} AS ${sqlType})`;
  };

  /**
   * @param {string} predicateSql
   * @returns {string}
   */
  const sqlServerBoolToBit = (predicateSql) => {
    return `(CASE WHEN ${predicateSql} THEN CAST(1 AS BIT) ELSE CAST(0 AS BIT) END)`;
  };

  /**
   * @param {ExprNode} node
   * @param {boolean} castString
   * @returns {string}
   */
  function toSqlLegacy(node, castString) {
    switch (node.type) {
      case "value":
        throw new Error("Value placeholder '_' is not supported in SQL folding");
      case "column": {
        if (ctx.knownColumns && !ctx.knownColumns.includes(node.name)) {
          throw new Error(`Unknown column '${node.name}'`);
        }
        return `${ctx.alias}.${ctx.quoteIdentifier(node.name)}`;
      }
      case "literal":
        if (node.value == null) return "NULL";
        if (typeof node.value === "string") {
          // Avoid "could not determine data type of parameter $n" for Postgres
          // when the literal appears without a type context (e.g. SELECT ?).
          const placeholder = param(node.value);
          return castString ? ctx.dialect.castText(placeholder) : placeholder;
        }
        if (typeof node.value === "number") {
          return castParam(param(node.value), types.number);
        }
        if (typeof node.value === "boolean") {
          return castParam(param(node.value), types.boolean);
        }
        throw new Error(`Unsupported literal type '${typeof node.value}'`);
      case "unary": {
        const rhs = toSqlLegacy(node.arg, false);
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
        if (["==", "===", "!=", "!=="].includes(node.op)) {
          const leftNull = node.left.type === "literal" && node.left.value == null;
          const rightNull = node.right.type === "literal" && node.right.value == null;
          if (leftNull && rightNull) {
            // Match JS semantics: `null == null` is true, `null != null` is false.
            const isNotEquals = node.op === "!=" || node.op === "!==";
            if (ctx.dialect.name === "sqlserver") {
              return isNotEquals ? "(1=0)" : "(1=1)";
            }
            return isNotEquals ? "(FALSE)" : "(TRUE)";
          }
          if (leftNull || rightNull) {
            const other = leftNull ? node.right : node.left;
            const otherSql = toSqlLegacy(other, true);
            const isNotEquals = node.op === "!=" || node.op === "!==";
            return `(${otherSql} IS ${isNotEquals ? "NOT " : ""}NULL)`;
          }
        }

        const leftIsStringLiteral = node.left.type === "literal" && typeof node.left.value === "string";
        const rightIsStringLiteral = node.right.type === "literal" && typeof node.right.value === "string";
        if (node.op === "+" && (leftIsStringLiteral || rightIsStringLiteral)) {
          throw new Error("String concatenation via '+' is not supported in SQL folding");
        }

        const castStringOperands = leftIsStringLiteral && rightIsStringLiteral;
        const left = toSqlLegacy(node.left, castStringOperands);
        const right = toSqlLegacy(node.right, castStringOperands);
        const op = binaryOpToSql(node.op);
        return `(${left} ${op} ${right})`;
      }
      case "ternary": {
        const test = toSqlLegacy(node.test, false);
        const cons = toSqlLegacy(node.consequent, true);
        const alt = toSqlLegacy(node.alternate, true);
        return `(CASE WHEN ${test} THEN ${cons} ELSE ${alt} END)`;
      }
      case "call":
        if (node.callee.toLowerCase() === "date") {
          if (node.args.length !== 1) {
            throw new Error("date() expects exactly 1 argument");
          }
          const arg0 = node.args[0];
          if (arg0.type !== "literal" || typeof arg0.value !== "string") {
            throw new Error('date() expects a string literal like date("2020-01-01")');
          }
          return castParam(param(parseDateLiteral(arg0.value)), types.date);
        }
        throw new Error(`Unsupported function '${node.callee}'`);
      default: {
        /** @type {never} */
        const exhausted = node;
        throw new Error(`Unsupported node '${exhausted.type}'`);
      }
    }
  }

  /**
   * Compile an expression into a SQL Server predicate (boolean expression).
   * SQL Server predicates cannot use BIT values directly (e.g. `CASE WHEN t.[Flag] THEN ...` is invalid);
   * they must be comparisons like `t.[Flag] = 1`.
   *
   * @param {ExprNode} node
   * @returns {string}
   */
  function toSqlServerPredicate(node) {
    switch (node.type) {
      case "value":
        throw new Error("Value placeholder '_' is not supported in SQL folding");
      case "column": {
        if (ctx.knownColumns && !ctx.knownColumns.includes(node.name)) {
          throw new Error(`Unknown column '${node.name}'`);
        }
        const col = `${ctx.alias}.${ctx.quoteIdentifier(node.name)}`;
        return `(${col} = 1)`;
      }
      case "literal":
        if (node.value == null) return "(1=0)";
        if (typeof node.value === "boolean") return node.value ? "(1=1)" : "(1=0)";
        // Fall back to SQL Server truthiness for numbers; other types are not supported
        // in predicate positions (matching our conservative folding approach).
        if (typeof node.value === "number") return `(${castParam(param(node.value), types.number)} <> 0)`;
        throw new Error("Unsupported predicate expression in SQL Server folding");
      case "unary": {
        if (node.op !== "!") throw new Error(`Unsupported unary operator '${node.op}' in SQL Server predicate context`);
        return `(NOT ${toSqlServerPredicate(node.arg)})`;
      }
      case "binary": {
        if (node.op === "&&" || node.op === "||") {
          const left = toSqlServerPredicate(node.left);
          const right = toSqlServerPredicate(node.right);
          const op = node.op === "&&" ? "AND" : "OR";
          return `(${left} ${op} ${right})`;
        }

        if (["==", "===", "!=", "!=="].includes(node.op)) {
          const leftNull = node.left.type === "literal" && node.left.value == null;
          const rightNull = node.right.type === "literal" && node.right.value == null;
          if (leftNull && rightNull) {
            const isNotEquals = node.op === "!=" || node.op === "!==";
            return isNotEquals ? "(1=0)" : "(1=1)";
          }
          if (leftNull || rightNull) {
            const other = leftNull ? node.right : node.left;
            const otherSql = toSqlServerValue(other, true);
            const isNotEquals = node.op === "!=" || node.op === "!==";
            return `(${otherSql} IS ${isNotEquals ? "NOT " : ""}NULL)`;
          }
        }

        const leftIsStringLiteral = node.left.type === "literal" && typeof node.left.value === "string";
        const rightIsStringLiteral = node.right.type === "literal" && typeof node.right.value === "string";
        if (node.op === "+" && (leftIsStringLiteral || rightIsStringLiteral)) {
          throw new Error("String concatenation via '+' is not supported in SQL folding");
        }

        const castStringOperands = leftIsStringLiteral && rightIsStringLiteral;
        const left = toSqlServerValue(node.left, castStringOperands);
        const right = toSqlServerValue(node.right, castStringOperands);
        const op = binaryOpToSql(node.op);

        if (!["=", "!=", "<", "<=", ">", ">="].includes(op) && op !== "AND" && op !== "OR") {
          throw new Error("Unsupported predicate operator in SQL Server folding");
        }
        return `(${left} ${op} ${right})`;
      }
      case "ternary": {
        // Searched CASE yields a scalar, not a predicate.
        throw new Error("Ternary expressions are not supported in SQL Server predicate context");
      }
      case "call":
        throw new Error(`Unsupported function '${node.callee}' in SQL Server predicate context`);
      default: {
        /** @type {never} */
        const exhausted = node;
        throw new Error(`Unsupported node '${exhausted.type}'`);
      }
    }
  }

  /**
   * @param {ExprNode} node
   * @param {boolean} castString
   * @returns {string}
   */
  function toSqlServerValue(node, castString) {
    switch (node.type) {
      case "value":
        throw new Error("Value placeholder '_' is not supported in SQL folding");
      case "column":
      case "call":
        // Defer to the legacy implementation for column references and supported functions (`date()`).
        return toSqlLegacy(node, castString);
      case "literal":
        if (node.value == null) return "NULL";
        if (typeof node.value === "boolean") return `CAST(${node.value ? 1 : 0} AS BIT)`;
        return toSqlLegacy(node, castString);
      case "unary": {
        if (node.op === "!") {
          return sqlServerBoolToBit(toSqlServerPredicate(node));
        }
        return toSqlLegacy(node, castString);
      }
      case "binary": {
        if (node.op === "&&" || node.op === "||") {
          return sqlServerBoolToBit(toSqlServerPredicate(node));
        }

        // Comparisons yield predicates; turn them into BIT for value context.
        if (["==", "===", "!=", "!==", "<", "<=", ">", ">="].includes(node.op)) {
          return sqlServerBoolToBit(toSqlServerPredicate(node));
        }

        return toSqlLegacy(node, castString);
      }
      case "ternary": {
        const test = toSqlServerPredicate(node.test);
        const cons = toSqlServerValue(node.consequent, true);
        const alt = toSqlServerValue(node.alternate, true);
        return `(CASE WHEN ${test} THEN ${cons} ELSE ${alt} END)`;
      }
      default: {
        /** @type {never} */
        const exhausted = node;
        throw new Error(`Unsupported node '${exhausted.type}'`);
      }
    }
  }

  const sql = isSqlServer ? toSqlServerValue(expr, true) : toSqlLegacy(expr, true);
  return { sql, params };
}
