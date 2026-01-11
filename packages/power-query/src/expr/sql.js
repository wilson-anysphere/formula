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
   * @param {ExprNode} node
   * @param {boolean} castString
   * @returns {string}
   */
  function toSql(node, castString) {
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
        const rhs = toSql(node.arg, false);
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
            return isNotEquals ? "(FALSE)" : "(TRUE)";
          }
          if (leftNull || rightNull) {
            const other = leftNull ? node.right : node.left;
            const otherSql = toSql(other, true);
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
        const left = toSql(node.left, castStringOperands);
        const right = toSql(node.right, castStringOperands);
        const op = binaryOpToSql(node.op);
        return `(${left} ${op} ${right})`;
      }
      case "ternary": {
        const test = toSql(node.test, false);
        const cons = toSql(node.consequent, true);
        const alt = toSql(node.alternate, true);
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

  return { sql: toSql(expr, true), params };
}
