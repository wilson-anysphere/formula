import "./ast.js";

import { parseExpr } from "./parser.js";
import { bindExprColumns, evaluateExpr } from "./evaluate.js";
import { compileExprToSql } from "./sql.js";
import { collectExprColumnRefs } from "./refs.js";

export { parseExpr, bindExprColumns, evaluateExpr, compileExprToSql, collectExprColumnRefs };

/**
 * Strip the optional leading `=` and normalize whitespace.
 *
 * @param {string} formula
 * @returns {string}
 */
export function normalizeFormulaText(formula) {
  let expr = formula.trim();
  if (expr.startsWith("=")) expr = expr.slice(1).trim();
  if (expr === "") {
    throw new Error("Formula cannot be empty");
  }
  return expr;
}

/**
 * Parse a formula string into an expression AST.
 *
 * @param {string} formula
 * @returns {import("./ast.js").ExprNode}
 */
export function parseFormula(formula) {
  return parseExpr(normalizeFormulaText(formula));
}

