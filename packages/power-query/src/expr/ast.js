/**
 * Minimal expression language used by Power Query `addColumn` formulas.
 *
 * The language is intentionally small and deterministic so it can be:
 *  - evaluated locally without `eval` / `new Function`
 *  - compiled to SQL for folding
 *  - analyzed to extract referenced columns (Parquet projection planning)
 *
 * This file holds the shared JSDoc types used across the expression engine.
 */

/**
 * @typedef {{
 *   start: number;
 *   end: number;
 }} ExprSpan
 */

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
 *   span: ExprSpan;
 * }} ExprToken
 */

/**
 * @typedef {{
 *   type: "literal";
 *   value: null | boolean | number | string;
 * }} LiteralExpr
 */

/**
 * @typedef {{
 *   type: "column";
 *   name: string;
 *   // Optional binding for fast evaluation.
 *   index?: number;
 * }} ColumnExpr
 */

/**
 * @typedef {{
 *   type: "unary";
 *   op: "!" | "+" | "-";
 *   arg: ExprNode;
 * }} UnaryExpr
 */

/**
 * @typedef {{
 *   type: "binary";
 *   op: string;
 *   left: ExprNode;
 *   right: ExprNode;
 * }} BinaryExpr
 */

/**
 * @typedef {{
 *   type: "ternary";
 *   test: ExprNode;
 *   consequent: ExprNode;
 *   alternate: ExprNode;
 * }} TernaryExpr
 */

/**
 * @typedef {{
 *   type: "call";
 *   callee: string;
 *   args: ExprNode[];
 * }} CallExpr
 */

/**
 * @typedef {LiteralExpr | ColumnExpr | UnaryExpr | BinaryExpr | TernaryExpr | CallExpr} ExprNode
 */

export {};

