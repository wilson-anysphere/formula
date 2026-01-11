/**
 * Typed AST definitions for the pragmatic M language subset supported by
 * `packages/power-query`.
 *
 * This is intentionally JS + JSDoc so it can run in Node without a TS build.
 */

/**
 * @typedef {import("./errors.js").MLocation} MLocation
 * @typedef {import("./errors.js").MSpan} MSpan
 */

/**
 * @typedef {{
 *   type: "Program";
 *   expression: MExpression;
 *   span: MSpan;
 * }} MProgram
 */

/**
 * @typedef {{
 *   type: "LetExpression";
 *   bindings: MLetBinding[];
 *   body: MExpression;
 *   span: MSpan;
 * }} MLetExpression
 */

/**
 * @typedef {{
 *   name: MIdentifierName;
 *   value: MExpression;
 *   span: MSpan;
 * }} MLetBinding
 */

/**
 * @typedef {{
 *   name: string;
 *   quoted: boolean;
 *   span: MSpan;
 * }} MIdentifierName
 */

/**
 * @typedef {{
 *   type: "Identifier";
 *   parts: string[];
 *   quoted?: boolean;
 *   span: MSpan;
 * }} MIdentifier
 */

/**
 * @typedef {{
 *   type: "Literal";
 *   value: null | boolean | number | string | Date;
 *   literalType: "null" | "boolean" | "number" | "string" | "date";
 *   span: MSpan;
 * }} MLiteral
 */

/**
 * @typedef {{
 *   type: "ListExpression";
 *   elements: MExpression[];
 *   span: MSpan;
 * }} MListExpression
 */

/**
 * @typedef {{
 *   type: "RecordExpression";
 *   fields: { key: string; value: MExpression; span: MSpan }[];
 *   span: MSpan;
 * }} MRecordExpression
 */

/**
 * Field access (`x[Field]`), or implicit field access (`[Field]` inside `each`).
 * @typedef {{
 *   type: "FieldAccessExpression";
 *   base: MExpression | null;
 *   field: string;
 *   span: MSpan;
 * }} MFieldAccessExpression
 */

/**
 * Item access (`x{0}` or `x{[Key="Value"]}`)
 * @typedef {{
 *   type: "ItemAccessExpression";
 *   base: MExpression;
 *   key: MExpression;
 *   span: MSpan;
 * }} MItemAccessExpression
 */

/**
 * @typedef {{
 *   type: "CallExpression";
 *   callee: MExpression;
 *   args: MExpression[];
 *   span: MSpan;
 * }} MCallExpression
 */

/**
 * @typedef {{
 *   type: "EachExpression";
 *   body: MExpression;
 *   span: MSpan;
 * }} MEachExpression
 */

/**
 * @typedef {{
 *   type: "UnaryExpression";
 *   operator: "not" | "+" | "-";
 *   argument: MExpression;
 *   span: MSpan;
 * }} MUnaryExpression
 */

/**
 * @typedef {{
 *   type: "BinaryExpression";
 *   operator: "and" | "or" | "=" | "<>" | "<" | "<=" | ">" | ">=" | "+" | "-" | "*" | "/" | "&";
 *   left: MExpression;
 *   right: MExpression;
 *   span: MSpan;
 * }} MBinaryExpression
 */

/**
 * `type number`, `type text`, etc.
 * @typedef {{
 *   type: "TypeExpression";
 *   name: string;
 *   span: MSpan;
 * }} MTypeExpression
 */

/**
 * @typedef {{
 *   type: "ParenthesizedExpression";
 *   expression: MExpression;
 *   span: MSpan;
 * }} MParenthesizedExpression
 */

/**
 * @typedef {MLetExpression | MIdentifier | MLiteral | MListExpression | MRecordExpression | MFieldAccessExpression | MItemAccessExpression | MCallExpression | MEachExpression | MUnaryExpression | MBinaryExpression | MTypeExpression | MParenthesizedExpression} MExpression
 */

/**
 * @param {MLocation} start
 * @param {MLocation} end
 * @returns {MSpan}
 */
export function span(start, end) {
  return { start, end };
}

export {};

