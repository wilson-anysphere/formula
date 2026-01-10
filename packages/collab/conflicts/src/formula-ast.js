/**
 * @param {import("./formula-parser.js").FormulaAst} a
 * @param {import("./formula-parser.js").FormulaAst} b
 * @returns {boolean}
 */
export function astEquals(a, b) {
  if (a.type !== b.type) return false;

  switch (a.type) {
    case "number":
      return a.value === /** @type {any} */ (b).value;
    case "cell":
      return a.col === /** @type {any} */ (b).col && a.row === /** @type {any} */ (b).row;
    case "range":
      return astEquals(a.start, /** @type {any} */ (b).start) && astEquals(a.end, /** @type {any} */ (b).end);
    case "name":
      return a.name === /** @type {any} */ (b).name;
    case "call": {
      const bb = /** @type {any} */ (b);
      if (a.name !== bb.name) return false;
      if (a.args.length !== bb.args.length) return false;
      for (let i = 0; i < a.args.length; i += 1) {
        if (!astEquals(a.args[i], bb.args[i])) return false;
      }
      return true;
    }
    case "unary":
      return a.op === /** @type {any} */ (b).op && astEquals(a.expr, /** @type {any} */ (b).expr);
    case "binary":
      return (
        a.op === /** @type {any} */ (b).op &&
        astEquals(a.left, /** @type {any} */ (b).left) &&
        astEquals(a.right, /** @type {any} */ (b).right)
      );
    default:
      return false;
  }
}

/**
 * Returns true if `needle` is an AST subtree of `haystack`.
 *
 * Used as a heuristic for "extension/subset" formula relationships
 * (see docs/06-collaboration.md).
 *
 * @param {import("./formula-parser.js").FormulaAst} needle
 * @param {import("./formula-parser.js").FormulaAst} haystack
 * @returns {boolean}
 */
export function isAstSubtree(needle, haystack) {
  if (astEquals(needle, haystack)) return true;

  switch (haystack.type) {
    case "range":
      return isAstSubtree(needle, haystack.start) || isAstSubtree(needle, haystack.end);
    case "call":
      return haystack.args.some((arg) => isAstSubtree(needle, arg));
    case "unary":
      return isAstSubtree(needle, haystack.expr);
    case "binary":
      return isAstSubtree(needle, haystack.left) || isAstSubtree(needle, haystack.right);
    default:
      return false;
  }
}

