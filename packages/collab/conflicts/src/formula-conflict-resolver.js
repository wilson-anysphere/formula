import { normalizeFormula, parseFormula } from "../../../versioning/src/index.js";

/**
 * @typedef {object} FormulaConflictDecision
 * @property {"equivalent"|"prefer-local"|"prefer-remote"|"conflict"} kind
 * @property {string} reason
 * @property {string} [chosenFormula]
 */

/**
 * Implements the conflict resolver described in docs/06-collaboration.md.
 *
 * - If formulas are AST-equivalent, auto-resolve.
 * - If one is an extension/subset (subtree), prefer the extension.
 * - Otherwise surface a conflict UI.
 *
 * @param {object} input
 * @param {string|null|undefined} input.localFormula
 * @param {string|null|undefined} input.remoteFormula
 * @returns {FormulaConflictDecision}
 */
export function resolveFormulaConflict(input) {
  const localFormula = (input.localFormula ?? "").trim();
  const remoteFormula = (input.remoteFormula ?? "").trim();

  // Use the existing semantic normalization helper (Task 3) for equivalence.
  const normLocal = normalizeFormula(localFormula);
  const normRemote = normalizeFormula(remoteFormula);
  if (normLocal === normRemote) {
    return { kind: "equivalent", chosenFormula: remoteFormula, reason: "ast-equivalent" };
  }

  const parsedLocal = tryParse(localFormula);
  const parsedRemote = tryParse(remoteFormula);

  if (parsedLocal && parsedRemote) {
    if (isAstSubtree(parsedLocal, parsedRemote) && !astEquals(parsedLocal, parsedRemote)) {
      return { kind: "prefer-remote", chosenFormula: remoteFormula, reason: "remote-is-extension" };
    }

    if (isAstSubtree(parsedRemote, parsedLocal) && !astEquals(parsedLocal, parsedRemote)) {
      return { kind: "prefer-local", chosenFormula: localFormula, reason: "local-is-extension" };
    }
  }

  // Fallback: crude substring heuristic if parsing fails.
  const textLocal = normalizeFormulaText(localFormula);
  const textRemote = normalizeFormulaText(remoteFormula);

  if (textRemote.includes(textLocal)) {
    return { kind: "prefer-remote", chosenFormula: remoteFormula, reason: "remote-contains-local" };
  }

  if (textLocal.includes(textRemote)) {
    return { kind: "prefer-local", chosenFormula: localFormula, reason: "local-contains-remote" };
  }

  return { kind: "conflict", reason: "non-equivalent" };
}

/**
 * @param {string} formula
 * @returns {import("../../../versioning/src/formula/parse.js").AstNode|null}
 */
function tryParse(formula) {
  try {
    if (!formula) return null;
    return parseFormula(formula);
  } catch {
    return null;
  }
}

/**
 * @param {string} formula
 */
function normalizeFormulaText(formula) {
  const stripped = formula.trim().replace(/^\s*=\s*/, "");
  return stripped.replaceAll(/\s+/g, "").toUpperCase();
}

/**
 * @param {import("../../../versioning/src/formula/parse.js").AstNode} a
 * @param {import("../../../versioning/src/formula/parse.js").AstNode} b
 */
function astEquals(a, b) {
  if (a.type !== b.type) return false;

  switch (a.type) {
    case "Number":
      return a.value === /** @type {any} */ (b).value;
    case "String":
      return a.value === /** @type {any} */ (b).value;
    case "Cell":
      return a.ref.toUpperCase() === /** @type {any} */ (b).ref.toUpperCase();
    case "Name":
      return a.name.toUpperCase() === /** @type {any} */ (b).name.toUpperCase();
    case "Unary":
      return a.op === /** @type {any} */ (b).op && astEquals(a.expr, /** @type {any} */ (b).expr);
    case "Percent":
      return astEquals(a.expr, /** @type {any} */ (b).expr);
    case "Range":
      return astEquals(a.start, /** @type {any} */ (b).start) && astEquals(a.end, /** @type {any} */ (b).end);
    case "Binary":
      return (
        a.op === /** @type {any} */ (b).op &&
        astEquals(a.left, /** @type {any} */ (b).left) &&
        astEquals(a.right, /** @type {any} */ (b).right)
      );
    case "Function": {
      const bb = /** @type {any} */ (b);
      if (a.name.toUpperCase() !== bb.name.toUpperCase()) return false;
      if (a.args.length !== bb.args.length) return false;
      for (let i = 0; i < a.args.length; i += 1) {
        if (!astEquals(a.args[i], bb.args[i])) return false;
      }
      return true;
    }
    default:
      return false;
  }
}

/**
 * @param {import("../../../versioning/src/formula/parse.js").AstNode} needle
 * @param {import("../../../versioning/src/formula/parse.js").AstNode} haystack
 */
function isAstSubtree(needle, haystack) {
  if (astEquals(needle, haystack)) return true;

  switch (haystack.type) {
    case "Unary":
      return isAstSubtree(needle, haystack.expr);
    case "Percent":
      return isAstSubtree(needle, haystack.expr);
    case "Binary":
      return isAstSubtree(needle, haystack.left) || isAstSubtree(needle, haystack.right);
    case "Range":
      return isAstSubtree(needle, haystack.start) || isAstSubtree(needle, haystack.end);
    case "Function":
      return haystack.args.some((arg) => isAstSubtree(needle, arg));
    default:
      return false;
  }
}
