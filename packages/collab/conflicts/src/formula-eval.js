import { colToNumber, numberToCol } from "./cell-ref.js";
import { parseFormula } from "./formula-parser.js";

/**
 * Best-effort evaluator used for conflict previews. It intentionally supports a
 * very small subset of Excel semantics.
 *
 * @param {string} formula
 * @param {object} [opts]
 * @param {(ref: { col: number, row: number }) => any} [opts.getCellValue] 0-based row/col.
 * @returns {{ ok: true, value: any } | { ok: false, error: string }}
 */
export function tryEvaluateFormula(formula, opts = {}) {
  try {
    const ast = parseFormula(formula);
    const value = evalAst(ast, opts);
    return { ok: true, value };
  } catch (err) {
    return { ok: false, error: err instanceof Error ? err.message : String(err) };
  }
}

/**
 * @param {import("./formula-parser.js").FormulaAst} ast
 * @param {object} opts
 * @returns {any}
 */
function evalAst(ast, opts) {
  switch (ast.type) {
    case "number":
      return Number.parseFloat(ast.value);
    case "cell": {
      if (!opts.getCellValue) throw new Error("No getCellValue() provided.");
      const col = colToNumber(ast.col);
      const row = ast.row - 1;
      return opts.getCellValue({ col, row });
    }
    case "range": {
      if (!opts.getCellValue) throw new Error("No getCellValue() provided.");
      const startCol = colToNumber(ast.start.col);
      const endCol = colToNumber(ast.end.col);
      const startRow = ast.start.row - 1;
      const endRow = ast.end.row - 1;

      /** @type {Array<any>} */
      const values = [];
      for (let r = Math.min(startRow, endRow); r <= Math.max(startRow, endRow); r += 1) {
        for (let c = Math.min(startCol, endCol); c <= Math.max(startCol, endCol); c += 1) {
          values.push(opts.getCellValue({ col: c, row: r }));
        }
      }
      return values;
    }
    case "name":
      throw new Error(`Cannot evaluate named reference: ${ast.name}`);
    case "unary": {
      const v = evalAst(ast.expr, opts);
      if (typeof v !== "number") throw new Error("Unary op on non-number");
      return ast.op === "-" ? -v : v;
    }
    case "binary": {
      const l = evalAst(ast.left, opts);
      const r = evalAst(ast.right, opts);
      if (typeof l !== "number" || typeof r !== "number") {
        throw new Error("Binary op on non-number");
      }
      switch (ast.op) {
        case "+":
          return l + r;
        case "-":
          return l - r;
        case "*":
          return l * r;
        case "/":
          return l / r;
        default:
          throw new Error(`Unsupported op: ${ast.op}`);
      }
    }
    case "call": {
      const name = ast.name.toUpperCase();
      const args = ast.args.map((a) => evalAst(a, opts));
      if (name === "SUM") {
        const flat = args.flatMap((v) => (Array.isArray(v) ? v : [v]));
        return flat.reduce((acc, v) => acc + (typeof v === "number" ? v : 0), 0);
      }
      if (name === "MIN") {
        const flat = args.flatMap((v) => (Array.isArray(v) ? v : [v])).filter((v) => typeof v === "number");
        if (flat.length === 0) throw new Error("MIN() with no numeric args");
        return Math.min(...flat);
      }
      if (name === "MAX") {
        const flat = args.flatMap((v) => (Array.isArray(v) ? v : [v])).filter((v) => typeof v === "number");
        if (flat.length === 0) throw new Error("MAX() with no numeric args");
        return Math.max(...flat);
      }
      throw new Error(`Unsupported function: ${name}`);
    }
    default:
      throw new Error(`Unknown AST node: ${/** @type {any} */ (ast).type}`);
  }
}

/**
 * @param {{col: number, row: number}} ref
 */
export function cellRefToA1(ref) {
  return `${numberToCol(ref.col)}${ref.row + 1}`;
}

