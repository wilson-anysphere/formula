import { colToNumber, numberToCol } from "./cell-ref.js";
import { parseFormula } from "./formula-parser.js";

// Conflict previews should never attempt to materialize Excel-scale ranges into JS arrays.
// Keep evaluation bounded so large formulas like `=SUM(A:A)` don't OOM the renderer.
const DEFAULT_MAX_EVAL_RANGE_CELLS = 200_000;

/**
 * Best-effort evaluator used for conflict previews. It intentionally supports a
 * very small subset of Excel semantics.
 *
 * @param {string} formula
 * @param {object} [opts]
 * @param {(ref: { col: number, row: number }) => any} [opts.getCellValue] 0-based row/col.
 * @param {number} [opts.maxRangeCells] Maximum number of cells to materialize for a range reference.
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
  const maxRangeCells = Number.isFinite(opts.maxRangeCells) && opts.maxRangeCells > 0 ? Math.floor(opts.maxRangeCells) : DEFAULT_MAX_EVAL_RANGE_CELLS;

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

      const minRow = Math.min(startRow, endRow);
      const maxRow = Math.max(startRow, endRow);
      const minCol = Math.min(startCol, endCol);
      const maxCol = Math.max(startCol, endCol);
      const rows = maxRow - minRow + 1;
      const cols = maxCol - minCol + 1;
      const cellCount = rows * cols;
      if (!Number.isFinite(cellCount) || cellCount < 0) {
        throw new Error(`Invalid range size (rows=${rows}, cols=${cols}).`);
      }
      if (cellCount > maxRangeCells) {
        throw new Error(`Range too large to evaluate (${cellCount} cells; max=${maxRangeCells}).`);
      }

      /** @type {Array<any>} */
      const values = [];
      for (let r = minRow; r <= maxRow; r += 1) {
        for (let c = minCol; c <= maxCol; c += 1) {
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
        let sum = 0;
        for (const arg of args) {
          if (Array.isArray(arg)) {
            for (const v of arg) sum += typeof v === "number" ? v : 0;
          } else {
            sum += typeof arg === "number" ? arg : 0;
          }
        }
        return sum;
      }
      if (name === "MIN") {
        let min = Number.POSITIVE_INFINITY;
        let hasNumber = false;
        for (const arg of args) {
          if (Array.isArray(arg)) {
            for (const v of arg) {
              if (typeof v !== "number") continue;
              if (!hasNumber) {
                min = v;
                hasNumber = true;
              } else if (v < min) {
                min = v;
              }
            }
          } else if (typeof arg === "number") {
            if (!hasNumber) {
              min = arg;
              hasNumber = true;
            } else if (arg < min) {
              min = arg;
            }
          }
        }
        if (!hasNumber) throw new Error("MIN() with no numeric args");
        return min;
      }
      if (name === "MAX") {
        let max = Number.NEGATIVE_INFINITY;
        let hasNumber = false;
        for (const arg of args) {
          if (Array.isArray(arg)) {
            for (const v of arg) {
              if (typeof v !== "number") continue;
              if (!hasNumber) {
                max = v;
                hasNumber = true;
              } else if (v > max) {
                max = v;
              }
            }
          } else if (typeof arg === "number") {
            if (!hasNumber) {
              max = arg;
              hasNumber = true;
            } else if (arg > max) {
              max = arg;
            }
          }
        }
        if (!hasNumber) throw new Error("MAX() with no numeric args");
        return max;
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
