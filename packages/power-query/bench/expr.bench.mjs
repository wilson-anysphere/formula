import { performance } from "node:perf_hooks";

import { bindExprColumns, evaluateExpr, parseFormula } from "../src/expr/index.js";

function fmtMs(ms) {
  return `${ms.toFixed(1)}ms`;
}

/**
 * @param {number} rowCount
 */
function makeColumns(rowCount) {
  const a = new Float64Array(rowCount);
  const b = new Float64Array(rowCount);
  for (let i = 0; i < rowCount; i++) {
    a[i] = i % 10_000;
    b[i] = (i % 1_000) * 0.25;
  }
  return { a, b };
}

/**
 * @param {string} label
 * @param {number} rowCount
 */
function bench(label, rowCount) {
  const { a, b } = makeColumns(rowCount);
  const getColumnIndex = (name) => {
    if (name === "A") return 0;
    if (name === "B") return 1;
    throw new Error(`unknown column ${name}`);
  };

  const formula = "=[A] * 2 + [B] / 3";

  const parseStart = performance.now();
  const ast = parseFormula(formula);
  const bound = bindExprColumns(ast, getColumnIndex);
  const parseEnd = performance.now();

  const values = [0, 0];

  const evalStart = performance.now();
  let sum = 0;
  for (let i = 0; i < rowCount; i++) {
    values[0] = a[i];
    values[1] = b[i];
    const v = evaluateExpr(bound, values);
    // Touch the value so work isn't optimized away.
    if (typeof v === "number") sum += v;
  }
  const evalEnd = performance.now();

  const baselineStart = performance.now();
  let baselineSum = 0;
  for (let i = 0; i < rowCount; i++) {
    const v = a[i] * 2 + b[i] / 3;
    baselineSum += v;
  }
  const baselineEnd = performance.now();

  console.log(
    `${label}: parse+bind=${fmtMs(parseEnd - parseStart)} eval=${fmtMs(evalEnd - evalStart)} baseline=${fmtMs(baselineEnd - baselineStart)} checksum=${sum.toFixed(1)} baselineChecksum=${baselineSum.toFixed(1)}`,
  );
}

const SMALL = 100_000;
const LARGE = 1_000_000;

console.log("Power Query expression benchmark");
console.log(`Node ${process.version}`);
console.log("");

bench(`expr ${SMALL.toLocaleString()} rows`, SMALL);
bench(`expr ${LARGE.toLocaleString()} rows`, LARGE);

