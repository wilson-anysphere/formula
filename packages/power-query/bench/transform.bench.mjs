import { performance } from "node:perf_hooks";

import { arrowTableFromColumns } from "../../data-io/src/index.js";

import { ArrowTableAdapter } from "../src/arrowTable.js";
import { DataTable } from "../src/table.js";
import { applyOperation } from "../src/steps.js";

function fmtMs(ms) {
  return `${ms.toFixed(1)}ms`;
}

function mem() {
  const { heapUsed } = process.memoryUsage();
  return `${(heapUsed / 1024 / 1024).toFixed(1)}MB`;
}

/**
 * @param {number} rowCount
 */
function makeArrowTable(rowCount) {
  const regions = ["East", "West", "North", "South"];
  const products = ["A", "B", "C", "D", "E"];

  const region = new Array(rowCount);
  const product = new Array(rowCount);
  const sales = new Float64Array(rowCount);

  for (let i = 0; i < rowCount; i++) {
    region[i] = regions[i & 3];
    product[i] = products[i % products.length];
    sales[i] = (i % 10_000) * 0.5;
  }

  return new ArrowTableAdapter(
    arrowTableFromColumns({
      Region: region,
      Product: product,
      Sales: sales,
    }),
  );
}

/**
 * @param {number} rowCount
 */
function makeDataTable(rowCount) {
  const regions = ["East", "West", "North", "South"];
  const products = ["A", "B", "C", "D", "E"];
  const columns = [
    { name: "Region", type: "string" },
    { name: "Product", type: "string" },
    { name: "Sales", type: "number" },
  ];

  const rows = new Array(rowCount);
  for (let i = 0; i < rowCount; i++) {
    rows[i] = [regions[i & 3], products[i % products.length], (i % 10_000) * 0.5];
  }
  return new DataTable(columns, rows);
}

function runPipeline(table) {
  let current = table;
  current = applyOperation(current, {
    type: "filterRows",
    predicate: { type: "comparison", column: "Sales", operator: "greaterThan", value: 2_000 },
  });
  current = applyOperation(current, { type: "removeColumns", columns: ["Product"] });
  current = applyOperation(current, { type: "sortRows", sortBy: [{ column: "Sales", direction: "descending" }] });
  current = applyOperation(current, {
    type: "groupBy",
    groupColumns: ["Region"],
    aggregations: [{ column: "Sales", op: "sum", as: "Total Sales" }],
  });
  current = applyOperation(current, { type: "sortRows", sortBy: [{ column: "Total Sales", direction: "descending" }] });
  return current;
}

async function bench(label, makeTable) {
  const beforeMem = mem();
  const start = performance.now();
  const table = makeTable();
  const created = performance.now();
  const out = runPipeline(table);
  const end = performance.now();

  // Touch the result so work isn't optimized away.
  const head = out.head(3).toGrid();
  void head;

  console.log(
    `${label}: create=${fmtMs(created - start)} run=${fmtMs(end - created)} total=${fmtMs(end - start)} rowsOut=${out.rowCount} heap=${beforeMem} -> ${mem()}`,
  );
}

const SMALL = 100_000;
const LARGE = 1_000_000;

console.log("Power Query benchmark (JS, single-threaded)");
console.log(`Node ${process.version}`);
console.log("");

await bench(`Arrow ${SMALL.toLocaleString()} rows`, () => makeArrowTable(SMALL));
await bench(`DataTable ${SMALL.toLocaleString()} rows`, () => makeDataTable(SMALL));

console.log("");
await bench(`Arrow ${LARGE.toLocaleString()} rows`, () => makeArrowTable(LARGE));

console.log("");
console.log("Note: DataTable 1M rows is intentionally skipped (row-array overhead is too large).");
