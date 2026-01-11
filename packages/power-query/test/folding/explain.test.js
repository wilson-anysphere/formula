import assert from "node:assert/strict";
import test from "node:test";

import { QueryFoldingEngine } from "../../src/folding/sql.js";

test("explain: addColumn uses the shared expr engine (no ReferenceError)", () => {
  const folding = new QueryFoldingEngine();
  const query = {
    id: "q_explain_add",
    name: "Explain addColumn",
    source: { type: "database", connection: {}, query: "SELECT * FROM sales", dialect: "postgres", columns: ["Sales"] },
    steps: [{ id: "s1", name: "Add", operation: { type: "addColumn", name: "Double", formula: "=[Sales] * 2" } }],
  };

  const result = folding.explain(query, { dialect: "postgres" });
  assert.equal(result.plan.type, "sql");
  assert.equal(result.steps.length, 1);
  assert.equal(result.steps[0].status, "folded");
});

test("explain: marks unsafe addColumn formulas as local with unsafe_formula reason", () => {
  const folding = new QueryFoldingEngine();
  const query = {
    id: "q_explain_unsafe",
    name: "Explain unsafe",
    source: { type: "database", connection: {}, query: "SELECT * FROM sales", dialect: "postgres", columns: ["Sales"] },
    steps: [{ id: "s1", name: "Add", operation: { type: "addColumn", name: "Bad", formula: "=Math.abs([Sales])" } }],
  };

  const result = folding.explain(query, { dialect: "postgres" });
  assert.equal(result.plan.type, "hybrid");
  assert.equal(result.steps[0].status, "local");
  assert.equal(result.steps[0].reason, "unsafe_formula");
});

