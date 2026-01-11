import assert from "node:assert/strict";
import test from "node:test";

import { QueryFoldingEngine } from "../../src/folding/sql.js";

test("explain: fully foldable query marks every step as folded", () => {
  const folding = new QueryFoldingEngine();
  const query = {
    id: "q_db",
    name: "DB Query",
    source: { type: "database", connection: {}, query: "SELECT * FROM sales" },
    steps: [
      { id: "s1", name: "Select", operation: { type: "selectColumns", columns: ["Region", "Sales"] } },
      {
        id: "s2",
        name: "Filter",
        operation: { type: "filterRows", predicate: { type: "comparison", column: "Region", operator: "equals", value: "East" } },
      },
    ],
  };

  const explained = folding.explain(query, { dialect: "postgres" });
  assert.equal(explained.plan.type, "sql");
  assert.deepEqual(
    explained.steps.map((s) => s.status),
    ["folded", "folded"],
  );
  assert.ok(explained.steps[0].sqlFragment);
  assert.equal(explained.steps.at(-1)?.sqlFragment, explained.plan.sql);
});

test("explain: hybrid query marks folded prefix + local suffix with a reason", () => {
  const folding = new QueryFoldingEngine();
  const query = {
    id: "q_hybrid",
    name: "Hybrid",
    source: { type: "database", connection: {}, query: "SELECT * FROM sales" },
    steps: [
      {
        id: "s1",
        name: "Filter",
        operation: { type: "filterRows", predicate: { type: "comparison", column: "Sales", operator: "greaterThan", value: 0 } },
      },
      { id: "s2", name: "Fill Down", operation: { type: "fillDown", columns: ["Region"] } },
    ],
  };

  const explained = folding.explain(query, { dialect: "postgres" });
  assert.equal(explained.plan.type, "hybrid");
  assert.equal(explained.steps[0].status, "folded");
  assert.equal(explained.steps[1].status, "local");
  assert.equal(explained.steps[1].reason, "unsupported_op");
});

test("explain: missing dialect marks all steps as local with missing_dialect", () => {
  const folding = new QueryFoldingEngine();
  const query = {
    id: "q_missing_dialect",
    name: "Missing dialect",
    source: { type: "database", connection: {}, query: "SELECT * FROM sales" },
    steps: [{ id: "s1", name: "Filter", operation: { type: "filterRows", predicate: { type: "comparison", column: "Region", operator: "equals", value: "East" } } }],
  };

  const explained = folding.explain(query);
  assert.equal(explained.plan.type, "local");
  assert.deepEqual(explained.steps, [
    { stepId: "s1", opType: "filterRows", status: "local", reason: "missing_dialect" },
  ]);
});

test("explain: merge across different connections marks merge step local with different_connection", () => {
  const folding = new QueryFoldingEngine();

  const right = {
    id: "q_right",
    name: "Targets",
    source: { type: "database", connection: { name: "db2" }, query: "SELECT * FROM targets" },
    steps: [{ id: "r1", name: "Select", operation: { type: "selectColumns", columns: ["Id", "Target"] } }],
  };

  const left = {
    id: "q_left",
    name: "Sales",
    source: { type: "database", connection: { name: "db1" }, query: "SELECT * FROM sales" },
    steps: [
      { id: "l1", name: "Select", operation: { type: "selectColumns", columns: ["Id", "Sales"] } },
      { id: "l2", name: "Merge", operation: { type: "merge", rightQuery: "q_right", joinType: "left", leftKey: "Id", rightKey: "Id" } },
    ],
  };

  const explained = folding.explain(left, { dialect: "postgres", queries: { q_right: right } });
  assert.equal(explained.plan.type, "hybrid");
  assert.equal(explained.steps[0].status, "folded");
  assert.equal(explained.steps[1].status, "local");
  assert.equal(explained.steps[1].reason, "different_connection");
});

