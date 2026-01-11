import assert from "node:assert/strict";
import test from "node:test";

import { QueryFoldingEngine } from "../../src/folding/sql.js";
import { getSqlSourceId } from "../../src/privacy/sourceId.js";

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

test("explain: unsupported table ops stop folding with unsupported_op", () => {
  const folding = new QueryFoldingEngine();
  const query = {
    id: "q_explain_unsupported",
    name: "Explain unsupported",
    source: { type: "database", connection: {}, query: "SELECT * FROM sales", dialect: "postgres" },
    steps: [{ id: "s1", name: "Promote", operation: { type: "promoteHeaders" } }],
  };

  const result = folding.explain(query, { dialect: "postgres" });
  assert.equal(result.plan.type, "hybrid");
  assert.equal(result.steps[0].status, "local");
  assert.equal(result.steps[0].reason, "unsupported_op");
});

test("explain: merge blocked by privacy levels marks step local with privacy_firewall", () => {
  const folding = new QueryFoldingEngine();
  const sharedConnection = { id: "db1" };

  const right = {
    id: "q_right",
    name: "Targets",
    source: { type: "database", connection: sharedConnection, query: "SELECT * FROM targets", dialect: "postgres" },
    steps: [{ id: "r1", name: "Select", operation: { type: "selectColumns", columns: ["Id", "Target"] } }],
  };

  const left = {
    id: "q_left",
    name: "Sales",
    // Explicit connectionId that does *not* match `connection.id` so the folding
    // engine's stable source ids differ even though the connection handle is the same.
    source: { type: "database", connectionId: "connA", connection: sharedConnection, query: "SELECT * FROM sales", dialect: "postgres" },
    steps: [
      { id: "l1", name: "Select", operation: { type: "selectColumns", columns: ["Id"] } },
      {
        id: "l2",
        name: "Merge",
        operation: { type: "merge", rightQuery: "q_right", joinType: "left", leftKeys: ["Id"], rightKeys: ["Id"], joinMode: "flat" },
      },
    ],
  };

  const explained = folding.explain(left, {
    dialect: "postgres",
    queries: { q_right: right },
    privacyMode: "enforce",
    privacyLevelsBySourceId: {
      [getSqlSourceId("connA")]: "organizational",
      [getSqlSourceId(sharedConnection)]: "public",
    },
  });

  assert.equal(explained.plan.type, "hybrid");
  assert.equal(explained.steps[0].status, "folded");
  assert.equal(explained.steps[1].status, "local");
  assert.equal(explained.steps[1].reason, "privacy_firewall");
  assert.ok(Array.isArray(explained.plan.diagnostics) && explained.plan.diagnostics.length > 0);
});

test("explain: nested join is not folded (unsupported_join_mode)", () => {
  const folding = new QueryFoldingEngine();
  const connection = { id: "db1" };

  const right = {
    id: "q_right_nested",
    name: "Targets",
    source: { type: "database", connection, query: "SELECT * FROM targets", dialect: "postgres" },
    steps: [{ id: "r1", name: "Select", operation: { type: "selectColumns", columns: ["Id", "Target"] } }],
  };

  const left = {
    id: "q_left_nested",
    name: "Sales",
    source: { type: "database", connection, query: "SELECT * FROM sales", dialect: "postgres" },
    steps: [
      { id: "l1", name: "Select", operation: { type: "selectColumns", columns: ["Id"] } },
      {
        id: "l2",
        name: "Nested Join",
        operation: {
          type: "merge",
          rightQuery: "q_right_nested",
          joinType: "left",
          leftKeys: ["Id"],
          rightKeys: ["Id"],
          joinMode: "nested",
          newColumnName: "Matches",
        },
      },
    ],
  };

  const explained = folding.explain(left, { dialect: "postgres", queries: { q_right_nested: right } });
  assert.equal(explained.plan.type, "hybrid");
  assert.deepEqual(
    explained.steps.map((s) => [s.opType, s.status]),
    [
      ["selectColumns", "folded"],
      ["merge", "local"],
    ],
  );
  assert.equal(explained.steps[1].reason, "unsupported_join_mode");
});

test("explain: nested join + expand folds as a flattened join", () => {
  const folding = new QueryFoldingEngine();
  const connection = { id: "db1" };

  const right = {
    id: "q_right_nested_expand",
    name: "Targets",
    source: { type: "database", connection, query: "SELECT * FROM targets", dialect: "postgres" },
    steps: [{ id: "r1", name: "Select", operation: { type: "selectColumns", columns: ["Id", "Target"] } }],
  };

  const left = {
    id: "q_left_nested_expand",
    name: "Sales",
    source: { type: "database", connection, query: "SELECT * FROM sales", dialect: "postgres" },
    steps: [
      { id: "l1", name: "Select", operation: { type: "selectColumns", columns: ["Id", "Target", "Sales"] } },
      {
        id: "l2",
        name: "Nested Join",
        operation: {
          type: "merge",
          rightQuery: "q_right_nested_expand",
          joinType: "left",
          leftKeys: ["Id"],
          rightKeys: ["Id"],
          joinMode: "nested",
          newColumnName: "Matches",
        },
      },
      { id: "l3", name: "Expand", operation: { type: "expandTableColumn", column: "Matches", columns: ["Target"] } },
    ],
  };

  const explained = folding.explain(left, { dialect: "postgres", queries: { q_right_nested_expand: right } });
  assert.equal(explained.plan.type, "sql");
  assert.deepEqual(
    explained.steps.map((s) => [s.opType, s.status]),
    [
      ["selectColumns", "folded"],
      ["merge", "folded"],
      ["expandTableColumn", "folded"],
    ],
  );
  assert.match(explained.plan.sql, /\bJOIN\b/);
  assert.match(explained.plan.sql, /"Target\.1"/);
});

test("explain: merge with comparer is not folded (unsupported_comparer)", () => {
  const folding = new QueryFoldingEngine();
  const connection = { id: "db1" };

  const right = {
    id: "q_right_comparer",
    name: "Scores",
    source: { type: "database", connection, query: "SELECT * FROM scores", dialect: "postgres", columns: ["Name", "Score"] },
    steps: [],
  };

  const left = {
    id: "q_left_comparer",
    name: "People",
    source: { type: "database", connection, query: "SELECT * FROM people", dialect: "postgres", columns: ["Id", "Name"] },
    steps: [
      { id: "l1", name: "Select", operation: { type: "selectColumns", columns: ["Id", "Name"] } },
      {
        id: "l2",
        name: "Merge (ignore case)",
        operation: {
          type: "merge",
          rightQuery: "q_right_comparer",
          joinType: "inner",
          leftKeys: ["Name"],
          rightKeys: ["Name"],
          joinMode: "flat",
          comparer: { comparer: "ordinalIgnoreCase", caseSensitive: false },
        },
      },
    ],
  };

  const explained = folding.explain(left, { dialect: "postgres", queries: { q_right_comparer: right } });
  assert.equal(explained.plan.type, "hybrid");
  assert.equal(explained.steps[0].status, "folded");
  assert.equal(explained.steps[1].status, "local");
  assert.equal(explained.steps[1].reason, "unsupported_comparer");
});
