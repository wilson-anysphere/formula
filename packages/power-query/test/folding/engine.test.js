import assert from "node:assert/strict";
import test from "node:test";

import { QueryEngine } from "../../src/engine.js";
import { DataTable } from "../../src/table.js";
import { SqlConnector } from "../../src/connectors/sql.js";

test("QueryEngine: executes folded SQL when database dialect is provided", async () => {
  /** @type {{ sql: string, params: unknown[] | undefined } | null} */
  let observed = null;

  const engine = new QueryEngine({
    connectors: {
      sql: new SqlConnector({
        querySql: async (_connection, sql, options) => {
          observed = { sql, params: options?.params };
          // Simulate database applying the filter by returning only matching rows.
          return DataTable.fromGrid(
            [
              ["Region", "Sales"],
              ["East", 100],
            ],
            { hasHeaders: true, inferTypes: true },
          );
        },
      }),
    },
  });

  const query = {
    id: "q_fold",
    name: "Fold",
    source: { type: "database", connection: { id: "db1" }, query: "SELECT * FROM sales", dialect: "postgres" },
    steps: [
      {
        id: "s1",
        name: "Filter",
        operation: { type: "filterRows", predicate: { type: "comparison", column: "Region", operator: "equals", value: "East" } },
      },
    ],
  };

  const result = await engine.executeQuery(query, { queries: {} }, {});
  assert.ok(observed, "expected SQL connector to be invoked");
  assert.match(observed.sql, /WHERE/);
  assert.deepEqual(observed.params, ["East"]);
  assert.deepEqual(result.toGrid(), [
    ["Region", "Sales"],
    ["East", 100],
  ]);
});

test("QueryEngine: pushes ExecuteOptions.limit down when the plan fully folds to SQL", async () => {
  /** @type {{ sql: string, params: unknown[] | undefined } | null} */
  let observed = null;

  const engine = new QueryEngine({
    connectors: {
      sql: new SqlConnector({
        querySql: async (_connection, sql, options) => {
          observed = { sql, params: options?.params };
          // Pretend the database applied both the filter and limit.
          return DataTable.fromGrid(
            [
              ["Region", "Sales"],
              ["East", 100],
            ],
            { hasHeaders: true, inferTypes: true },
          );
        },
      }),
    },
  });

  const query = {
    id: "q_fold_limit",
    name: "Fold + Limit",
    source: { type: "database", connection: { id: "db1" }, query: "SELECT * FROM sales", dialect: "postgres" },
    steps: [
      {
        id: "s1",
        name: "Filter",
        operation: { type: "filterRows", predicate: { type: "comparison", column: "Region", operator: "equals", value: "East" } },
      },
    ],
  };

  await engine.executeQuery(query, { queries: {} }, { limit: 10 });
  assert.ok(observed, "expected SQL connector to be invoked");
  assert.match(observed.sql, /\bLIMIT\b/);
  assert.deepEqual(observed.params, ["East", 10]);
});

test("QueryEngine: without a dialect, executes steps locally (no folding)", async () => {
  /** @type {{ sql: string, params: unknown[] | undefined } | null} */
  let observed = null;

  const engine = new QueryEngine({
    connectors: {
      sql: new SqlConnector({
        querySql: async (_connection, sql, options) => {
          observed = { sql, params: options?.params };
          // Return unfiltered rows; local engine should filter them.
          return DataTable.fromGrid(
            [
              ["Region", "Sales"],
              ["East", 100],
              ["West", 200],
            ],
            { hasHeaders: true, inferTypes: true },
          );
        },
      }),
    },
  });

  const query = {
    id: "q_local",
    name: "Local",
    source: { type: "database", connection: { id: "db1" }, query: "SELECT * FROM sales" },
    steps: [
      {
        id: "s1",
        name: "Filter",
        operation: { type: "filterRows", predicate: { type: "comparison", column: "Region", operator: "equals", value: "East" } },
      },
    ],
  };

  const result = await engine.executeQuery(query, { queries: {} }, {});
  assert.ok(observed, "expected SQL connector to be invoked");
  assert.equal(observed.sql, "SELECT * FROM sales");
  assert.equal(observed.params, undefined);
  assert.deepEqual(result.toGrid(), [
    ["Region", "Sales"],
    ["East", 100],
  ]);
});

test("QueryEngine: executes hybrid folded SQL then runs remaining steps locally", async () => {
  /** @type {{ sql: string, params: unknown[] | undefined } | null} */
  let observed = null;

  const engine = new QueryEngine({
    connectors: {
      sql: new SqlConnector({
        querySql: async (_connection, sql, options) => {
          observed = { sql, params: options?.params };
          return DataTable.fromGrid(
            [
              ["Region", "Sales"],
              ["East", 100],
              [null, 150],
            ],
            { hasHeaders: true, inferTypes: true },
          );
        },
      }),
    },
  });

  const query = {
    id: "q_hybrid",
    name: "Hybrid",
    source: { type: "database", connection: { id: "db1" }, query: "SELECT * FROM sales", dialect: "postgres" },
    steps: [
      {
        id: "s1",
        name: "Filter",
        operation: { type: "filterRows", predicate: { type: "comparison", column: "Sales", operator: "greaterThan", value: 0 } },
      },
      {
        id: "s2",
        name: "Fill Down",
        operation: { type: "fillDown", columns: ["Region"] },
      },
    ],
  };

  const result = await engine.executeQuery(query, { queries: {} }, {});
  assert.ok(observed, "expected SQL connector to be invoked");
  assert.match(observed.sql, /WHERE/);
  assert.deepEqual(observed.params, [0]);

  // fillDown should run locally after the SQL query.
  assert.deepEqual(result.toGrid(), [
    ["Region", "Sales"],
    ["East", 100],
    ["East", 150],
  ]);
});

test("QueryEngine: folds merge into a single SQL query when both sides are foldable", async () => {
  /** @type {{ sql: string, params: unknown[] | undefined }[]} */
  const calls = [];
  const connection = {};

  const engine = new QueryEngine({
    connectors: {
      sql: new SqlConnector({
        querySql: async (_connection, sql, options) => {
          calls.push({ sql, params: options?.params });
          return DataTable.fromGrid(
            [
              ["Id", "Region", "Target"],
              [1, "East", 10],
            ],
            { hasHeaders: true, inferTypes: true },
          );
        },
      }),
    },
  });

  const right = {
    id: "q_right",
    name: "Targets",
    source: { type: "database", connection, query: "SELECT * FROM targets" },
    steps: [{ id: "r1", name: "Select", operation: { type: "selectColumns", columns: ["Id", "Target"] } }],
  };

  const left = {
    id: "q_left",
    name: "Sales",
    source: { type: "database", connection, query: "SELECT * FROM sales", dialect: "postgres" },
    steps: [
      { id: "l1", name: "Select", operation: { type: "selectColumns", columns: ["Id", "Region"] } },
      { id: "l2", name: "Merge", operation: { type: "merge", rightQuery: "q_right", joinType: "left", leftKey: "Id", rightKey: "Id" } },
    ],
  };

  const result = await engine.executeQuery(left, { queries: { q_right: right } }, {});
  assert.equal(calls.length, 1, "expected a single SQL roundtrip when merge folds");
  assert.match(calls[0].sql, /\bJOIN\b/);
  assert.deepEqual(result.toGrid(), [
    ["Id", "Region", "Target"],
    [1, "East", 10],
  ]);
});

test("QueryEngine: folds append into a single SQL query when schemas are compatible", async () => {
  /** @type {{ sql: string, params: unknown[] | undefined }[]} */
  const calls = [];
  const connection = {};

  const engine = new QueryEngine({
    connectors: {
      sql: new SqlConnector({
        querySql: async (_connection, sql, options) => {
          calls.push({ sql, params: options?.params });
          return DataTable.fromGrid(
            [
              ["Id", "Value"],
              [1, "a"],
              [2, "b"],
            ],
            { hasHeaders: true, inferTypes: true },
          );
        },
      }),
    },
  });

  const other = {
    id: "q_other",
    name: "Other",
    source: { type: "database", connection, query: "SELECT * FROM b" },
    steps: [{ id: "o1", name: "Select", operation: { type: "selectColumns", columns: ["Value", "Id"] } }],
  };

  const base = {
    id: "q_base",
    name: "Base",
    source: { type: "database", connection, query: "SELECT * FROM a", dialect: "postgres" },
    steps: [
      { id: "b1", name: "Select", operation: { type: "selectColumns", columns: ["Id", "Value"] } },
      { id: "b2", name: "Append", operation: { type: "append", queries: ["q_other"] } },
    ],
  };

  const result = await engine.executeQuery(base, { queries: { q_other: other } }, {});
  assert.equal(calls.length, 1, "expected a single SQL roundtrip when append folds");
  assert.match(calls[0].sql, /\bUNION ALL\b/);
  assert.deepEqual(result.toGrid(), [
    ["Id", "Value"],
    [1, "a"],
    [2, "b"],
  ]);
});
