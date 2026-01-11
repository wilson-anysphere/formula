import assert from "node:assert/strict";
import test from "node:test";

import { QueryEngine } from "../../src/engine.js";
import { SqlConnector } from "../../src/connectors/sql.js";
import { QueryFoldingEngine } from "../../src/folding/sql.js";
import { getSqlDialect } from "../../src/folding/dialect.js";
import { normalizeSqlServerPlaceholders } from "../../src/folding/placeholders.js";
import { DataTable } from "../../src/table.js";

test("sqlserver: quotes identifiers using brackets", () => {
  const dialect = getSqlDialect("sqlserver");
  assert.equal(dialect.quoteIdentifier("Region"), "[Region]");
  assert.equal(dialect.quoteIdentifier("a]b"), "[a]]b]");
});

test("sqlserver: ORDER BY emulates NULLS FIRST/LAST", () => {
  const folding = new QueryFoldingEngine();
  const query = {
    id: "q_sqlserver_sort",
    name: "SQL Server Sort",
    source: { type: "database", connection: {}, query: "SELECT * FROM sales" },
    steps: [
      {
        id: "s1",
        name: "Sort",
        operation: { type: "sortRows", sortBy: [{ column: "Sales", direction: "descending", nulls: "first" }] },
      },
    ],
  };

  const plan = folding.compile(query, { dialect: "sqlserver" });
  assert.deepEqual(plan, {
    type: "sql",
    sql: "SELECT * FROM (SELECT * FROM sales) AS t ORDER BY (CASE WHEN t.[Sales] IS NULL THEN 1 ELSE 0 END) DESC, t.[Sales] DESC",
    params: [],
  });
});

test("sqlserver: sort followed by take folds to TOP with ORDER BY (no nested ORDER BY)", () => {
  const folding = new QueryFoldingEngine();
  const query = {
    id: "q_sqlserver_sort_take",
    name: "SQL Server Sort + Take",
    source: { type: "database", connection: {}, query: "SELECT * FROM sales" },
    steps: [
      {
        id: "s1",
        name: "Sort",
        operation: { type: "sortRows", sortBy: [{ column: "Sales", direction: "descending", nulls: "first" }] },
      },
      {
        id: "s2",
        name: "Take",
        operation: { type: "take", count: 5 },
      },
    ],
  };

  const plan = folding.compile(query, { dialect: "sqlserver" });
  assert.deepEqual(plan, {
    type: "sql",
    sql: "SELECT TOP (?) * FROM (SELECT * FROM sales) AS t ORDER BY (CASE WHEN t.[Sales] IS NULL THEN 1 ELSE 0 END) DESC, t.[Sales] DESC",
    params: [5],
  });
});

test("sqlserver: sort followed by renameColumn keeps valid ORDER BY", () => {
  const folding = new QueryFoldingEngine();
  const query = {
    id: "q_sqlserver_sort_rename",
    name: "SQL Server Sort + Rename",
    source: { type: "database", connection: {}, query: "SELECT * FROM sales", columns: ["Sales"] },
    steps: [
      {
        id: "s1",
        name: "Sort",
        operation: { type: "sortRows", sortBy: [{ column: "Sales", direction: "descending", nulls: "first" }] },
      },
      {
        id: "s2",
        name: "Rename",
        operation: { type: "renameColumn", oldName: "Sales", newName: "Amount" },
      },
    ],
  };

  const plan = folding.compile(query, { dialect: "sqlserver" });
  assert.deepEqual(plan, {
    type: "sql",
    sql: "SELECT * FROM (SELECT t.[Sales] AS [Amount] FROM (SELECT * FROM sales) AS t) AS t ORDER BY (CASE WHEN t.[Amount] IS NULL THEN 1 ELSE 0 END) DESC, t.[Amount] DESC",
    params: [],
  });
});

test("sqlserver: placeholder normalization converts ? -> @pN (ignores strings/comments)", () => {
  const sql = "SELECT [col?] AS q, '?' AS lit FROM t -- ? comment\nWHERE a = ? AND b = ? /* ? block */";
  assert.equal(
    normalizeSqlServerPlaceholders(sql, 2),
    "SELECT [col?] AS q, '?' AS lit FROM t -- ? comment\nWHERE a = @p1 AND b = @p2 /* ? block */",
  );
});

test("sqlserver: Date params are normalized to ISO strings without timezone suffix", () => {
  const folding = new QueryFoldingEngine();
  const query = {
    id: "q_sqlserver_date",
    name: "SQL Server Date",
    source: { type: "database", connection: {}, query: "SELECT * FROM events" },
    steps: [
      {
        id: "s1",
        name: "Filter",
        operation: {
          type: "filterRows",
          predicate: { type: "comparison", column: "CreatedAt", operator: "equals", value: new Date("2020-01-02T03:04:05.678Z") },
        },
      },
    ],
  };

  const plan = folding.compile(query, { dialect: "sqlserver" });
  assert.equal(plan.type, "sql");
  assert.equal(plan.sql, "SELECT * FROM (SELECT * FROM events) AS t WHERE (t.[CreatedAt] = ?)");
  assert.deepEqual(plan.params, ["2020-01-02T03:04:05.678"]);
});

test("sqlserver: addColumn supports boolean columns in ternary predicates", () => {
  const folding = new QueryFoldingEngine();
  const query = {
    id: "q_sqlserver_bool_ternary",
    name: "SQL Server bool ternary",
    source: { type: "database", connection: {}, query: "SELECT * FROM users", columns: ["IsActive"] },
    steps: [{ id: "s1", name: "Add", operation: { type: "addColumn", name: "ActiveFlag", formula: "=[IsActive] ? 1 : 0" } }],
  };

  const plan = folding.compile(query, { dialect: "sqlserver" });
  assert.deepEqual(plan, {
    type: "sql",
    sql: "SELECT t.*, (CASE WHEN (t.[IsActive] = 1) THEN CAST(? AS FLOAT) ELSE CAST(? AS FLOAT) END) AS [ActiveFlag] FROM (SELECT * FROM users) AS t",
    params: [1, 0],
  });
});

test("sqlserver: addColumn comparisons return BIT values (not predicates)", () => {
  const folding = new QueryFoldingEngine();
  const query = {
    id: "q_sqlserver_bool_value",
    name: "SQL Server bool value",
    source: { type: "database", connection: {}, query: "SELECT * FROM users", columns: ["Sales"] },
    steps: [{ id: "s1", name: "Add", operation: { type: "addColumn", name: "IsBig", formula: "=[Sales] > 100" } }],
  };

  const plan = folding.compile(query, { dialect: "sqlserver" });
  assert.deepEqual(plan, {
    type: "sql",
    sql: "SELECT t.*, (CASE WHEN (t.[Sales] > CAST(? AS FLOAT)) THEN CAST(1 AS BIT) ELSE CAST(0 AS BIT) END) AS [IsBig] FROM (SELECT * FROM users) AS t",
    params: [100],
  });
});

test("sqlserver: folding explain + QueryEngine execution normalize placeholders + preserve explain steps", async () => {
  /** @type {{ sql: string, params: unknown[] | undefined } | null} */
  let observed = null;

  const engine = new QueryEngine({
    connectors: {
      sql: new SqlConnector({
        querySql: async (_connection, sql, options) => {
          observed = { sql, params: options?.params };
          return DataTable.fromGrid(
            [
              ["Name", "Sales"],
              ["x", 100],
            ],
            { hasHeaders: true, inferTypes: true },
          );
        },
      }),
    },
  });

  const query = {
    id: "q_sqlserver_e2e",
    name: "SQL Server e2e",
    source: { type: "database", connection: { id: "db1" }, query: "SELECT * FROM items", dialect: "sqlserver" },
    steps: [
      {
        id: "s1",
        name: "Filter",
        operation: {
          type: "filterRows",
          predicate: {
            type: "and",
            predicates: [
              { type: "comparison", column: "Name", operator: "contains", value: "x" },
              { type: "comparison", column: "Sales", operator: "greaterThan", value: 0 },
            ],
          },
        },
      },
      {
        id: "s2",
        name: "Sort",
        operation: { type: "sortRows", sortBy: [{ column: "Sales", direction: "descending", nulls: "last" }] },
      },
    ],
  };

  const { meta } = await engine.executeQueryWithMeta(query, { queries: {} }, {});
  assert.ok(observed, "expected SQL connector to be invoked");
  assert.match(observed.sql, /\bORDER BY\b/);
  assert.match(observed.sql, /@p1/);
  assert.match(observed.sql, /@p2/);
  assert.deepEqual(observed.params, ["%x%", 0]);

  assert.ok(meta.folding, "expected folding metadata");
  assert.equal(meta.folding.dialect, "sqlserver");
  assert.equal(meta.folding.planType, "sql");
  // `meta.folding.sql` should reflect the executed (normalized) SQL.
  assert.match(meta.folding.sql, /@p1/);
  assert.match(meta.folding.sql, /@p2/);
  assert.deepEqual(
    meta.folding.steps.map((s) => s.status),
    ["folded", "folded"],
  );
});

test("sqlserver: compile/explain for a simple folded query stays in sync", () => {
  const folding = new QueryFoldingEngine();
  const query = {
    id: "q_sqlserver_explain",
    name: "SQL Server explain",
    source: { type: "database", connection: {}, query: "SELECT * FROM sales", dialect: "sqlserver" },
    steps: [{ id: "s1", name: "Take", operation: { type: "take", count: 5 } }],
  };

  const compiled = folding.compile(query, { dialect: "sqlserver" });
  const explained = folding.explain(query, { dialect: "sqlserver" });
  assert.equal(explained.plan.type, "sql");
  assert.deepEqual(explained.plan, compiled);
  assert.deepEqual(explained.steps.map((s) => s.status), ["folded"]);
});
