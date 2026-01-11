import assert from "node:assert/strict";
import test from "node:test";

import { QueryFoldingEngine } from "../../src/folding/sql.js";

test("compile: folds selectColumns/filterRows/groupBy into a parameterized SQL plan", () => {
  const folding = new QueryFoldingEngine();
  const query = {
    id: "q_db",
    name: "DB Query",
    source: { type: "database", connection: {}, query: "SELECT * FROM sales" },
    steps: [
      {
        id: "s1",
        name: "Select",
        operation: { type: "selectColumns", columns: ["Region", "Sales"] },
      },
      {
        id: "s2",
        name: "Filter",
        operation: { type: "filterRows", predicate: { type: "comparison", column: "Region", operator: "equals", value: "East" } },
      },
      {
        id: "s3",
        name: "Group",
        operation: { type: "groupBy", groupColumns: ["Region"], aggregations: [{ column: "Sales", op: "sum", as: "Total Sales" }] },
      },
    ],
  };

  const plan = folding.compile(query);
  assert.deepEqual(plan, {
    type: "sql",
    sql: 'SELECT t."Region", COALESCE(SUM(CASE WHEN TRIM(CAST(t."Sales" AS TEXT)) = \'\' THEN NULL WHEN TRIM(CAST(t."Sales" AS TEXT)) ~ \'^[+-]?([0-9]+([.][0-9]*)?|[.][0-9]+)([eE][+-]?[0-9]+)?$\' THEN (CASE WHEN isfinite(CAST(TRIM(CAST(t."Sales" AS TEXT)) AS DOUBLE PRECISION)) THEN CAST(TRIM(CAST(t."Sales" AS TEXT)) AS DOUBLE PRECISION) ELSE NULL END) ELSE NULL END), 0) AS "Total Sales" FROM (SELECT * FROM (SELECT t."Region", t."Sales" FROM (SELECT * FROM sales) AS t) AS t WHERE (t."Region" = ?)) AS t GROUP BY t."Region"',
    params: ["East"],
  });
});

test("compile: selectColumns with duplicate names breaks folding (SQL cannot return zero/duplicate columns safely)", () => {
  const folding = new QueryFoldingEngine();
  const query = {
    id: "q_select_dupe",
    name: "Select dupe",
    source: { type: "database", connection: {}, query: "SELECT * FROM sales" },
    steps: [{ id: "s1", name: "Select", operation: { type: "selectColumns", columns: ["Region", "Region"] } }],
  };

  const plan = folding.compile(query);
  assert.equal(plan.type, "hybrid");
  assert.equal(plan.sql, "SELECT * FROM sales");
  assert.deepEqual(plan.localSteps.map((s) => s.operation.type), ["selectColumns"]);
});

test("compile: groupBy with no columns + no aggregations breaks folding", () => {
  const folding = new QueryFoldingEngine();
  const query = {
    id: "q_group_empty",
    name: "Group empty",
    source: { type: "database", connection: {}, query: "SELECT * FROM sales" },
    steps: [{ id: "s1", name: "Group", operation: { type: "groupBy", groupColumns: [], aggregations: [] } }],
  };

  const plan = folding.compile(query);
  assert.equal(plan.type, "hybrid");
  assert.equal(plan.sql, "SELECT * FROM sales");
  assert.deepEqual(plan.localSteps.map((s) => s.operation.type), ["groupBy"]);
});

test("compile: LIKE predicates escape wildcard characters + use ESCAPE clause", () => {
  const folding = new QueryFoldingEngine();
  const query = {
    id: "q_like",
    name: "Like",
    source: { type: "database", connection: {}, query: "SELECT * FROM items" },
    steps: [
      {
        id: "s1",
        name: "Filter",
        operation: { type: "filterRows", predicate: { type: "comparison", column: "Name", operator: "contains", value: "50%_!\\test" } },
      },
    ],
  };

  const plan = folding.compile(query);
  assert.deepEqual(plan, {
    type: "sql",
    sql: 'SELECT * FROM (SELECT * FROM items) AS t WHERE (LOWER(COALESCE(CAST(t."Name" AS TEXT), \'\')) LIKE LOWER(?) ESCAPE \'!\')',
    params: ["%50!%!_!!\\test%"],
  });
});

test("compile: folds renameColumn when output columns are known", () => {
  const folding = new QueryFoldingEngine();
  const query = {
    id: "q_rename",
    name: "Rename",
    source: { type: "database", connection: {}, query: "SELECT * FROM sales" },
    steps: [
      { id: "s1", name: "Select", operation: { type: "selectColumns", columns: ["Region", "Sales"] } },
      { id: "s2", name: "Rename", operation: { type: "renameColumn", oldName: "Sales", newName: "Amount" } },
    ],
  };

  const plan = folding.compile(query);
  assert.deepEqual(plan, {
    type: "sql",
    sql: 'SELECT t."Region", t."Sales" AS "Amount" FROM (SELECT t."Region", t."Sales" FROM (SELECT * FROM sales) AS t) AS t',
    params: [],
  });
});

test("compile: renameColumn to an existing column breaks folding (matches local error semantics)", () => {
  const folding = new QueryFoldingEngine();
  const query = {
    id: "q_rename_break",
    name: "Rename Break",
    source: { type: "database", connection: {}, query: "SELECT * FROM sales" },
    steps: [
      { id: "s1", name: "Select", operation: { type: "selectColumns", columns: ["Region", "Sales"] } },
      { id: "s2", name: "Rename", operation: { type: "renameColumn", oldName: "Sales", newName: "Region" } },
    ],
  };

  const plan = folding.compile(query);
  assert.equal(plan.type, "hybrid");
  assert.equal(plan.sql, 'SELECT t."Region", t."Sales" FROM (SELECT * FROM sales) AS t');
  assert.deepEqual(plan.localSteps.map((s) => s.operation.type), ["renameColumn"]);
});

test("compile: folds changeType via CAST when output columns are known", () => {
  const folding = new QueryFoldingEngine();
  const query = {
    id: "q_cast",
    name: "Cast",
    source: { type: "database", connection: {}, query: "SELECT * FROM raw", columns: ["Value"] },
    steps: [{ id: "s1", name: "Type", operation: { type: "changeType", column: "Value", newType: "number" } }],
  };

  const plan = folding.compile(query);
  assert.deepEqual(plan, {
    type: "sql",
    sql: 'SELECT CASE WHEN TRIM(CAST(t."Value" AS TEXT)) = \'\' THEN NULL WHEN TRIM(CAST(t."Value" AS TEXT)) ~ \'^[+-]?([0-9]+([.][0-9]*)?|[.][0-9]+)([eE][+-]?[0-9]+)?$\' THEN (CASE WHEN isfinite(CAST(TRIM(CAST(t."Value" AS TEXT)) AS DOUBLE PRECISION)) THEN CAST(TRIM(CAST(t."Value" AS TEXT)) AS DOUBLE PRECISION) ELSE NULL END) ELSE NULL END AS "Value" FROM (SELECT * FROM raw) AS t',
    params: [],
  });
});

test("compile: changeType without a known projection breaks folding into a hybrid plan", () => {
  const folding = new QueryFoldingEngine();
  const query = {
    id: "q_cast_break",
    name: "Cast",
    source: { type: "database", connection: {}, query: "SELECT * FROM raw" },
    steps: [{ id: "s1", name: "Type", operation: { type: "changeType", column: "Value", newType: "number" } }],
  };

  const plan = folding.compile(query);
  assert.equal(plan.type, "hybrid");
  assert.equal(plan.sql, "SELECT * FROM raw");
  assert.deepEqual(plan.params, []);
  assert.deepEqual(plan.localSteps.map((s) => s.operation.type), ["changeType"]);
});

test("compile: folds addColumn for a safe subset of formula expressions", () => {
  const folding = new QueryFoldingEngine();
  const query = {
    id: "q_add",
    name: "Add",
    source: { type: "database", connection: {}, query: "SELECT * FROM sales" },
    steps: [
      { id: "s1", name: "Select", operation: { type: "selectColumns", columns: ["Sales"] } },
      { id: "s2", name: "Add", operation: { type: "addColumn", name: "Double", formula: "=[Sales] * 2" } },
    ],
  };

  const plan = folding.compile(query);
  assert.deepEqual(plan, {
    type: "sql",
    sql: 'SELECT t.*, (t."Sales" * ?) AS "Double" FROM (SELECT t."Sales" FROM (SELECT * FROM sales) AS t) AS t',
    params: [2],
  });
});

test("compile: folds addColumn with exponent number literals", () => {
  const folding = new QueryFoldingEngine();
  const query = {
    id: "q_add_exponent",
    name: "Add exponent",
    source: { type: "database", connection: {}, query: "SELECT * FROM sales" },
    steps: [
      { id: "s1", name: "Select", operation: { type: "selectColumns", columns: ["Sales"] } },
      { id: "s2", name: "Add", operation: { type: "addColumn", name: "Scaled", formula: "=[Sales] * 1e3" } },
    ],
  };

  const plan = folding.compile(query);
  assert.deepEqual(plan, {
    type: "sql",
    sql: 'SELECT t.*, (t."Sales" * ?) AS "Scaled" FROM (SELECT t."Sales" FROM (SELECT * FROM sales) AS t) AS t',
    params: [1000],
  });
});

test("compile: folds addColumn ternary expressions", () => {
  const folding = new QueryFoldingEngine();
  const query = {
    id: "q_add_ternary",
    name: "Add ternary",
    source: { type: "database", connection: {}, query: "SELECT * FROM sales" },
    steps: [
      { id: "s1", name: "Select", operation: { type: "selectColumns", columns: ["Sales"] } },
      { id: "s2", name: "Add", operation: { type: "addColumn", name: "Bucket", formula: '=[Sales] > 100 ? "big" : "small"' } },
    ],
  };

  const plan = folding.compile(query);
  assert.deepEqual(plan, {
    type: "sql",
    sql: 'SELECT t.*, (CASE WHEN (t."Sales" > ?) THEN ? ELSE ? END) AS "Bucket" FROM (SELECT t."Sales" FROM (SELECT * FROM sales) AS t) AS t',
    params: [100, "big", "small"],
  });
});

test("compile: folds addColumn null equality to IS NULL (SQL semantics match local)", () => {
  const folding = new QueryFoldingEngine();
  const query = {
    id: "q_add_null_eq",
    name: "Add null eq",
    source: { type: "database", connection: {}, query: "SELECT * FROM sales" },
    steps: [
      { id: "s1", name: "Select", operation: { type: "selectColumns", columns: ["Sales"] } },
      { id: "s2", name: "Add", operation: { type: "addColumn", name: "IsNull", formula: "=[Sales] == null ? 1 : 0" } },
    ],
  };

  const plan = folding.compile(query);
  assert.deepEqual(plan, {
    type: "sql",
    sql: 'SELECT t.*, (CASE WHEN (t."Sales" IS NULL) THEN ? ELSE ? END) AS "IsNull" FROM (SELECT t."Sales" FROM (SELECT * FROM sales) AS t) AS t',
    params: [1, 0],
  });
});

test("compile: folds addColumn string literals with escapes", () => {
  const folding = new QueryFoldingEngine();
  const payload = 'a"b\\c\n';
  const query = {
    id: "q_add_string_escape",
    name: "Add string escape",
    source: { type: "database", connection: {}, query: "SELECT * FROM sales" },
    steps: [{ id: "s1", name: "Add", operation: { type: "addColumn", name: "Text", formula: JSON.stringify(payload) } }],
  };

  const plan = folding.compile(query);
  assert.equal(plan.type, "sql");
  assert.ok(plan.sql.includes("?"));
  assert.deepEqual(plan.params, [payload]);
});

test("compile: folds addColumn with date() literal", () => {
  const folding = new QueryFoldingEngine();
  const query = {
    id: "q_add_date",
    name: "Add date",
    source: { type: "database", connection: {}, query: "SELECT * FROM sales" },
    steps: [{ id: "s1", name: "Add", operation: { type: "addColumn", name: "Day", formula: 'date("2020-01-01")' } }],
  };

  const plan = folding.compile(query);
  assert.deepEqual(plan, {
    type: "sql",
    sql: 'SELECT t.*, ? AS "Day" FROM (SELECT * FROM sales) AS t',
    params: ["2020-01-01T00:00:00.000Z"],
  });
});

test("compile: folds transformColumns identity casts when the formula parses to '_'", () => {
  const folding = new QueryFoldingEngine();
  const query = {
    id: "q_transform_identity",
    name: "Transform identity",
    source: { type: "database", connection: {}, query: "SELECT * FROM raw", columns: ["Value"] },
    steps: [
      {
        id: "s1",
        name: "Transform",
        operation: { type: "transformColumns", transforms: [{ column: "Value", formula: "=(( _ ))", newType: "string" }] },
      },
    ],
  };

  const plan = folding.compile(query);
  assert.deepEqual(plan, {
    type: "sql",
    sql: 'SELECT CAST(t."Value" AS TEXT) AS "Value" FROM (SELECT * FROM raw) AS t',
    params: [],
  });
});

test("compile: addColumn params come before nested query params (placeholder order)", () => {
  const folding = new QueryFoldingEngine();
  const query = {
    id: "q_add_param_order",
    name: "Add param order",
    source: { type: "database", connection: {}, query: "SELECT * FROM sales" },
    steps: [
      {
        id: "s1",
        name: "Filter",
        operation: { type: "filterRows", predicate: { type: "comparison", column: "Region", operator: "equals", value: "East" } },
      },
      { id: "s2", name: "Add", operation: { type: "addColumn", name: "Injected", formula: '"x"' } },
    ],
  };

  const plan = folding.compile(query);
  assert.deepEqual(plan, {
    type: "sql",
    sql: 'SELECT t.*, ? AS "Injected" FROM (SELECT * FROM (SELECT * FROM sales) AS t WHERE (t."Region" = ?)) AS t',
    params: ["x", "East"],
  });
});

test("compile: non-translatable addColumn formula breaks folding into a hybrid plan", () => {
  const folding = new QueryFoldingEngine();
  const query = {
    id: "q_add_break",
    name: "Add",
    source: { type: "database", connection: {}, query: "SELECT * FROM sales" },
    steps: [
      {
        id: "s1",
        name: "Filter",
        operation: { type: "filterRows", predicate: { type: "comparison", column: "Region", operator: "equals", value: "East" } },
      },
      { id: "s2", name: "Add", operation: { type: "addColumn", name: "Bad", formula: "=Math.abs([Sales])" } },
      { id: "s3", name: "Take", operation: { type: "take", count: 5 } },
    ],
  };

  const plan = folding.compile(query);
  assert.equal(plan.type, "hybrid");
  assert.deepEqual(plan.params, ["East"]);
  assert.equal(plan.sql, 'SELECT * FROM (SELECT * FROM sales) AS t WHERE (t."Region" = ?)');
  assert.deepEqual(plan.localSteps.map((s) => s.operation.type), ["addColumn", "take"]);
});

test("compile: folds merge (join) when both sides fully fold to SQL", () => {
  const folding = new QueryFoldingEngine();
  const connection = {};

  const right = {
    id: "q_right",
    name: "Targets",
    source: { type: "database", connection, query: "SELECT * FROM targets" },
    steps: [{ id: "s1", name: "Select", operation: { type: "selectColumns", columns: ["Id", "Target"] } }],
  };

  const left = {
    id: "q_left",
    name: "Sales",
    source: { type: "database", connection, query: "SELECT * FROM sales" },
    steps: [
      { id: "s1", name: "Select", operation: { type: "selectColumns", columns: ["Id", "Region", "Sales"] } },
      { id: "s2", name: "Merge", operation: { type: "merge", rightQuery: "q_right", joinType: "left", leftKey: "Id", rightKey: "Id" } },
    ],
  };

  const plan = folding.compile(left, { queries: { q_right: right } });
  assert.deepEqual(plan, {
    type: "sql",
    sql: 'SELECT l."Id" AS "Id", l."Region" AS "Region", l."Sales" AS "Sales", r."Target" AS "Target" FROM (SELECT t."Id", t."Region", t."Sales" FROM (SELECT * FROM sales) AS t) AS l LEFT JOIN (SELECT t."Id", t."Target" FROM (SELECT * FROM targets) AS t) AS r ON l."Id" IS NOT DISTINCT FROM r."Id"',
    params: [],
  });
});

test("compile: folds merge when connections are deep-equal but not referentially equal", () => {
  const folding = new QueryFoldingEngine();
  const leftConn = { host: "localhost", database: "db1" };
  const rightConn = { host: "localhost", database: "db1" };

  const right = {
    id: "q_right",
    name: "Targets",
    source: { type: "database", connection: rightConn, query: "SELECT * FROM targets" },
    steps: [{ id: "s1", name: "Select", operation: { type: "selectColumns", columns: ["Id", "Target"] } }],
  };

  const left = {
    id: "q_left",
    name: "Sales",
    source: { type: "database", connection: leftConn, query: "SELECT * FROM sales" },
    steps: [
      { id: "s1", name: "Select", operation: { type: "selectColumns", columns: ["Id", "Region", "Sales"] } },
      { id: "s2", name: "Merge", operation: { type: "merge", rightQuery: "q_right", joinType: "left", leftKey: "Id", rightKey: "Id" } },
    ],
  };

  const plan = folding.compile(left, { queries: { q_right: right }, getConnectionIdentity: (connection) => connection });
  assert.equal(plan.type, "sql");
  assert.match(plan.sql, /\bJOIN\b/);
});

test("compile: folds merge when both connections share an id without getConnectionIdentity", () => {
  const folding = new QueryFoldingEngine();
  const leftConn = { id: "db1", host: "localhost" };
  const rightConn = { id: "db1", host: "localhost" };

  const right = {
    id: "q_right",
    name: "Targets",
    source: { type: "database", connection: rightConn, query: "SELECT * FROM targets" },
    steps: [{ id: "s1", name: "Select", operation: { type: "selectColumns", columns: ["Id", "Target"] } }],
  };

  const left = {
    id: "q_left",
    name: "Sales",
    source: { type: "database", connection: leftConn, query: "SELECT * FROM sales" },
    steps: [
      { id: "s1", name: "Select", operation: { type: "selectColumns", columns: ["Id", "Region", "Sales"] } },
      { id: "s2", name: "Merge", operation: { type: "merge", rightQuery: "q_right", joinType: "left", leftKey: "Id", rightKey: "Id" } },
    ],
  };

  const plan = folding.compile(left, { queries: { q_right: right } });
  assert.equal(plan.type, "sql");
  assert.match(plan.sql, /\bJOIN\b/);
});

test("compile: merge across different database connections breaks folding", () => {
  const folding = new QueryFoldingEngine();

  const right = {
    id: "q_right",
    name: "Targets",
    source: { type: "database", connection: { name: "db2" }, query: "SELECT * FROM targets" },
    steps: [{ id: "s1", name: "Select", operation: { type: "selectColumns", columns: ["Id", "Target"] } }],
  };

  const left = {
    id: "q_left",
    name: "Sales",
    source: { type: "database", connection: { name: "db1" }, query: "SELECT * FROM sales" },
    steps: [
      { id: "s1", name: "Select", operation: { type: "selectColumns", columns: ["Id", "Sales"] } },
      { id: "s2", name: "Merge", operation: { type: "merge", rightQuery: "q_right", joinType: "left", leftKey: "Id", rightKey: "Id" } },
    ],
  };

  const plan = folding.compile(left, { queries: { q_right: right } });
  assert.equal(plan.type, "hybrid");
  assert.deepEqual(plan.params, []);
  assert.equal(plan.sql, 'SELECT t."Id", t."Sales" FROM (SELECT * FROM sales) AS t');
  assert.deepEqual(plan.localSteps.map((s) => s.operation.type), ["merge"]);
});

test("compile: folds append (UNION ALL) when schemas are compatible", () => {
  const folding = new QueryFoldingEngine();
  const connection = {};

  const other = {
    id: "q_other",
    name: "Other",
    source: { type: "database", connection, query: "SELECT * FROM b" },
    steps: [{ id: "s1", name: "Select", operation: { type: "selectColumns", columns: ["Value", "Id"] } }],
  };

  const base = {
    id: "q_base",
    name: "Base",
    source: { type: "database", connection, query: "SELECT * FROM a" },
    steps: [
      { id: "s1", name: "Select", operation: { type: "selectColumns", columns: ["Id", "Value"] } },
      { id: "s2", name: "Append", operation: { type: "append", queries: ["q_other"] } },
    ],
  };

  const plan = folding.compile(base, { queries: { q_other: other } });
  assert.deepEqual(plan, {
    type: "sql",
    sql: '(SELECT t."Id", t."Value" FROM (SELECT t."Id", t."Value" FROM (SELECT * FROM a) AS t) AS t) UNION ALL (SELECT t."Id", t."Value" FROM (SELECT t."Value", t."Id" FROM (SELECT * FROM b) AS t) AS t)',
    params: [],
  });
});

test("compile: folds append when connections are deep-equal but not referentially equal", () => {
  const folding = new QueryFoldingEngine();
  const baseConn = { host: "localhost", database: "db1" };
  const otherConn = { host: "localhost", database: "db1" };

  const other = {
    id: "q_other",
    name: "Other",
    source: { type: "database", connection: otherConn, query: "SELECT * FROM b" },
    steps: [{ id: "s1", name: "Select", operation: { type: "selectColumns", columns: ["Value", "Id"] } }],
  };

  const base = {
    id: "q_base",
    name: "Base",
    source: { type: "database", connection: baseConn, query: "SELECT * FROM a" },
    steps: [
      { id: "s1", name: "Select", operation: { type: "selectColumns", columns: ["Id", "Value"] } },
      { id: "s2", name: "Append", operation: { type: "append", queries: ["q_other"] } },
    ],
  };

  const plan = folding.compile(base, { queries: { q_other: other }, getConnectionIdentity: (connection) => connection });
  assert.equal(plan.type, "sql");
  assert.match(plan.sql, /\bUNION ALL\b/);
});

test("compile: folds append when both connections share an id without getConnectionIdentity", () => {
  const folding = new QueryFoldingEngine();
  const baseConn = { id: "db1", host: "localhost" };
  const otherConn = { id: "db1", host: "localhost" };

  const other = {
    id: "q_other",
    name: "Other",
    source: { type: "database", connection: otherConn, query: "SELECT * FROM b" },
    steps: [{ id: "s1", name: "Select", operation: { type: "selectColumns", columns: ["Value", "Id"] } }],
  };

  const base = {
    id: "q_base",
    name: "Base",
    source: { type: "database", connection: baseConn, query: "SELECT * FROM a" },
    steps: [
      { id: "s1", name: "Select", operation: { type: "selectColumns", columns: ["Id", "Value"] } },
      { id: "s2", name: "Append", operation: { type: "append", queries: ["q_other"] } },
    ],
  };

  const plan = folding.compile(base, { queries: { q_other: other } });
  assert.equal(plan.type, "sql");
  assert.match(plan.sql, /\bUNION ALL\b/);
});

test("compile: append across different database connections breaks folding", () => {
  const folding = new QueryFoldingEngine();

  const other = {
    id: "q_other",
    name: "Other",
    source: { type: "database", connection: { name: "db2" }, query: "SELECT * FROM b" },
    steps: [{ id: "s1", name: "Select", operation: { type: "selectColumns", columns: ["Id", "Value"] } }],
  };

  const base = {
    id: "q_base",
    name: "Base",
    source: { type: "database", connection: { name: "db1" }, query: "SELECT * FROM a" },
    steps: [
      { id: "s1", name: "Select", operation: { type: "selectColumns", columns: ["Id", "Value"] } },
      { id: "s2", name: "Append", operation: { type: "append", queries: ["q_other"] } },
    ],
  };

  const plan = folding.compile(base, { queries: { q_other: other } });
  assert.equal(plan.type, "hybrid");
  assert.deepEqual(plan.localSteps.map((s) => s.operation.type), ["append"]);
});

test("compile: folds take (LIMIT) into SQL", () => {
  const folding = new QueryFoldingEngine();
  const query = {
    id: "q_take",
    name: "Take",
    source: { type: "database", connection: {}, query: "SELECT * FROM sales" },
    steps: [{ id: "s1", name: "Take", operation: { type: "take", count: 10 } }],
  };

  const plan = folding.compile(query);
  assert.deepEqual(plan, { type: "sql", sql: "SELECT * FROM (SELECT * FROM sales) AS t LIMIT ?", params: [10] });
});

test("compile: folds removeColumns when projection is known", () => {
  const folding = new QueryFoldingEngine();
  const query = {
    id: "q_remove",
    name: "Remove",
    source: { type: "database", connection: {}, query: "SELECT * FROM sales" },
    steps: [
      { id: "s1", name: "Select", operation: { type: "selectColumns", columns: ["Region", "Product", "Sales"] } },
      { id: "s2", name: "Remove", operation: { type: "removeColumns", columns: ["Product"] } },
    ],
  };

  const plan = folding.compile(query);
  assert.deepEqual(plan, {
    type: "sql",
    sql: 'SELECT t."Region", t."Sales" FROM (SELECT t."Region", t."Product", t."Sales" FROM (SELECT * FROM sales) AS t) AS t',
    params: [],
  });
});

test("compile: dialect-specific quoting + NULL ordering (MySQL)", () => {
  const folding = new QueryFoldingEngine();
  const query = {
    id: "q_mysql_sort",
    name: "Sort",
    source: { type: "database", connection: {}, query: "SELECT * FROM sales" },
    steps: [{ id: "s1", name: "Sort", operation: { type: "sortRows", sortBy: [{ column: "Sales", direction: "descending", nulls: "first" }] } }],
  };

  const plan = folding.compile(query, { dialect: "mysql" });
  assert.deepEqual(plan, {
    type: "sql",
    sql: "SELECT * FROM (SELECT * FROM sales) AS t ORDER BY (t.`Sales` IS NULL) DESC, t.`Sales` DESC",
    params: [],
  });
});

test("compile: dialect-specific Date parameter formatting (MySQL)", () => {
  const folding = new QueryFoldingEngine();
  const query = {
    id: "q_mysql_date",
    name: "Filter Date",
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

  const plan = folding.compile(query, { dialect: "mysql" });
  assert.equal(plan.type, "sql");
  assert.equal(plan.sql, "SELECT * FROM (SELECT * FROM events) AS t WHERE (t.`CreatedAt` = ?)");
  assert.deepEqual(plan.params, ["2020-01-02 03:04:05"]);
});
