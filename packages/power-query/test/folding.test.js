import assert from "node:assert/strict";
import test from "node:test";

import { QueryFoldingEngine } from "../src/folding/sql.js";

test("QueryFoldingEngine compiles a foldable prefix to SQL", () => {
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
  assert.equal(plan.type, "sql");
  assert.match(plan.sql, /SELECT\s+t\."Region"/);
  assert.match(plan.sql, /WHERE\s+\(t\."Region"\s+=\s+'East'\)/);
  assert.match(plan.sql, /GROUP BY\s+t\."Region"/);
  assert.match(plan.sql, /SUM\(t\."Sales"\)\s+AS\s+"Total Sales"/);
});

test("QueryFoldingEngine returns a hybrid plan when folding breaks", () => {
  const folding = new QueryFoldingEngine();
  const query = {
    id: "q_db_break",
    name: "DB Query",
    source: { type: "database", connection: {}, query: "SELECT * FROM sales" },
    steps: [
      {
        id: "s1",
        name: "Filter",
        operation: { type: "filterRows", predicate: { type: "comparison", column: "Region", operator: "equals", value: "East" } },
      },
      {
        id: "s_break",
        name: "FillDown",
        operation: { type: "fillDown", columns: ["Region"] },
      },
      {
        id: "s_after",
        name: "Sort",
        operation: { type: "sortRows", sortBy: [{ column: "Sales", direction: "descending" }] },
      },
    ],
  };

  const plan = folding.compile(query);
  assert.equal(plan.type, "hybrid");
  assert.match(plan.sql, /WHERE/);
  assert.equal(plan.localSteps.length, 2);
  assert.equal(plan.localSteps[0].operation.type, "fillDown");
});

