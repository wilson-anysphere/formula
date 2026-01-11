import assert from "node:assert/strict";
import test from "node:test";

import { DataTable } from "../../src/table.js";
import { QueryEngine } from "../../src/engine.js";
import { compileMToQuery } from "../../src/m/compiler.js";

test("m_language execution: range -> filter/group/sort", async () => {
  const script = `
let
  Source = Range.FromValues({
    {"Date", "Region", "Sales"},
    {#date(2024, 1, 1), "East", 100},
    {#date(2024, 1, 2), "East", 150},
    {#date(2024, 1, 3), "West", 200}
  }),
  #"Filtered Rows" = Table.SelectRows(Source, each [Date] >= #date(2024, 1, 2) and [Region] = "East"),
  #"Grouped Rows" = Table.Group(#"Filtered Rows", {"Region"}, {{"Total Sales", each List.Sum([Sales])}}),
  #"Sorted Rows" = Table.Sort(#"Grouped Rows", {{"Total Sales", Order.Descending}})
in
  #"Sorted Rows"
`;

  const query = compileMToQuery(script);
  const engine = new QueryEngine();
  const result = await engine.executeQuery(query, {}, {});
  assert.deepEqual(result.toGrid(), [
    ["Region", "Total Sales"],
    ["East", 150],
  ]);
});

test("m_language execution: add/rename/type/select/remove", async () => {
  const script = `
let
  Source = Range.FromValues({
    {"Region", "Sales", "Product"},
    {"East", "100", "A"},
    {"West", "200", "B"}
  }),
  #"Changed Type" = Table.TransformColumnTypes(Source, {{"Sales", type number}}),
  #"Added Column" = Table.AddColumn(#"Changed Type", "Double", each [Sales] * 2),
  #"Removed Columns" = Table.RemoveColumns(#"Added Column", {"Product"}),
  #"Selected Columns" = Table.SelectColumns(#"Removed Columns", {"Region", "Sales", "Double"}),
  #"Renamed Columns" = Table.RenameColumns(#"Selected Columns", {{"Double", "Double Sales"}, {"Region", "Area"}})
in
  #"Renamed Columns"
`;

  const query = compileMToQuery(script);
  const engine = new QueryEngine();
  const result = await engine.executeQuery(query, {}, {});
  assert.deepEqual(result.toGrid(), [
    ["Area", "Sales", "Double Sales"],
    ["East", 100, 200],
    ["West", 200, 400],
  ]);
});

test("m_language execution: filldown/replace/split", async () => {
  const script = `
let
  Source = Range.FromValues({
    {"Region", "FullName"},
    {"East", "A,B"},
    {null, "C,D"},
    {"West", "E,F"}
  }),
  #"Filled Down" = Table.FillDown(Source, {"Region"}),
  #"Replaced Value" = Table.ReplaceValue(#"Filled Down", "West", "W", Replacer.ReplaceText, {"Region"}),
  #"Split Column" = Table.SplitColumn(#"Replaced Value", "FullName", ",")
in
  #"Split Column"
`;

  const query = compileMToQuery(script);
  const engine = new QueryEngine();
  const result = await engine.executeQuery(query, {}, {});
  assert.deepEqual(result.toGrid(), [
    ["Region", "FullName", "FullName.2"],
    ["East", "A", "B"],
    ["East", "C", "D"],
    ["W", "E", "F"],
  ]);
});

test("m_language execution: Table.SelectRows compares records by value", async () => {
  const script = `
let
  Source = Range.FromValues({
    {"Key", "Value"},
    {[a=1, b=2], 1},
    {[b=2, a=1], 2},
    {[a=1, b=3], 3}
  }),
  #"Filtered Rows" = Table.SelectRows(Source, each [Key] = [a=1, b=2])
in
  #"Filtered Rows"
`;

  const query = compileMToQuery(script);
  const engine = new QueryEngine();
  const result = await engine.executeQuery(query, {}, {});
  assert.deepEqual(result.toGrid(), [
    ["Key", "Value"],
    [{ a: 1, b: 2 }, 1],
    [{ b: 2, a: 1 }, 2],
  ]);
});

test("m_language execution: constant equality compares dates and records by value", async () => {
  const script = `
let
  Source = Range.FromValues({
    {"Value"},
    {#date(2024, 1, 1) = #date(2024, 1, 1)},
    {#date(2024, 1, 1) <> #date(2024, 1, 1)},
    {[a=1, b=2] = [b=2, a=1]}
  })
in
  Source
`;

  const query = compileMToQuery(script);
  const engine = new QueryEngine();
  const result = await engine.executeQuery(query, {}, {});
  assert.deepEqual(result.toGrid(), [["Value"], [true], [false], [true]]);
});

test("m_language execution: Web.Contents + Table.RemoveColumns", async () => {
  const script = `
let
  Source = Web.Contents("https://example.com/data", [Method="POST", Headers=[Authorization="Bearer token"]]),
  #"Removed Columns" = Table.RemoveColumns(Source, {"Secret"})
in
  #"Removed Columns"
`;

  const query = compileMToQuery(script);
  const engine = new QueryEngine({
    apiAdapter: {
      fetchTable: async (_url, _options) =>
        DataTable.fromGrid(
          [
            ["Secret", "Value"],
            ["x", 1],
          ],
          { hasHeaders: true, inferTypes: true },
        ),
    },
  });
  const result = await engine.executeQuery(query, {}, {});
  assert.deepEqual(result.toGrid(), [["Value"], [1]]);
});

test("m_language execution: Odbc.Query + Table.SelectColumns", async () => {
  const script = `
let
  Source = Odbc.Query("dsn=mydb", "select * from sales"),
  #"Selected Columns" = Table.SelectColumns(Source, {"Region"})
in
  #"Selected Columns"
`;

  const query = compileMToQuery(script);
  const engine = new QueryEngine({
    databaseAdapter: {
      querySql: async (_connection, _sql) =>
        DataTable.fromGrid(
          [
            ["Region", "Sales"],
            ["East", 100],
            ["West", 200],
          ],
          { hasHeaders: true, inferTypes: true },
        ),
    },
  });
  const result = await engine.executeQuery(query, {}, {});
  assert.deepEqual(result.toGrid(), [
    ["Region"],
    ["East"],
    ["West"],
  ]);
});

test("m_language execution: if/then/else in addColumn", async () => {
  const script = `
let
  Source = Range.FromValues({
    {"Sales"},
    {10},
    {-5},
    {0}
  }),
  #"Added Column" = Table.AddColumn(Source, "Positive", each if [Sales] > 0 then [Sales] else null)
in
  #"Added Column"
`;

  const query = compileMToQuery(script);
  const engine = new QueryEngine();
  const result = await engine.executeQuery(query, {}, {});
  assert.deepEqual(result.toGrid(), [
    ["Sales", "Positive"],
    [10, 10],
    [-5, null],
    [0, null],
  ]);
});

test("m_language execution: if/then/else in filter predicate", async () => {
  const script = `
let
  Source = Range.FromValues({
    {"Sales"},
    {50},
    {150}
  }),
  #"Filtered Rows" = Table.SelectRows(Source, each if [Sales] > 100 then true else false)
in
  #"Filtered Rows"
`;

  const query = compileMToQuery(script);
  const engine = new QueryEngine();
  const result = await engine.executeQuery(query, {}, {});
  assert.deepEqual(result.toGrid(), [["Sales"], [150]]);
});

test("m_language execution: Table.Combine + Table.Distinct", async () => {
  const q1 = compileMToQuery(
    `
let
  Source = Range.FromValues({
    {"Id"},
    {1},
    {2}
  })
in
  Source
`,
    { id: "q1", name: "Q1" },
  );

  const q2 = compileMToQuery(
    `
let
  Source = Range.FromValues({
    {"Id"},
    {2},
    {3}
  })
in
  Source
`,
    { id: "q2", name: "Q2" },
  );

  const combined = compileMToQuery(
    `
let
  A = Query.Reference("q1"),
  B = Query.Reference("q2"),
  #"Appended Queries" = Table.Combine({A, B}),
  #"Removed Duplicates" = Table.Distinct(#"Appended Queries")
in
  #"Removed Duplicates"
`,
    { id: "q_combined", name: "Combined" },
  );

  assert.deepEqual(combined.steps.map((s) => s.operation.type), ["append", "distinctRows"]);

  const engine = new QueryEngine();
  const result = await engine.executeQuery(combined, { queries: { q1, q2 } }, {});
  assert.deepEqual(result.toGrid(), [
    ["Id"],
    [1],
    [2],
    [3],
  ]);
});

test("m_language execution: Table.Join maps to merge", async () => {
  const sales = compileMToQuery(
    `
let
  Source = Range.FromValues({
    {"Id", "Sales"},
    {1, 100},
    {2, 200}
  })
in
  Source
`,
    { id: "q_sales", name: "Sales" },
  );

  const targets = compileMToQuery(
    `
let
  Source = Range.FromValues({
    {"Id", "Target"},
    {2, "B"}
  })
in
  Source
`,
    { id: "q_targets", name: "Targets" },
  );

  const joinQuery = compileMToQuery(
    `
let
  Sales = Query.Reference("q_sales"),
  Targets = Query.Reference("q_targets"),
  #"Merged Queries" = Table.Join(Sales, {"Id"}, Targets, {"Id"}, JoinKind.LeftOuter)
in
  #"Merged Queries"
`,
    { id: "q_join", name: "Join" },
  );

  assert.equal(joinQuery.steps[0]?.operation.type, "merge");

  const engine = new QueryEngine();
  const result = await engine.executeQuery(joinQuery, { queries: { q_sales: sales, q_targets: targets } }, {});
  assert.deepEqual(result.toGrid(), [
    ["Id", "Sales", "Target"],
    [1, 100, null],
    [2, 200, "B"],
  ]);
});
