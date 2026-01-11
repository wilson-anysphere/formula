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

