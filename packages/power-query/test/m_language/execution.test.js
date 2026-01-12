import assert from "node:assert/strict";
import test from "node:test";

import { DataTable } from "../../src/table.js";
import { QueryEngine } from "../../src/engine.js";
import { ODataConnector } from "../../src/connectors/odata.js";
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

test("m_language execution: OData.Feed + Table.SelectRows + Table.SelectColumns (folded)", async () => {
  /** @type {string[]} */
  const urls = [];
  const connector = new ODataConnector({
    fetch: async (url) => {
      urls.push(String(url));
      return {
        ok: true,
        status: 200,
        headers: {
          get: () => "application/json",
        },
        async json() {
          return { value: [{ Id: 1, Name: "A" }] };
        },
      };
    },
  });

  const script = `
let
  Source = OData.Feed("https://example.com/odata/Products"),
  #"Filtered Rows" = Table.SelectRows(Source, each [Price] > 20),
  #"Selected Columns" = Table.SelectColumns(#"Filtered Rows", {"Id", "Name"})
in
  #"Selected Columns"
`;

  const query = compileMToQuery(script);
  const engine = new QueryEngine({ connectors: { odata: connector } });
  const result = await engine.executeQuery(query, {}, {});

  assert.equal(urls.length, 1);
  assert.equal(urls[0], "https://example.com/odata/Products?$select=Id,Name&$filter=Price%20gt%2020");
  assert.deepEqual(result.toGrid(), [
    ["Id", "Name"],
    [1, "A"],
  ]);
});

test("m_language execution: OData.Feed + select/remove/take/skip (folded)", async () => {
  /** @type {string[]} */
  const urls = [];
  const connector = new ODataConnector({
    fetch: async (url) => {
      urls.push(String(url));
      return {
        ok: true,
        status: 200,
        headers: {
          get: () => "application/json",
        },
        async json() {
          return { value: [] };
        },
      };
    },
  });

  const script = `
let
  Source = OData.Feed("https://example.com/odata/Products"),
  #"Selected Columns" = Table.SelectColumns(Source, {"Id", "Name", "Price"}),
  #"Removed Columns" = Table.RemoveColumns(#"Selected Columns", {"Price"}),
  #"Kept First Rows" = Table.FirstN(#"Removed Columns", 10),
  #"Removed Top Rows" = Table.Skip(#"Kept First Rows", 3)
in
  #"Removed Top Rows"
`;

  const query = compileMToQuery(script);
  const engine = new QueryEngine({ connectors: { odata: connector } });
  const result = await engine.executeQuery(query, {}, {});

  assert.equal(urls.length, 1);
  assert.equal(urls[0], "https://example.com/odata/Products?$select=Id,Name&$skip=3&$top=7");
  assert.deepEqual(result.toGrid(), [["Id", "Name"]]);
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

test("m_language execution: promote/demote headers", async () => {
  const script = `
let
  Source = Range.FromValues(
    {
      {"Region", "Sales"},
      {"East", 100},
      {"West", 200}
    },
    [HasHeaders = false]
  ),
  #"Promoted Headers" = Table.PromoteHeaders(Source),
  #"Demoted Headers" = Table.DemoteHeaders(#"Promoted Headers")
in
  #"Demoted Headers"
`;

  const query = compileMToQuery(script);
  const engine = new QueryEngine();
  const result = await engine.executeQuery(query, {}, {});
  assert.deepEqual(result.toGrid(), [
    ["Column1", "Column2"],
    ["Region", "Sales"],
    ["East", 100],
    ["West", 200],
  ]);
});

test("m_language execution: addIndex/reorder/skip/firstN", async () => {
  const script = `
let
  Source = Range.FromValues({
    {"Region", "Sales"},
    {"East", 100},
    {"West", 200},
    {"South", 300}
  }),
  #"Added Index" = Table.AddIndexColumn(Source, "Index", 1, 1),
  #"Reordered Columns" = Table.ReorderColumns(#"Added Index", {"Index", "Region", "Sales"}),
  #"Removed Top Rows" = Table.Skip(#"Reordered Columns", 1),
  #"Kept First Rows" = Table.FirstN(#"Removed Top Rows", 1)
in
  #"Kept First Rows"
`;

  const query = compileMToQuery(script);
  const engine = new QueryEngine();
  const result = await engine.executeQuery(query, {}, {});
  assert.deepEqual(result.toGrid(), [
    ["Index", "Region", "Sales"],
    [2, "West", 200],
  ]);
});

test("m_language execution: whitelisted text/number/date functions in addColumn", async () => {
  const script = `
let
  Source = Range.FromValues({
    {"Name", "Sales", "When"},
    {" alice ", 12.345, "2020-01-01"},
    {"Bob", 67.891, "2020-01-02"}
  }),
  #"Upper Trimmed" = Table.AddColumn(Source, "NameUpper", each Text.Upper(Text.Trim([Name]))),
  #"Name Length" = Table.AddColumn(#"Upper Trimmed", "NameLen", each Text.Length([Name])),
  #"Has A" = Table.AddColumn(#"Name Length", "HasA", each Text.Contains([Name], "a")),
  #"Rounded" = Table.AddColumn(#"Has A", "SalesRounded", each Number.Round([Sales], 1)),
  #"Add Days" = Table.AddColumn(#"Rounded", "WhenPlus", each Date.AddDays(Date.FromText([When]), 1))
in
  #"Add Days"
`;

  const query = compileMToQuery(script);
  const engine = new QueryEngine();
  const result = await engine.executeQuery(query, {}, {});
  assert.deepEqual(result.toGrid(), [
    ["Name", "Sales", "When", "NameUpper", "NameLen", "HasA", "SalesRounded", "WhenPlus"],
    [" alice ", 12.345, "2020-01-01", "ALICE", 7, true, 12.3, new Date(Date.UTC(2020, 0, 2))],
    ["Bob", 67.891, "2020-01-02", "BOB", 3, false, 67.9, new Date(Date.UTC(2020, 0, 3))],
  ]);
});

test("m_language execution: combineColumns + splitColumn (named)", async () => {
  const script = `
let
  Source = Range.FromValues({
    {"First", "Last"},
    {"A", "B"},
    {"C", "D"}
  }),
  #"Merged Columns" = Table.CombineColumns(Source, {"First", "Last"}, Combiner.CombineTextByDelimiter("-", QuoteStyle.None), "Full"),
  #"Split Column" = Table.SplitColumn(#"Merged Columns", "Full", "-", {"First2", "Last2"})
in
  #"Split Column"
`;

  const query = compileMToQuery(script);
  const engine = new QueryEngine();
  const result = await engine.executeQuery(query, {}, {});
  assert.deepEqual(result.toGrid(), [
    ["First2", "Last2"],
    ["A", "B"],
    ["C", "D"],
  ]);
});

test("m_language execution: transformColumnNames", async () => {
  const script = `
let
  Source = Range.FromValues({
    {"Name", "Region"},
    {"A", "East"}
  }),
  #"Transformed Column Names" = Table.TransformColumnNames(Source, Text.Upper)
in
  #"Transformed Column Names"
`;

  const query = compileMToQuery(script);
  const engine = new QueryEngine();
  const result = await engine.executeQuery(query, {}, {});
  assert.deepEqual(result.toGrid(), [
    ["NAME", "REGION"],
    ["A", "East"],
  ]);
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
  // Table.Join compiles to a flat merge.
  assert.deepEqual(joinQuery.steps[0]?.operation.leftKeys, ["Id"]);
  assert.deepEqual(joinQuery.steps[0]?.operation.rightKeys, ["Id"]);
  assert.equal(joinQuery.steps[0]?.operation.joinMode, "flat");

  const engine = new QueryEngine();
  const result = await engine.executeQuery(joinQuery, { queries: { q_sales: sales, q_targets: targets } }, {});
  assert.deepEqual(result.toGrid(), [
    ["Id", "Sales", "Target"],
    [1, 100, null],
    [2, 200, "B"],
  ]);
});

test("m_language execution: Table.Join accepts numeric join kind constants", async () => {
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
    { id: "q_sales_num_join", name: "Sales" },
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
    { id: "q_targets_num_join", name: "Targets" },
  );

  const joinQuery = compileMToQuery(
    `
let
  Sales = Query.Reference("q_sales_num_join"),
  Targets = Query.Reference("q_targets_num_join"),
  #"Merged Queries" = Table.Join(Sales, {"Id"}, Targets, {"Id"}, 1)
in
  #"Merged Queries"
`,
    { id: "q_join_num", name: "Join numeric" },
  );

  assert.equal(joinQuery.steps[0]?.operation.type, "merge");
  assert.equal(joinQuery.steps[0]?.operation.joinType, "left");

  const engine = new QueryEngine();
  const result = await engine.executeQuery(joinQuery, { queries: { q_sales_num_join: sales, q_targets_num_join: targets } }, {});
  assert.deepEqual(result.toGrid(), [
    ["Id", "Sales", "Target"],
    [1, 100, null],
    [2, 200, "B"],
  ]);
});

test("m_language execution: Table.Join accepts JoinAlgorithm constants", async () => {
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
    { id: "q_sales_join_alg", name: "Sales" },
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
    { id: "q_targets_join_alg", name: "Targets" },
  );

  const joinQuery = compileMToQuery(
    `
let
  Sales = Query.Reference("q_sales_join_alg"),
  Targets = Query.Reference("q_targets_join_alg"),
  #"Merged Queries" = Table.Join(Sales, {"Id"}, Targets, {"Id"}, JoinKind.LeftOuter, JoinAlgorithm.LeftHash)
in
  #"Merged Queries"
`,
    { id: "q_join_alg", name: "Join algorithm" },
  );

  assert.equal(joinQuery.steps[0]?.operation.type, "merge");
  assert.equal(joinQuery.steps[0]?.operation.joinAlgorithm, "leftHash");

  const engine = new QueryEngine();
  const result = await engine.executeQuery(joinQuery, { queries: { q_sales_join_alg: sales, q_targets_join_alg: targets } }, {});
  assert.deepEqual(result.toGrid(), [
    ["Id", "Sales", "Target"],
    [1, 100, null],
    [2, 200, "B"],
  ]);
});

test("m_language execution: JoinKind/Comparer constants can be bound in let expressions", async () => {
  const left = compileMToQuery(
    `
let
  Source = Range.FromValues({
    {"Id", "Name"},
    {1, "Alice"},
    {2, "Bob"}
  })
in
  Source
`,
    { id: "q_left_const_binding", name: "Left" },
  );

  const right = compileMToQuery(
    `
let
  Source = Range.FromValues({
    {"Name", "Score"},
    {"alice", 10},
    {"BOB", 20}
  })
in
  Source
`,
    { id: "q_right_const_binding", name: "Right" },
  );

  const joinQuery = compileMToQuery(
    `
let
  Left = Query.Reference("q_left_const_binding"),
  Right = Query.Reference("q_right_const_binding"),
  Kind = JoinKind.Inner,
  Comp = Comparer.OrdinalIgnoreCase,
  #"Merged Queries" = Table.Join(Left, {"Name"}, Right, {"Name"}, Kind, null, Comp)
in
  #"Merged Queries"
`,
    { id: "q_join_const_binding", name: "Join constants" },
  );

  assert.equal(joinQuery.steps[0]?.operation.type, "merge");
  assert.equal(joinQuery.steps[0]?.operation.joinType, "inner");
  assert.equal(joinQuery.steps[0]?.operation.comparer?.caseSensitive, false);

  const engine = new QueryEngine();
  const result = await engine.executeQuery(
    joinQuery,
    { queries: { q_left_const_binding: left, q_right_const_binding: right } },
    {},
  );
  assert.deepEqual(result.toGrid(), [
    ["Id", "Name", "Score"],
    [1, "Alice", 10],
    [2, "Bob", 20],
  ]);
});

test("m_language execution: Table.Join supports Comparer.OrdinalIgnoreCase", async () => {
  const left = compileMToQuery(
    `
let
  Source = Range.FromValues({
    {"Id", "Name"},
    {1, "Alice"},
    {2, "Bob"}
  })
in
  Source
`,
    { id: "q_left_case", name: "Left" },
  );

  const right = compileMToQuery(
    `
let
  Source = Range.FromValues({
    {"Name", "Score"},
    {"alice", 10},
    {"BOB", 20}
  })
in
  Source
`,
    { id: "q_right_case", name: "Right" },
  );

  const joinQuery = compileMToQuery(
    `
let
  Left = Query.Reference("q_left_case"),
  Right = Query.Reference("q_right_case"),
  #"Merged Queries" = Table.Join(Left, {"Name"}, Right, {"Name"}, JoinKind.Inner, null, Comparer.OrdinalIgnoreCase)
in
  #"Merged Queries"
`,
    { id: "q_join_case", name: "Join Case" },
  );

  assert.equal(joinQuery.steps[0]?.operation.type, "merge");
  assert.equal(joinQuery.steps[0]?.operation.comparer?.caseSensitive, false);

  const engine = new QueryEngine();
  const result = await engine.executeQuery(joinQuery, { queries: { q_left_case: left, q_right_case: right } }, {});
  assert.deepEqual(result.toGrid(), [
    ["Id", "Name", "Score"],
    [1, "Alice", 10],
    [2, "Bob", 20],
  ]);
});

test("m_language execution: Table.Join ordinal comparer record is case sensitive", async () => {
  const left = compileMToQuery(
    `
let
  Source = Range.FromValues({
    {"Id", "Name"},
    {1, "Alice"}
  })
in
  Source
`,
    { id: "q_left_case_sensitive", name: "Left" },
  );

  const right = compileMToQuery(
    `
let
  Source = Range.FromValues({
    {"Name", "Score"},
    {"alice", 10}
  })
in
  Source
`,
    { id: "q_right_case_sensitive", name: "Right" },
  );

  const joinQuery = compileMToQuery(
    `
let
  Left = Query.Reference("q_left_case_sensitive"),
  Right = Query.Reference("q_right_case_sensitive"),
  #"Merged Queries" = Table.Join(Left, {"Name"}, Right, {"Name"}, JoinKind.Inner, null, [comparer = "ordinal"])
in
  #"Merged Queries"
`,
    { id: "q_join_case_sensitive", name: "Join Case Sensitive" },
  );

  assert.equal(joinQuery.steps[0]?.operation.type, "merge");
  assert.equal(joinQuery.steps[0]?.operation.comparer?.caseSensitive, true);

  const engine = new QueryEngine();
  const result = await engine.executeQuery(
    joinQuery,
    { queries: { q_left_case_sensitive: left, q_right_case_sensitive: right } },
    {},
  );
  assert.deepEqual(result.toGrid(), [["Id", "Name", "Score"]]);
});

test("m_language execution: Table.Join supports comparer lists (per-key)", async () => {
  const left = compileMToQuery(
    `
let
  Source = Range.FromValues({
    {"First", "Last", "Id"},
    {"Alice", "Smith", 1},
    {"ALICE", "Smith", 2},
    {"Alice", "SMITH", 3}
  })
in
  Source
`,
    { id: "q_left_comparer_list", name: "Left" },
  );

  const right = compileMToQuery(
    `
let
  Source = Range.FromValues({
    {"First", "Last", "Score"},
    {"alice", "Smith", 10},
    {"alice", "SMITH", 20}
  })
in
  Source
`,
    { id: "q_right_comparer_list", name: "Right" },
  );

  const joinQuery = compileMToQuery(
    `
let
  Left = Query.Reference("q_left_comparer_list"),
  Right = Query.Reference("q_right_comparer_list"),
  #"Merged Queries" = Table.Join(Left, {"First", "Last"}, Right, {"First", "Last"}, JoinKind.Inner, null, {Comparer.OrdinalIgnoreCase, Comparer.Ordinal})
in
  #"Merged Queries"
`,
    { id: "q_join_comparer_list", name: "Join comparer list" },
  );

  assert.equal(joinQuery.steps[0]?.operation.type, "merge");
  assert.deepEqual(joinQuery.steps[0]?.operation.comparers, [
    { comparer: "ordinalIgnoreCase", caseSensitive: false },
    { comparer: "ordinal", caseSensitive: true },
  ]);

  const engine = new QueryEngine();
  const result = await engine.executeQuery(
    joinQuery,
    { queries: { q_left_comparer_list: left, q_right_comparer_list: right } },
    {},
  );
  assert.deepEqual(result.toGrid(), [
    ["First", "Last", "Id", "Score"],
    ["Alice", "Smith", 1, 10],
    ["ALICE", "Smith", 2, 10],
    ["Alice", "SMITH", 3, 20],
  ]);
});

test("m_language execution: Table.Join supports multi-key joins (nulls compare equal)", async () => {
  const left = compileMToQuery(
    `
let
  Source = Range.FromValues({
    {"Id", "Region", "Sales"},
    {1, "East", 100},
    {1, "West", 200},
    {2, "East", 300},
    {3, null, 400}
  })
in
  Source
`,
    { id: "q_left", name: "Left" },
  );

  const right = compileMToQuery(
    `
let
  Source = Range.FromValues({
    {"Id", "Region", "Target"},
    {1, "East", "A"},
    {1, "West", "B"},
    {3, null, "C"}
  })
in
  Source
`,
    { id: "q_right", name: "Right" },
  );

  const joinQuery = compileMToQuery(
    `
let
  Left = Query.Reference("q_left"),
  Right = Query.Reference("q_right"),
  #"Merged Queries" = Table.Join(Left, {"Id", "Region"}, Right, {"Id", "Region"}, JoinKind.LeftOuter)
in
  #"Merged Queries"
`,
    { id: "q_join_multi", name: "Join Multi" },
  );

  assert.equal(joinQuery.steps[0]?.operation.type, "merge");
  assert.deepEqual(joinQuery.steps[0]?.operation.leftKeys, ["Id", "Region"]);
  assert.deepEqual(joinQuery.steps[0]?.operation.rightKeys, ["Id", "Region"]);

  const engine = new QueryEngine();
  const result = await engine.executeQuery(joinQuery, { queries: { q_left: left, q_right: right } }, {});
  assert.deepEqual(result.toGrid(), [
    ["Id", "Region", "Sales", "Target"],
    [1, "East", 100, "A"],
    [1, "West", 200, "B"],
    [2, "East", 300, null],
    [3, null, 400, "C"],
  ]);
});

test("m_language execution: Table.NestedJoin produces nested table columns", async () => {
  const left = compileMToQuery(
    `
let
  Source = Range.FromValues({
    {"Id", "Sales"},
    {1, 100},
    {2, 200},
    {3, 300}
  })
in
  Source
`,
    { id: "q_left_nested", name: "Left" },
  );

  const right = compileMToQuery(
    `
let
  Source = Range.FromValues({
    {"Id", "Target"},
    {2, "B"},
    {2, "C"}
  })
in
  Source
`,
    { id: "q_right_nested", name: "Right" },
  );

  const nestedJoin = compileMToQuery(
    `
let
  Left = Query.Reference("q_left_nested"),
  Right = Query.Reference("q_right_nested"),
  #"Merged Queries" = Table.NestedJoin(Left, {"Id"}, Right, {"Id"}, "Matches", JoinKind.LeftOuter)
in
  #"Merged Queries"
`,
    { id: "q_nested_join", name: "Nested join" },
  );

  assert.equal(nestedJoin.steps[0]?.operation.type, "merge");
  assert.equal(nestedJoin.steps[0]?.operation.joinMode, "nested");
  assert.equal(nestedJoin.steps[0]?.operation.newColumnName, "Matches");

  const engine = new QueryEngine();
  const result = await engine.executeQuery(nestedJoin, { queries: { q_left_nested: left, q_right_nested: right } }, {});

  const matchesIdx = result.getColumnIndex("Matches");

  const row0 = result.getCell(0, matchesIdx);
  assert.ok(row0 instanceof DataTable);
  assert.deepEqual(row0.toGrid(), [["Id", "Target"]]);

  const row1 = result.getCell(1, matchesIdx);
  assert.ok(row1 instanceof DataTable);
  assert.deepEqual(row1.toGrid(), [
    ["Id", "Target"],
    [2, "B"],
    [2, "C"],
  ]);
});

test("m_language execution: Table.AddJoinColumn is an alias for Table.NestedJoin", async () => {
  const left = compileMToQuery(
    `
let
  Source = Range.FromValues({
    {"Id", "Sales"},
    {1, 100},
    {2, 200},
    {3, 300}
  })
in
  Source
`,
    { id: "q_left_add_join", name: "Left" },
  );

  const right = compileMToQuery(
    `
let
  Source = Range.FromValues({
    {"Id", "Target"},
    {2, "B"},
    {2, "C"}
  })
in
  Source
`,
    { id: "q_right_add_join", name: "Right" },
  );

  const nestedJoin = compileMToQuery(
    `
let
  Left = Query.Reference("q_left_add_join"),
  Right = Query.Reference("q_right_add_join"),
  #"Merged Queries" = Table.AddJoinColumn(Left, {"Id"}, Right, {"Id"}, "Matches", JoinKind.LeftOuter)
in
  #"Merged Queries"
`,
    { id: "q_add_join", name: "AddJoinColumn" },
  );

  assert.equal(nestedJoin.steps[0]?.operation.type, "merge");
  assert.equal(nestedJoin.steps[0]?.operation.joinMode, "nested");
  assert.equal(nestedJoin.steps[0]?.operation.newColumnName, "Matches");

  const engine = new QueryEngine();
  const result = await engine.executeQuery(nestedJoin, { queries: { q_left_add_join: left, q_right_add_join: right } }, {});

  const matchesIdx = result.getColumnIndex("Matches");

  const row0 = result.getCell(0, matchesIdx);
  assert.ok(row0 instanceof DataTable);
  assert.deepEqual(row0.toGrid(), [["Id", "Target"]]);

  const row1 = result.getCell(1, matchesIdx);
  assert.ok(row1 instanceof DataTable);
  assert.deepEqual(row1.toGrid(), [
    ["Id", "Target"],
    [2, "B"],
    [2, "C"],
  ]);
});

test("m_language execution: Table.NestedJoin supports comparer lists (per-key)", async () => {
  const left = compileMToQuery(
    `
let
  Source = Range.FromValues({
    {"First", "Last", "Id"},
    {"Alice", "Smith", 1},
    {"ALICE", "Smith", 2},
    {"Alice", "SMITH", 3},
    {"Bob", "Smith", 4}
  })
in
  Source
`,
    { id: "q_left_nested_comparers", name: "Left" },
  );

  const right = compileMToQuery(
    `
let
  Source = Range.FromValues({
    {"First", "Last", "Score"},
    {"alice", "Smith", 10},
    {"alice", "SMITH", 20}
  })
in
  Source
`,
    { id: "q_right_nested_comparers", name: "Right" },
  );

  const nestedJoin = compileMToQuery(
    `
let
  Left = Query.Reference("q_left_nested_comparers"),
  Right = Query.Reference("q_right_nested_comparers"),
  #"Merged Queries" = Table.NestedJoin(
    Left,
    {"First", "Last"},
    Right,
    {"First", "Last"},
    "Matches",
    JoinKind.LeftOuter,
    null,
    {Comparer.OrdinalIgnoreCase, Comparer.Ordinal}
  )
in
  #"Merged Queries"
`,
    { id: "q_nested_join_comparers", name: "Nested join comparers" },
  );

  assert.equal(nestedJoin.steps[0]?.operation.type, "merge");
  assert.deepEqual(nestedJoin.steps[0]?.operation.comparers, [
    { comparer: "ordinalIgnoreCase", caseSensitive: false },
    { comparer: "ordinal", caseSensitive: true },
  ]);

  const engine = new QueryEngine();
  const result = await engine.executeQuery(
    nestedJoin,
    { queries: { q_left_nested_comparers: left, q_right_nested_comparers: right } },
    {},
  );

  const matchesIdx = result.getColumnIndex("Matches");

  const row0 = result.getCell(0, matchesIdx);
  assert.ok(row0 instanceof DataTable);
  assert.deepEqual(row0.toGrid(), [
    ["First", "Last", "Score"],
    ["alice", "Smith", 10],
  ]);

  const row1 = result.getCell(1, matchesIdx);
  assert.ok(row1 instanceof DataTable);
  assert.deepEqual(row1.toGrid(), [
    ["First", "Last", "Score"],
    ["alice", "Smith", 10],
  ]);

  const row2 = result.getCell(2, matchesIdx);
  assert.ok(row2 instanceof DataTable);
  assert.deepEqual(row2.toGrid(), [
    ["First", "Last", "Score"],
    ["alice", "SMITH", 20],
  ]);

  const row3 = result.getCell(3, matchesIdx);
  assert.ok(row3 instanceof DataTable);
  assert.deepEqual(row3.toGrid(), [["First", "Last", "Score"]]);
});

test("m_language execution: Table.ExpandTableColumn expands nested joins (empty nested tables keep the row)", async () => {
  const left = compileMToQuery(
    `
let
  Source = Range.FromValues({
    {"Id", "Target", "Sales"},
    {1, "L1", 100},
    {2, "L2", 200},
    {3, "L3", 300}
  })
in
  Source
`,
    { id: "q_left_expand", name: "Left" },
  );

  const right = compileMToQuery(
    `
let
  Source = Range.FromValues({
    {"Id", "Target"},
    {2, "R2"},
    {2, "R2b"}
  })
in
  Source
`,
    { id: "q_right_expand", name: "Right" },
  );

  const query = compileMToQuery(
    `
let
  Left = Query.Reference("q_left_expand"),
  Right = Query.Reference("q_right_expand"),
  #"Merged Queries" = Table.NestedJoin(Left, {"Id"}, Right, {"Id"}, "Matches", JoinKind.LeftOuter),
  #"Expanded Matches" = Table.ExpandTableColumn(#"Merged Queries", "Matches", {"Target"})
in
  #"Expanded Matches"
`,
    { id: "q_expand", name: "Expand" },
  );

  assert.deepEqual(query.steps.map((s) => s.operation.type), ["merge", "expandTableColumn"]);

  const engine = new QueryEngine();
  const result = await engine.executeQuery(query, { queries: { q_left_expand: left, q_right_expand: right } }, {});
  assert.deepEqual(result.toGrid(), [
    ["Id", "Target", "Sales", "Target.1"],
    [1, "L1", 100, null],
    [2, "L2", 200, "R2"],
    [2, "L2", 200, "R2b"],
    [3, "L3", 300, null],
  ]);
});
