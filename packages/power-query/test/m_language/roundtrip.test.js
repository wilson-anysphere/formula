import assert from "node:assert/strict";
import test from "node:test";

import { compileMToQuery } from "../../src/m/compiler.js";
import { prettyPrintQueryToM } from "../../src/m/pretty.js";

/**
 * @param {unknown} value
 * @returns {any}
 */
function toJson(value) {
  return JSON.parse(JSON.stringify(value));
}

test("m_language round-trip: prettyPrintQueryToM", () => {
  const script = `
let
  Source = Range.FromValues({
    {"Region", "Sales"},
    {"East", 100},
    {"West", 200}
  }),
  #"Filtered Rows" = Table.SelectRows(Source, each [Region] = "East"),
  #"Added Column" = Table.AddColumn(#"Filtered Rows", "Double", each [Sales] * 2)
in
  #"Added Column"
`;

  const query = compileMToQuery(script);
  const printed = prettyPrintQueryToM(query);
  const query2 = compileMToQuery(printed);
  assert.deepEqual(toJson(query2), toJson(query));
});

test("m_language round-trip: prettyPrintQueryToM (if/then/else)", () => {
  const script = `
let
  Source = Range.FromValues({
    {"Sales"},
    {50},
    {150}
  }),
  #"Added Column" = Table.AddColumn(Source, "Bucket", each if [Sales] > 100 then "High" else "Low")
in
  #"Added Column"
`;

  const query = compileMToQuery(script);
  const printed = prettyPrintQueryToM(query);
  const query2 = compileMToQuery(printed);
  assert.deepEqual(toJson(query2), toJson(query));
});

test("m_language round-trip: prettyPrintQueryToM (OData.Feed)", () => {
  const script = `
let
  Source = OData.Feed("https://example.com/odata/Products", [Headers=[Authorization="Bearer token"]]),
  #"Selected Columns" = Table.SelectColumns(Source, {"Id", "Name"})
in
  #"Selected Columns"
`;

  const query = compileMToQuery(script);
  const printed = prettyPrintQueryToM(query);
  const query2 = compileMToQuery(printed);
  assert.deepEqual(toJson(query2), toJson(query));
});

test("m_language round-trip: prettyPrintQueryToM (OData.Feed options)", () => {
  const script = `
let
  Source = OData.Feed(
    "https://example.com/odata/Products",
    [
      Headers = [Authorization = "Bearer token", Accept = "application/json"],
      Auth = [Type = "OAuth2", ProviderId = "example-provider", Scopes = {"scope1", "scope2"}],
      RowsPath = "data.items"
    ]
  ),
  #"Selected Columns" = Table.SelectColumns(Source, {"Id"})
in
  #"Selected Columns"
`;

  const query = compileMToQuery(script);
  const printed = prettyPrintQueryToM(query);
  const query2 = compileMToQuery(printed);
  assert.deepEqual(toJson(query2), toJson(query));
});

test("m_language round-trip: prettyPrintQueryToM (whitelisted formula calls + new table ops)", () => {
  const script = `
let
  Source = Range.FromValues({
    {"Name", "Sales", "When"},
    {" alice ", 12.345, "2020-01-01"},
    {"Bob", 67.891, "2020-01-02"}
  }),
  #"Upper Trimmed" = Table.AddColumn(Source, "NameUpper", each Text.Upper(Text.Trim([Name]))),
  #"Rounded" = Table.AddColumn(#"Upper Trimmed", "SalesRounded", each Number.Round([Sales], 1)),
  #"Add Days" = Table.AddColumn(#"Rounded", "WhenPlus", each Date.AddDays(Date.FromText([When]), 1)),
  #"Added Index" = Table.AddIndexColumn(#"Add Days", "Index"),
  #"Reordered Columns" = Table.ReorderColumns(#"Added Index", {"Index", "Name", "Sales", "When", "NameUpper", "SalesRounded", "WhenPlus"})
in
  #"Reordered Columns"
`;

  const query = compileMToQuery(script);
  const printed = prettyPrintQueryToM(query);
  const query2 = compileMToQuery(printed);
  assert.deepEqual(toJson(query2), toJson(query));
});

test("m_language round-trip: prettyPrintQueryToM (nested join + expand)", () => {
  const script = `
let
  Left = Query.Reference("q_left"),
  Right = Query.Reference("q_right"),
  #"Merged Queries" = Table.NestedJoin(Left, {"Id"}, Right, {"Id"}, "Matches", JoinKind.LeftOuter),
  #"Expanded Matches" = Table.ExpandTableColumn(#"Merged Queries", "Matches", {"Target"})
in
  #"Expanded Matches"
`;

  const query = compileMToQuery(script);
  const printed = prettyPrintQueryToM(query);
  const query2 = compileMToQuery(printed);
  assert.deepEqual(toJson(query2), toJson(query));
});

test("m_language round-trip: prettyPrintQueryToM (add join column)", () => {
  const script = `
let
  Left = Query.Reference("q_left"),
  Right = Query.Reference("q_right"),
  #"Merged Queries" = Table.AddJoinColumn(Left, {"Id"}, Right, {"Id"}, "Matches", JoinKind.LeftOuter)
in
  #"Merged Queries"
`;

  const query = compileMToQuery(script);
  const printed = prettyPrintQueryToM(query);
  const query2 = compileMToQuery(printed);
  assert.deepEqual(toJson(query2), toJson(query));
});

test("m_language round-trip: prettyPrintQueryToM (join comparer)", () => {
  const script = `
let
  Left = Query.Reference("q_left"),
  Right = Query.Reference("q_right"),
  #"Merged Queries" = Table.Join(Left, {"Name"}, Right, {"Name"}, JoinKind.Inner, null, Comparer.OrdinalIgnoreCase)
in
  #"Merged Queries"
`;

  const query = compileMToQuery(script);
  const printed = prettyPrintQueryToM(query);
  const query2 = compileMToQuery(printed);
  assert.deepEqual(toJson(query2), toJson(query));
});

test("m_language round-trip: prettyPrintQueryToM (join algorithm + comparer)", () => {
  const script = `
let
  Left = Query.Reference("q_left"),
  Right = Query.Reference("q_right"),
  #"Merged Queries" = Table.Join(Left, {"Name"}, Right, {"Name"}, JoinKind.Inner, JoinAlgorithm.LeftHash, Comparer.OrdinalIgnoreCase)
in
  #"Merged Queries"
`;

  const query = compileMToQuery(script);
  const printed = prettyPrintQueryToM(query);
  const query2 = compileMToQuery(printed);
  assert.deepEqual(toJson(query2), toJson(query));
});

test("m_language round-trip: prettyPrintQueryToM (join comparer list)", () => {
  const script = `
let
  Left = Query.Reference("q_left"),
  Right = Query.Reference("q_right"),
  #"Merged Queries" = Table.Join(Left, {"First", "Last"}, Right, {"First", "Last"}, JoinKind.Inner, null, {Comparer.OrdinalIgnoreCase, Comparer.Ordinal})
in
  #"Merged Queries"
`;

  const query = compileMToQuery(script);
  const printed = prettyPrintQueryToM(query);
  const query2 = compileMToQuery(printed);
  assert.deepEqual(toJson(query2), toJson(query));
});

test("m_language round-trip: prettyPrintQueryToM (nested join comparer list)", () => {
  const script = `
let
  Left = Query.Reference("q_left"),
  Right = Query.Reference("q_right"),
  #"Merged Queries" = Table.NestedJoin(Left, {"First", "Last"}, Right, {"First", "Last"}, "Matches", JoinKind.LeftOuter, null, {Comparer.OrdinalIgnoreCase, Comparer.Ordinal})
in
  #"Merged Queries"
`;

  const query = compileMToQuery(script);
  const printed = prettyPrintQueryToM(query);
  const query2 = compileMToQuery(printed);
  assert.deepEqual(toJson(query2), toJson(query));
});

test("m_language round-trip: prettyPrintQueryToM (expanded scalar types)", () => {
  const script = `
let
  Source = Range.FromValues({
    {"When", "Zone", "Time", "Dur", "Dec", "Bin"},
    {"2024-01-01T01:02:03.004Z", "2024-01-01T01:02:03.004Z", "06:00:00", "P1DT12H", "123.450", "AQID"}
  }),
  #"Changed Type" = Table.TransformColumnTypes(Source, {{"When", type datetime}}),
  #"Changed Type 2" = Table.TransformColumnTypes(#"Changed Type", {{"Zone", type datetimezone}}),
  #"Changed Type 3" = Table.TransformColumnTypes(#"Changed Type 2", {{"Time", type time}}),
  #"Changed Type 4" = Table.TransformColumnTypes(#"Changed Type 3", {{"Dur", type duration}}),
  #"Changed Type 5" = Table.TransformColumnTypes(#"Changed Type 4", {{"Dec", Decimal.Type}}),
  #"Changed Type 6" = Table.TransformColumnTypes(#"Changed Type 5", {{"Bin", type binary}})
in
  #"Changed Type 6"
`;

  const query = compileMToQuery(script);
  const printed = prettyPrintQueryToM(query);
  const query2 = compileMToQuery(printed);
  assert.deepEqual(toJson(query2), toJson(query));
});

test("m_language round-trip: prettyPrintQueryToM (#datetime fractional seconds)", () => {
  const script = `
let
  Source = Range.FromValues({
    {"When"},
    {#datetime(2024, 1, 1, 1, 2, 3.004)}
  })
in
  Source
`;

  const query = compileMToQuery(script);
  const printed = prettyPrintQueryToM(query);
  assert.match(printed, /#datetime\(2024, 1, 1, 1, 2, 3\.004\)/);
  const query2 = compileMToQuery(printed);
  assert.deepEqual(toJson(query2), toJson(query));
});

test("m_language round-trip: prettyPrintQueryToM (#time/#duration/#datetimezone constants)", () => {
  const script = `
let
  Source = Range.FromValues({
    {"Zone", "Time", "Dur"},
    {#datetimezone(2024, 1, 1, 1, 2, 3.004, 2, 0), #time(6, 30, 0), #duration(1, 12, 0, 0)}
  })
in
  Source
`;

  const query = compileMToQuery(script);
  const printed = prettyPrintQueryToM(query);
  assert.match(printed, /#datetimezone\(2024, 1, 1, 1, 2, 3\.004, 2, 0\)/);
  assert.match(printed, /#time\(6, 30, 0\)/);
  assert.match(printed, /#duration\(1, 12, 0, 0\)/);
  const query2 = compileMToQuery(printed);
  assert.deepEqual(toJson(query2), toJson(query));
});

test("m_language round-trip: prettyPrintQueryToM (Decimal.FromText + Binary.FromText)", () => {
  const script = `
let
  Source = Range.FromValues({
    {"Dec", "Bin"},
    {Decimal.FromText("123.450"), Binary.FromText("AQID", BinaryEncoding.Base64)}
  })
in
  Source
`;

  const query = compileMToQuery(script);
  const printed = prettyPrintQueryToM(query);
  assert.match(printed, /Decimal\.FromText\(\"123\.450\"\)/);
  assert.match(printed, /Binary\.FromText\(\"AQID\", BinaryEncoding\.Base64\)/);
  const query2 = compileMToQuery(printed);
  assert.deepEqual(toJson(query2), toJson(query));
});
