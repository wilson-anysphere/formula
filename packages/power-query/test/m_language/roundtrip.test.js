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
