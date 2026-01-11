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

