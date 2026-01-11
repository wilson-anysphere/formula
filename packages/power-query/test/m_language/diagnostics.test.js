import assert from "node:assert/strict";
import test from "node:test";

import { parseM } from "../../src/m/parser.js";

test("m_language diagnostics: parse errors include location and suggestion", () => {
  assert.throws(
    () => parseM("Table.SelectColumns(Source,)"),
    (err) => {
      assert.equal(err.name, "MLanguageSyntaxError");
      assert.match(err.message, /line 1, column \d+/);
      assert.match(err.message, /Expected:/);
      assert.match(err.message, /Suggestion:/);
      return true;
    },
  );
});

