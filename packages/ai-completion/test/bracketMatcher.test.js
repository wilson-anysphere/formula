import assert from "node:assert/strict";
import test from "node:test";

import { findMatchingBracketEnd } from "../src/bracketMatcher.js";

test("findMatchingBracketEnd treats '[' inside external workbook prefixes as a literal character", () => {
  // Workbook name contains a literal `[`, which does *not* introduce nesting in workbook prefixes.
  // (The name is: `A1[Name.xlsx`.)
  const src = "[A1[Name.xlsx]Sheet1!A1";
  const end = findMatchingBracketEnd(src, 0, src.length);
  assert.equal(end, src.indexOf("]Sheet1") + 1);
});

test("findMatchingBracketEnd handles escaped closing brackets inside external workbook prefixes", () => {
  // Workbook name is literally: `Book[Name[Part[More]Name.xlsx` (contains 3 `[` chars and one `]`),
  // so Excel encodes the `]` as `]]`. The prefix still ends at the closing `]` before `Sheet1!`.
  const src = "[Book[Name[Part[More]]Name.xlsx]Sheet1!A1";
  const end = findMatchingBracketEnd(src, 0, src.length);
  assert.equal(end, src.indexOf("]Sheet1") + 1);
});

test("findMatchingBracketEnd does not misclassify incomplete structured references as workbook prefixes", () => {
  // `[[A]B` has a `[` that would require a second closing bracket in structured-ref syntax, but
  // the prefix ends before that. The matcher should return null rather than matching the internal
  // `]` as if it were a workbook prefix.
  const src = "[[A]B";
  const end = findMatchingBracketEnd(src, 0, src.length);
  assert.equal(end, null);
});
