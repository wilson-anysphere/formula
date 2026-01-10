import test from "node:test";
import assert from "node:assert/strict";

import { excelWildcardTest } from "../index.js";

test("excel wildcards: * and ?", () => {
  assert.equal(excelWildcardTest("foobar", "foo"), true);
  assert.equal(excelWildcardTest("foobar", "foo", { matchEntireCell: true }), false);

  assert.equal(excelWildcardTest("foobar", "f*r"), true);
  assert.equal(excelWildcardTest("foobar", "f?o"), true);
  assert.equal(excelWildcardTest("foobar", "f??b"), true);
});

test("excel wildcards: escape with ~", () => {
  assert.equal(excelWildcardTest("*", "~*", { matchEntireCell: true }), true);
  assert.equal(excelWildcardTest("a?b", "a~?b", { matchEntireCell: true }), true);
  assert.equal(excelWildcardTest("a~b", "a~~b", { matchEntireCell: true }), true);
});

test("excel wildcards: case sensitivity", () => {
  assert.equal(excelWildcardTest("foo", "FOO"), true);
  assert.equal(excelWildcardTest("foo", "FOO", { matchCase: true }), false);
});

test("excel wildcards: match newlines", () => {
  assert.equal(excelWildcardTest("a\nb", "a*b", { matchEntireCell: true }), true);
});
