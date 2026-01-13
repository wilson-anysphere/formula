import assert from "node:assert/strict";
import test from "node:test";

import { FunctionRegistry } from "../src/functionRegistry.js";

test("FunctionRegistry loads the Rust function catalog (HLOOKUP is present)", () => {
  const registry = new FunctionRegistry();
  assert.ok(registry.getFunction("SEQUENCE"), "Expected SEQUENCE (catalog-only) to be present");
  assert.ok(registry.getFunction("XLOOKUP"), "Expected XLOOKUP to be present");
  assert.ok(registry.getFunction("_xlfn.XLOOKUP"), "Expected _xlfn.XLOOKUP alias to be present");
  assert.ok(registry.isRangeArg("_xlfn.XLOOKUP", 1), "Expected _xlfn.XLOOKUP arg2 to be a range");
  assert.equal(registry.getFunction("SUM")?.minArgs, 0, "Expected SUM minArgs to come from catalog");
  assert.equal(registry.getFunction("SUM")?.maxArgs, 255, "Expected SUM maxArgs to come from catalog");
  assert.equal(registry.getArgType("PV", 0), "number", "Expected PV arg1 type to come from catalog arg_types");
  assert.ok(
    registry.getFunction("HLOOKUP"),
    `Expected HLOOKUP to be present, got: ${registry.list().map(f => f.name).join(", ")}`
  );
});

test("FunctionRegistry falls back to curated defaults when catalog is missing/invalid", () => {
  const missingCatalog = new FunctionRegistry(undefined, { catalog: null });
  assert.ok(missingCatalog.getFunction("SUM"), "Expected SUM to exist in fallback registry");
  assert.equal(
    missingCatalog.getFunction("SEQUENCE"),
    undefined,
    "Expected catalog-only functions to be absent when catalog is missing"
  );

  const invalidCatalog = new FunctionRegistry(undefined, { catalog: { functions: [{ nope: true }] } });
  assert.ok(invalidCatalog.getFunction("SUM"), "Expected SUM to exist in fallback registry");
  assert.equal(
    invalidCatalog.getFunction("SEQUENCE"),
    undefined,
    "Expected catalog-only functions to be absent when catalog is invalid"
  );
});

test("FunctionRegistry uses curated range metadata for common multi-range functions", () => {
  const registry = new FunctionRegistry();

  // SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)
  assert.ok(registry.isRangeArg("SUMIFS", 0), "Expected SUMIFS arg1 to be a range");
  assert.ok(registry.isRangeArg("SUMIFS", 1), "Expected SUMIFS arg2 to be a range");
  assert.equal(registry.isRangeArg("SUMIFS", 2), false, "Expected SUMIFS arg3 to be a value");
  assert.ok(registry.isRangeArg("SUMIFS", 3), "Expected SUMIFS arg4 (criteria_range2) to be a range");
  assert.equal(registry.isRangeArg("SUMIFS", 4), false, "Expected SUMIFS arg5 (criteria2) to be a value");

  // _xlfn aliases should preserve the curated signatures.
  assert.ok(registry.isRangeArg("_xlfn.SUMIFS", 0), "Expected _xlfn.SUMIFS arg1 to be a range");
  assert.ok(registry.isRangeArg("_xlfn.FILTER", 0), "Expected _xlfn.FILTER arg1 to be a range");
  assert.ok(registry.isRangeArg("_xlfn.FILTER", 1), "Expected _xlfn.FILTER arg2 to be a range");

  // TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...)
  assert.equal(registry.isRangeArg("TEXTJOIN", 0), false, "Expected TEXTJOIN delimiter not to be a range");
  assert.equal(registry.isRangeArg("TEXTJOIN", 1), false, "Expected TEXTJOIN ignore_empty not to be a range");
  assert.ok(registry.isRangeArg("TEXTJOIN", 2), "Expected TEXTJOIN text1 to be a range");
  assert.ok(registry.isRangeArg("TEXTJOIN", 3), "Expected TEXTJOIN text2 to be a range (varargs)");
});
