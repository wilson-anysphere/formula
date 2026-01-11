import assert from "node:assert/strict";
import test from "node:test";

import { FunctionRegistry } from "../src/functionRegistry.js";

test("FunctionRegistry loads the Rust function catalog (HLOOKUP is present)", () => {
  const registry = new FunctionRegistry();
  assert.ok(registry.getFunction("XLOOKUP"), "Expected XLOOKUP to be present");
  assert.ok(registry.getFunction("_xlfn.XLOOKUP"), "Expected _xlfn.XLOOKUP alias to be present");
  assert.ok(registry.isRangeArg("_xlfn.XLOOKUP", 1), "Expected _xlfn.XLOOKUP arg2 to be a range");
  assert.equal(registry.getFunction("SUM")?.minArgs, 0, "Expected SUM minArgs to come from catalog");
  assert.equal(registry.getFunction("SUM")?.maxArgs, 255, "Expected SUM maxArgs to come from catalog");
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
