import assert from "node:assert/strict";
import test from "node:test";

import { FunctionRegistry } from "../src/functionRegistry.js";

test("FunctionRegistry loads the Rust function catalog (HLOOKUP is present)", () => {
  const registry = new FunctionRegistry();
  assert.ok(registry.getFunction("XLOOKUP"), "Expected XLOOKUP to be present");
  assert.ok(
    registry.getFunction("HLOOKUP"),
    `Expected HLOOKUP to be present, got: ${registry.list().map(f => f.name).join(", ")}`
  );
});

test("FunctionRegistry falls back to curated defaults when catalog is missing/invalid", () => {
  const missingCatalog = new FunctionRegistry(undefined, { catalog: null });
  assert.ok(missingCatalog.getFunction("SUM"), "Expected SUM to exist in fallback registry");
  assert.equal(
    missingCatalog.getFunction("HLOOKUP"),
    undefined,
    "Expected catalog-only functions to be absent when catalog is missing"
  );

  const invalidCatalog = new FunctionRegistry(undefined, { catalog: { functions: [{ nope: true }] } });
  assert.ok(invalidCatalog.getFunction("SUM"), "Expected SUM to exist in fallback registry");
  assert.equal(
    invalidCatalog.getFunction("HLOOKUP"),
    undefined,
    "Expected catalog-only functions to be absent when catalog is invalid"
  );
});
