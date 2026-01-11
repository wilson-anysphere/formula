import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";

import FUNCTION_CATALOG from "../../../shared/functionCatalog.mjs";

test("function catalog artifact matches committed JSON and is sorted", async () => {
  const jsonPath = fileURLToPath(new URL("../../../shared/functionCatalog.json", import.meta.url));
  const raw = await readFile(jsonPath, "utf8");
  const parsed = JSON.parse(raw);

  assert.deepEqual(parsed, FUNCTION_CATALOG, "Expected functionCatalog.mjs to match functionCatalog.json");

  const allowedTypes = new Set(["any", "number", "text", "bool"]);
  const allowedVolatility = new Set(["non_volatile", "volatile"]);

  const names = FUNCTION_CATALOG.functions.map((fn) => fn.name);
  assert.ok(names.length > 20, `Expected a non-trivial catalog, got ${names.length} functions`);

  for (const fn of FUNCTION_CATALOG.functions) {
    assert.equal(fn.name, fn.name.toUpperCase(), `Expected function name to be uppercase, got ${fn.name}`);
    assert.ok(Number.isInteger(fn.min_args) && fn.min_args >= 0, `Expected min_args to be >= 0 for ${fn.name}`);
    assert.ok(Number.isInteger(fn.max_args) && fn.max_args >= fn.min_args, `Expected max_args >= min_args for ${fn.name}`);
    assert.ok(allowedVolatility.has(fn.volatility), `Expected volatility to be valid for ${fn.name}`);
    assert.ok(allowedTypes.has(fn.return_type), `Expected return_type to be valid for ${fn.name}`);
    assert.ok(Array.isArray(fn.arg_types), `Expected arg_types to be an array for ${fn.name}`);
    for (const t of fn.arg_types) {
      assert.ok(allowedTypes.has(t), `Expected arg_types entry to be valid for ${fn.name}: ${t}`);
    }
    if (fn.max_args === 0) {
      assert.equal(fn.arg_types.length, 0, `Expected zero-arg functions to have empty arg_types for ${fn.name}`);
    }
  }

  const sorted = [...names].sort();
  assert.deepEqual(names, sorted, "Expected function catalog to be sorted by name for deterministic diffs");
  assert.equal(new Set(names).size, names.length, "Expected function catalog to contain unique names");

  assert.ok(names.includes("XLOOKUP"), "Expected XLOOKUP to be present in the catalog");
});
