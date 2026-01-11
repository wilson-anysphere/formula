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

  const names = FUNCTION_CATALOG.functions.map((fn) => fn.name);
  assert.ok(names.length > 20, `Expected a non-trivial catalog, got ${names.length} functions`);

  for (const name of names) {
    assert.equal(name, name.toUpperCase(), `Expected function name to be uppercase, got ${name}`);
  }

  const sorted = [...names].sort();
  assert.deepEqual(names, sorted, "Expected function catalog to be sorted by name for deterministic diffs");
  assert.equal(new Set(names).size, names.length, "Expected function catalog to contain unique names");

  assert.ok(names.includes("XLOOKUP"), "Expected XLOOKUP to be present in the catalog");
});

