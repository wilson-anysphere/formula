import assert from "node:assert/strict";
import test from "node:test";

import { RagIndex } from "../src/rag.js";

test("RagIndex.indexSheet: deleteByPrefix does not collide for sheet names that share `-region-` prefixes", async () => {
  const rag = new RagIndex();

  const sheetWithPrefix = {
    name: "Sales-region-2024",
    values: [
      ["Name", "Revenue"],
      ["North", 1],
    ],
  };
  const sheetPrefix = {
    name: "Sales",
    values: [
      ["Name", "Revenue"],
      ["South", 2],
    ],
  };

  const indexedA = await rag.indexSheet(sheetWithPrefix);
  assert.ok(indexedA.chunkCount > 0);
  assert.equal(rag.store.size, indexedA.chunkCount);

  const indexedB = await rag.indexSheet(sheetPrefix);
  assert.ok(indexedB.chunkCount > 0);

  // Regression: previously indexing "Sales" would call `deleteByPrefix("Sales-region-")`,
  // which deleted the chunks for "Sales-region-2024" as well.
  const sheetNamesInStore = new Set(Array.from(rag.store.items.values()).map((item) => item.metadata.sheetName));
  assert.ok(sheetNamesInStore.has("Sales-region-2024"), "expected store to still contain chunks for Sales-region-2024");
  assert.ok(sheetNamesInStore.has("Sales"), "expected store to contain chunks for Sales");
  assert.equal(rag.store.size, indexedA.chunkCount + indexedB.chunkCount);
});

