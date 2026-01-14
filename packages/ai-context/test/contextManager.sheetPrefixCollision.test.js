import assert from "node:assert/strict";
import test from "node:test";

import { ContextManager } from "../src/contextManager.js";

test("ContextManager LRU eviction: deleteByPrefix does not collide for sheet names that share `-region-` prefixes", async () => {
  // Force eviction on the second indexed sheet.
  const cm = new ContextManager({ sheetIndexCacheLimit: 1 });

  const sheetPrefix = {
    name: "Sales",
    values: [
      ["Name", "Revenue"],
      ["South", 2],
    ],
  };
  const sheetWithPrefix = {
    name: "Sales-region-2024",
    values: [
      ["Name", "Revenue"],
      ["North", 1],
    ],
  };

  await cm.buildContext({ sheet: sheetPrefix, query: "South" });
  assert.ok(cm.ragIndex.store.size > 0);

  // Indexing the second sheet should evict the first (LRU limit=1) and delete only its chunks.
  await cm.buildContext({ sheet: sheetWithPrefix, query: "North" });

  const sheetNamesInStore = new Set(
    Array.from(cm.ragIndex.store.items.values()).map((item) => item.metadata.sheetName),
  );
  assert.deepEqual(sheetNamesInStore, new Set(["Sales-region-2024"]));
});

