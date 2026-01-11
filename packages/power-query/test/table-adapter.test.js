import assert from "node:assert/strict";
import test from "node:test";

import { CacheManager } from "../src/cache/cache.js";
import { MemoryCacheStore } from "../src/cache/memory.js";
import { QueryEngine } from "../src/engine.js";
import { DataTable } from "../src/table.js";

test("QueryEngine resolves table sources via tableAdapter when context.tables is missing", async () => {
  let tableCalls = 0;

  const engine = new QueryEngine({
    cache: new CacheManager({ store: new MemoryCacheStore() }),
    tableAdapter: {
      getTable: async (tableName) => {
        tableCalls += 1;
        assert.equal(tableName, "Sales");
        return DataTable.fromGrid(
          [
            ["Region", "Sales"],
            ["East", 100],
            ["West", 200],
          ],
          { hasHeaders: true, inferTypes: true },
        );
      },
    },
  });

  const query = {
    id: "q_table",
    name: "Sales Table",
    source: { type: "table", table: "Sales" },
    steps: [],
  };

  // Without a table signature, caching should be bypassed even when a cache manager exists.
  assert.equal(await engine.getCacheKey(query, {}, {}), null);
  assert.equal(tableCalls, 0, "expected cache key computation to not load tables");

  const table = await engine.executeQuery(query, {}, {});
  assert.deepEqual(table.toGrid(), [
    ["Region", "Sales"],
    ["East", 100],
    ["West", 200],
  ]);
  assert.equal(tableCalls, 1);
});

