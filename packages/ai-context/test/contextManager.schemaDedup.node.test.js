import assert from "node:assert/strict";
import test from "node:test";

import { ContextManager } from "../src/contextManager.js";
import { extractSheetSchema } from "../src/schema.js";

function instrumentIterationPasses(records) {
  let passes = 0;
  const proxy = new Proxy(records, {
    get(target, prop, receiver) {
      if (prop === Symbol.iterator) {
        return function () {
          passes += 1;
          return target[Symbol.iterator]();
        };
      }
      return Reflect.get(target, prop, receiver);
    },
  });
  return { proxy, getPasses: () => passes };
}

test("ContextManager.buildContext: extracts sheet schema only once per call (no duplicate extractSheetSchema pass)", async () => {
  const ragIndex = {
    store: { size: 0 },
    async indexSheet(sheet, options = {}) {
      const schema = extractSheetSchema(sheet, { signal: options.signal });
      return { schema, chunkCount: 0 };
    },
    async search() {
      return [];
    },
  };

  const cm = new ContextManager({
    tokenBudgetTokens: 1_000_000,
    redactor: (t) => t,
    ragIndex: /** @type {any} */ (ragIndex),
    // Avoid stableHashValue/schema signature computation; this test only cares about schema dedup.
    cacheSheetIndex: false,
  });

  // Include explicit table metadata so `extractSheetSchema` will iterate `sheet.tables`.
  const tables = [{ name: "Table1", range: "A1:B3" }];
  const { proxy: tablesProxy, getPasses } = instrumentIterationPasses(tables);

  const sheet = {
    name: "Sheet1",
    values: [
      ["Region", "Revenue"],
      ["North", 1000],
      ["South", 2000],
    ],
    tables: tablesProxy,
  };

  await cm.buildContext({ sheet, query: "revenue" });

  // `extractSheetSchema` iterates `sheet.tables` once per invocation. If `buildContext` regresses to
  // extracting schema separately from indexing, this pass count would increase to 2.
  assert.equal(getPasses(), 1, `expected 1 iteration pass over sheet.tables, got ${getPasses()}`);
});

