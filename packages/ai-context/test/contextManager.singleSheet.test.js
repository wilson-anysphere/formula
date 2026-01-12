import assert from "node:assert/strict";
import test from "node:test";

import { ContextManager } from "../src/contextManager.js";

function makeSheet(values, name = "Sheet1") {
  return { name, values };
}

test("buildContext: repeated calls do not duplicate RAG chunks for the same sheet", async () => {
  const cm = new ContextManager({ tokenBudgetTokens: 1000 });
  const sheet = makeSheet([
    ["Region", "Revenue"],
    ["North", 1000],
    ["South", 2000],
  ]);

  const out1 = await cm.buildContext({ sheet, query: "revenue by region" });
  const size1 = cm.ragIndex.store.size;

  const out2 = await cm.buildContext({ sheet, query: "revenue by region" });
  const size2 = cm.ragIndex.store.size;

  assert.equal(size1, 1);
  assert.equal(size2, size1);
  assert.equal(out1.retrieved[0].range, "Sheet1!A1:B3");
  assert.equal(out2.retrieved[0].range, out1.retrieved[0].range);
});

test("buildContext: stale RAG chunks are removed when a sheet shrinks", async () => {
  const cm = new ContextManager({ tokenBudgetTokens: 1000 });

  const twoRegions = makeSheet([
    ["Region", "Revenue", "", "Cost"],
    ["North", 1000, "", 50],
    ["South", 2000, "", 60],
  ]);
  await cm.buildContext({ sheet: twoRegions, query: "revenue" });
  assert.equal(cm.ragIndex.store.size, 2);

  const oneRegion = makeSheet([
    ["Region", "Revenue"],
    ["North", 1000],
    ["South", 2000],
  ]);
  await cm.buildContext({ sheet: oneRegion, query: "revenue" });
  assert.equal(cm.ragIndex.store.size, 1);
});

test("buildContext: retrieved context is preserved under tight token budgets", async () => {
  const cm = new ContextManager({ tokenBudgetTokens: 20 });
  const sheet = makeSheet([
    ["Region", "Revenue"],
    ["North", 1000],
    ["South", 2000],
  ]);

  const out = await cm.buildContext({ sheet, query: "revenue by region" });
  assert.match(out.promptContext, /^## retrieved\b/m);
});

test("buildContext: caps matrix size to avoid Excel-scale allocations", async () => {
  const cm = new ContextManager({ tokenBudgetTokens: 1_000 });

  // 1,000 rows x 300 cols => 300,000 cells (> 200,000 cap). The ContextManager should
  // truncate columns so downstream schema + sampling work remains bounded.
  const values = Array.from({ length: 1_000 }, (_v, r) => {
    const row = Array.from({ length: 300 }, () => null);
    row[0] = r;
    return row;
  });
  const sheet = makeSheet(values);

  const out = await cm.buildContext({ sheet, query: "col1", sampleRows: 1 });
  assert.equal(out.sampledRows.length, 1);
  assert.equal(out.sampledRows[0].length, 200);
});
