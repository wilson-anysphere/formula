import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

import { stripComments } from "../../../apps/desktop/test/sourceTextUtils.js";

const browserSafeEntrypoints = [
  new URL("../src/index.js", import.meta.url),
  new URL("../src/contextManager.js", import.meta.url),
  new URL("../src/schema.js", import.meta.url),
  new URL("../src/rag.js", import.meta.url),
  new URL("../src/workbookSchema.js", import.meta.url),
  new URL("../src/summarizeSheet.js", import.meta.url),
  new URL("../src/summarizeWorkbookSchema.js", import.meta.url),
  new URL("../src/queryAware.js", import.meta.url),
  new URL("../src/sampling.js", import.meta.url),
  new URL("../src/tokenBudget.js", import.meta.url),
  new URL("../src/trimMessagesToBudget.js", import.meta.url),
  new URL("../src/budgetPlanner.js", import.meta.url),
  new URL("../src/dlp.js", import.meta.url),
  new URL("../src/abort.js", import.meta.url),
  new URL("../src/tsv.js", import.meta.url),
];

test("ai-context: browser-safe entrypoints do not contain static node:* imports", async () => {
  for (const url of browserSafeEntrypoints) {
    const code = stripComments(await readFile(url, "utf8"));
    assert.ok(!/from\s+["']node:/.test(code), `${url} should not statically import node:*`);
    assert.ok(!/import\(\s*["']node:/.test(code), `${url} should not dynamically import node:*`);
  }
});
