import { readFile } from "node:fs/promises";

import { expect, test } from "vitest";

const browserSafeEntrypoints = [
  new URL("../src/index.js", import.meta.url),
  new URL("../src/embedding/hashEmbedder.js", import.meta.url),
  new URL("../src/store/binaryStorage.js", import.meta.url),
  new URL("../src/store/inMemoryVectorStore.js", import.meta.url),
  new URL("../src/store/jsonVectorStore.js", import.meta.url),
  new URL("../src/store/sqliteVectorStore.js", import.meta.url),
  new URL("../src/pipeline/indexWorkbook.js", import.meta.url),
  new URL("../src/workbook/chunkWorkbook.js", import.meta.url),
  new URL("../src/workbook/chunkToText.js", import.meta.url),
  new URL("../src/workbook/fromSpreadsheetApi.js", import.meta.url),
  new URL("../src/workbook/rect.js", import.meta.url),
  new URL("../src/utils/abort.js", import.meta.url),
  new URL("../src/utils/hash.js", import.meta.url),
  new URL("../src/retrieval/searchWorkbookRag.js", import.meta.url),
  new URL("../src/retrieval/ranking.js", import.meta.url),
  new URL("../src/retrieval/rankResults.js", import.meta.url),
  new URL("../../../apps/desktop/src/ai/rag/index.js", import.meta.url),
];

test("browser-safe entrypoints do not contain static node:* imports", async () => {
  for (const url of browserSafeEntrypoints) {
    const code = await readFile(url, "utf8");
    expect(code, `${url} should not statically import node:*`).not.toMatch(/from\s+["']node:/);
    expect(code, `${url} should not statically import node:*`).not.toMatch(/import\(\s*["']node:/);
  }
});
