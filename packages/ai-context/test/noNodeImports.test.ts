import { readFile } from "node:fs/promises";

import { expect, test } from "vitest";

const browserSafeEntrypoints = [
  new URL("../src/index.js", import.meta.url),
  new URL("../src/contextManager.js", import.meta.url),
  new URL("../src/schema.js", import.meta.url),
  new URL("../src/rag.js", import.meta.url),
];

test("browser-safe entrypoints do not contain static node:* imports", async () => {
  for (const url of browserSafeEntrypoints) {
    const code = await readFile(url, "utf8");
    expect(code, `${url} should not statically import node:*`).not.toMatch(/from\s+["']node:/);
    expect(code, `${url} should not dynamically import node:*`).not.toMatch(/import\(\s*["']node:/);
  }
});

