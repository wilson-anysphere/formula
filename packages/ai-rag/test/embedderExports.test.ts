import { readdir } from "node:fs/promises";
import { expect, test } from "vitest";

import * as aiRag from "../src/index.js";

test("ai-rag only exports the hash embedder", () => {
  const embedderExports = Object.keys(aiRag).filter((name) => name.toLowerCase().endsWith("embedder"));
  expect(embedderExports).toEqual(["HashEmbedder"]);
});

test("ai-rag only ships the hash embedder implementation", async () => {
  const entries = await readdir(new URL("../src/embedding/", import.meta.url));
  const normalized = entries.slice().sort();
  expect(normalized).toEqual(["hashEmbedder.d.ts", "hashEmbedder.js"]);
});
