import { expect, test } from "vitest";

import * as aiRag from "../src/index.js";

test("ai-rag only exports the hash embedder", () => {
  const embedderExports = Object.keys(aiRag).filter((name) => name.toLowerCase().endsWith("embedder"));
  expect(embedderExports).toEqual(["HashEmbedder"]);
});
