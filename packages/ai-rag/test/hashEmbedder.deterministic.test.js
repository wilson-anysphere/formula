import assert from "node:assert/strict";
import test from "node:test";

import { HashEmbedder } from "../src/embedding/hashEmbedder.js";

test("HashEmbedder is deterministic for identical inputs", async () => {
  const embedder = new HashEmbedder({ dimension: 64 });
  const [a, b] = await embedder.embedTexts(["hello world", "hello world"]);
  assert.equal(a.length, 64);
  assert.equal(b.length, 64);
  assert.deepEqual(Array.from(a), Array.from(b));
});

test("HashEmbedder produces approximately unit-normalized vectors for non-empty inputs", async () => {
  const embedder = new HashEmbedder({ dimension: 64 });
  const [vec] = await embedder.embedTexts(["hello world"]);
  const norm = Math.sqrt(Array.from(vec).reduce((acc, x) => acc + x * x, 0));
  assert.ok(norm > 0.999 && norm < 1.001, `expected ~unit norm, got ${norm}`);
});

