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

test("HashEmbedder is deterministic across instances", async () => {
  const a = new HashEmbedder({ dimension: 64 });
  const b = new HashEmbedder({ dimension: 64 });
  const [va] = await a.embedTexts(["hello world"]);
  const [vb] = await b.embedTexts(["hello world"]);
  assert.deepEqual(Array.from(va), Array.from(vb));
});

test("HashEmbedder produces approximately unit-normalized vectors for non-empty inputs", async () => {
  const embedder = new HashEmbedder({ dimension: 64 });
  const [vec] = await embedder.embedTexts(["hello world"]);
  const norm = Math.sqrt(Array.from(vec).reduce((acc, x) => acc + x * x, 0));
  assert.ok(norm > 0.999 && norm < 1.001, `expected ~unit norm, got ${norm}`);
});

test("HashEmbedder returns a stable zero vector for empty/whitespace input", async () => {
  const embedder = new HashEmbedder({ dimension: 64 });
  const [vec] = await embedder.embedTexts(["   "]);
  assert.equal(vec.length, 64);
  const norm = Math.sqrt(Array.from(vec).reduce((acc, x) => acc + x * x, 0));
  assert.equal(norm, 0);
  assert.ok(Array.from(vec).every((n) => Number.isFinite(n) && n === 0));
});

test("HashEmbedder respects AbortSignal", async () => {
  const embedder = new HashEmbedder({ dimension: 64 });
  const abortController = new AbortController();
  abortController.abort();

  await assert.rejects(embedder.embedTexts(["hello world"], { signal: abortController.signal }), {
    name: "AbortError",
  });
});
