import assert from "node:assert/strict";
import test from "node:test";

import { HashEmbedder } from "../src/rag.js";

/**
 * @param {number[]} a
 * @param {number[]} b
 */
function cosineSimilarity(a, b) {
  assert.equal(a.length, b.length);
  let dot = 0;
  for (let i = 0; i < a.length; i++) dot += a[i] * b[i];
  return dot;
}

test("ai-context HashEmbedder treats underscores like token separators", async () => {
  const embedder = new HashEmbedder({ dimension: 512 });
  const spaced = await embedder.embed("user id");
  const snake = await embedder.embed("user_id");
  const unrelated = await embedder.embed("unrelated");

  const simSnake = cosineSimilarity(spaced, snake);
  const simUnrelated = cosineSimilarity(spaced, unrelated);

  assert.ok(simSnake > simUnrelated);
  assert.ok(simSnake > 0.99, `expected high similarity for snake_case, got ${simSnake}`);
  assert.ok(simUnrelated < 0.8, `expected low similarity for unrelated token, got ${simUnrelated}`);
});

