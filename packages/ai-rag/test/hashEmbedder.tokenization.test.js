import assert from "node:assert/strict";
import test from "node:test";

import { HashEmbedder } from "../src/embedding/hashEmbedder.js";
import { cosineSimilarity } from "../src/store/vectorMath.js";

test('HashEmbedder tokenization: "RevenueByRegion" matches "revenue by region"', async () => {
  const embedder = new HashEmbedder({ dimension: 256 });

  const [queryVec, pascalVec, unrelatedVec] = await embedder.embedTexts([
    "revenue by region",
    "RevenueByRegion",
    "employee salary table",
  ]);

  const simQueryPascal = cosineSimilarity(queryVec, pascalVec);
  const simQueryUnrelated = cosineSimilarity(queryVec, unrelatedVec);

  // Tokenization should split PascalCase identifiers, so these should be very similar.
  assert.ok(
    simQueryPascal > 0.8,
    `expected high similarity between "revenue by region" and "RevenueByRegion", got ${simQueryPascal}`,
  );
  assert.ok(
    simQueryPascal > simQueryUnrelated + 0.2,
    `expected RevenueByRegion similarity (${simQueryPascal}) to exceed unrelated similarity (${simQueryUnrelated}) by a margin`,
  );
});

test("HashEmbedder treats underscores like token separators", async () => {
  const embedder = new HashEmbedder({ dimension: 512 });
  const [spaced, snake, unrelated] = await embedder.embedTexts(["user id", "user_id", "unrelated"]);

  const simSnake = cosineSimilarity(spaced, snake);
  const simUnrelated = cosineSimilarity(spaced, unrelated);

  // "user id" and "user_id" should tokenize the same, so their vectors should be
  // nearly identical and closer than an unrelated token.
  assert.ok(simSnake > simUnrelated);
  assert.ok(simSnake > 0.99, `expected high similarity for snake_case, got ${simSnake}`);
  assert.ok(simUnrelated < 0.8, `expected low similarity for unrelated token, got ${simUnrelated}`);
});

test("HashEmbedder tokenization changes remain deterministic", async () => {
  const embedder = new HashEmbedder({ dimension: 128 });
  const [a, b] = await embedder.embedTexts(["RevenueByRegion2024", "RevenueByRegion2024"]);
  assert.deepEqual(Array.from(a), Array.from(b));

  const embedder2 = new HashEmbedder({ dimension: 128 });
  const [c] = await embedder2.embedTexts(["RevenueByRegion2024"]);
  assert.deepEqual(Array.from(a), Array.from(c));
});
