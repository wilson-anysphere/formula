import assert from "node:assert/strict";
import test from "node:test";

import { HashEmbedder } from "../src/embedding/hashEmbedder.js";

test("HashEmbedder token hash cache is deterministic and reused across calls", async () => {
  const embedder = new HashEmbedder({ dimension: 64, cacheSize: 10_000 });

  // @ts-ignore - `_debug` is a test-only internal.
  const debug = embedder._debug;

  const [a] = await embedder.embedTexts(["hello world"]);
  const hitsAfterFirst = debug?.tokenCacheHits ?? 0;
  const missesAfterFirst = debug?.tokenCacheMisses ?? 0;

  const [b] = await embedder.embedTexts(["hello world"]);

  assert.deepEqual(Array.from(a), Array.from(b));

  // Optional validation: when test debug counters are enabled, the second embed
  // should hit the token cache.
  if (debug) {
    assert.equal(hitsAfterFirst, 0);
    assert.ok(missesAfterFirst > 0);
    assert.ok(debug.tokenCacheHits > hitsAfterFirst);
    assert.equal(debug.tokenCacheMisses, missesAfterFirst);
  }
});

