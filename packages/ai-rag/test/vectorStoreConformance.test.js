import assert from "node:assert/strict";
import test from "node:test";

import { InMemoryBinaryStorage } from "../src/store/binaryStorage.js";
import { InMemoryVectorStore } from "../src/store/inMemoryVectorStore.js";
import { JsonVectorStore } from "../src/store/jsonVectorStore.js";
import { SqliteVectorStore } from "../src/store/sqliteVectorStore.js";

let sqlJsAvailable = true;
try {
  await import("sql.js");
} catch {
  sqlJsAvailable = false;
}

/**
 * @param {number} actual
 * @param {number} expected
 * @param {number} [eps]
 */
function assertApprox(actual, expected, eps = 1e-5) {
  assert.ok(
    Math.abs(actual - expected) <= eps,
    `Expected ${actual} to be within ${eps} of ${expected}`
  );
}

/**
 * Build a unit-length 3D vector with x component = `x` (and positive y).
 * @param {number} x
 */
function unitVec(x) {
  const y2 = 1 - x * x;
  // Guard tiny negative values from float rounding.
  const y = Math.sqrt(Math.max(0, y2));
  return [x, y, 0];
}

/**
 * @param {string} name
 * @param {() => Promise<any>} createStore
 * @param {{ skip?: boolean }} [opts]
 */
function defineVectorStoreConformanceSuite(name, createStore, opts) {
  test(`VectorStore conformance: ${name}`, { skip: opts?.skip }, async (t) => {
    const store = await createStore();

    try {
      await store.upsert([
        { id: "norm", vector: [3, 4, 0], metadata: { workbookId: "wb-norm", label: "Norm" } },

        { id: "list-a", vector: [1, 0, 0], metadata: { workbookId: "wb-list-1", label: "A" } },
        { id: "list-b", vector: [0, 1, 0], metadata: { workbookId: "wb-list-1", label: "B" } },
        { id: "list-c", vector: [0, 0, 1], metadata: { workbookId: "wb-list-2", label: "C" } },

        { id: "inc", vector: [1, 0, 0], metadata: { workbookId: "wb-include", label: "Inc" } },

        { id: "q-best", vector: [1, 0, 0], metadata: { workbookId: "wb-query", label: "best" } },
        { id: "q-0.8", vector: unitVec(0.8), metadata: { workbookId: "wb-query", label: "0.8" } },
        { id: "q-0.6", vector: unitVec(0.6), metadata: { workbookId: "wb-query", label: "0.6" } },
        { id: "q-0.0", vector: [0, 1, 0], metadata: { workbookId: "wb-query", label: "0.0" } },
        { id: "q-other-best", vector: [1, 0, 0], metadata: { workbookId: "wb-query-other", label: "other" } },

        { id: "drop1", vector: unitVec(1), metadata: { workbookId: "wb-filter", tag: "drop" } },
        { id: "drop2", vector: unitVec(0.9), metadata: { workbookId: "wb-filter", tag: "drop" } },
        { id: "keep1", vector: unitVec(0.8), metadata: { workbookId: "wb-filter", tag: "keep" } },
        { id: "keep2", vector: unitVec(0.7), metadata: { workbookId: "wb-filter", tag: "keep" } },
        { id: "keep3", vector: unitVec(0.6), metadata: { workbookId: "wb-filter", tag: "keep" } },
      ]);

      await t.test("upsert/get returns normalized vectors", async () => {
        const rec = await store.get("norm");
        assert.ok(rec);
        assert.equal(rec.metadata.workbookId, "wb-norm");
        assert.equal(rec.metadata.label, "Norm");

        assert.ok(rec.vector instanceof Float32Array);
        assert.equal(rec.vector.length, 3);

        assertApprox(Math.hypot(rec.vector[0], rec.vector[1], rec.vector[2]), 1);
        assertApprox(rec.vector[0], 0.6);
        assertApprox(rec.vector[1], 0.8);
        assertApprox(rec.vector[2], 0);
      });

      await t.test("list({ workbookId }) only returns that workbook", async () => {
        const res = await store.list({ workbookId: "wb-list-1" });
        assert.equal(res.length, 2);
        assert.deepEqual(
          res
            .map((r) => r.id)
            .slice()
            .sort(),
          ["list-a", "list-b"]
        );
        for (const r of res) assert.equal(r.metadata.workbookId, "wb-list-1");
      });

      await t.test("list({ includeVector:false }) elides vectors but keeps metadata", async () => {
        const res = await store.list({ workbookId: "wb-include", includeVector: false });
        assert.equal(res.length, 1);
        assert.equal(res[0].id, "inc");
        assert.equal(res[0].vector, undefined);
        assert.equal(res[0].metadata.workbookId, "wb-include");
        assert.equal(res[0].metadata.label, "Inc");
      });

      await t.test("query(vector, topK, { workbookId }) returns topK ordered desc", async () => {
        const hits = await store.query([1, 0, 0], 3, { workbookId: "wb-query" });
        assert.equal(hits.length, 3);
        assert.deepEqual(
          hits.map((h) => h.id),
          ["q-best", "q-0.8", "q-0.6"]
        );
        assert.ok(hits[0].score >= hits[1].score);
        assert.ok(hits[1].score >= hits[2].score);
      });

      await t.test("query with filter returns up to topK matching results", async () => {
        const hits = await store.query([1, 0, 0], 2, {
          workbookId: "wb-filter",
          filter: (metadata) => metadata.tag === "keep",
        });
        assert.equal(hits.length, 2);
        assert.deepEqual(
          hits.map((h) => h.id),
          ["keep1", "keep2"]
        );
        assert.ok(hits[0].score >= hits[1].score);
      });

      await t.test("AbortSignal: already-aborted signals reject for list/query", async () => {
        const abortController = new AbortController();
        abortController.abort();

        await assert.rejects(store.list({ workbookId: "wb-query", signal: abortController.signal }), {
          name: "AbortError",
        });

        await assert.rejects(store.query([1, 0, 0], 1, { workbookId: "wb-query", signal: abortController.signal }), {
          name: "AbortError",
        });
      });
    } finally {
      await store.close?.();
    }
  });
}

defineVectorStoreConformanceSuite("InMemoryVectorStore", async () => new InMemoryVectorStore({ dimension: 3 }));

defineVectorStoreConformanceSuite(
  "JsonVectorStore",
  async () => new JsonVectorStore({ dimension: 3, autoSave: false, storage: new InMemoryBinaryStorage() })
);

defineVectorStoreConformanceSuite(
  "SqliteVectorStore",
  async () => await SqliteVectorStore.create({ dimension: 3, autoSave: false, storage: new InMemoryBinaryStorage() }),
  { skip: !sqlJsAvailable }
);

