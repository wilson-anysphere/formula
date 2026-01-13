import assert from "node:assert/strict";
import test from "node:test";

import { InMemoryBinaryStorage } from "../src/store/binaryStorage.js";
import { JsonVectorStore } from "../src/store/jsonVectorStore.js";

test("JsonVectorStore persists vectors in v2 format and can reload them", async () => {
  const storage = new InMemoryBinaryStorage();

  const store1 = new JsonVectorStore({ storage, dimension: 3 });
  await store1.upsert([
    { id: "a", vector: [1, 0, 0], metadata: { label: "A" } },
    { id: "b", vector: [0, 1, 0], metadata: { label: "B" } },
  ]);

  const persisted = await storage.load();
  assert.ok(persisted);
  const json = new TextDecoder().decode(persisted);
  const parsed = JSON.parse(json);
  assert.equal(parsed.version, 2);
  assert.equal(parsed.dimension, 3);
  assert.ok(Array.isArray(parsed.records));
  assert.equal(typeof parsed.records[0].vector_b64, "string");
  assert.ok(!("vector" in parsed.records[0]));

  const store2 = new JsonVectorStore({ storage, dimension: 3 });
  const rec = await store2.get("a");
  assert.ok(rec);
  assert.equal(rec.metadata.label, "A");
  assert.ok(rec.vector instanceof Float32Array);

  const hits = await store2.query([1, 0, 0], 1);
  assert.equal(hits[0].id, "a");
});

test("JsonVectorStore can load v1 persisted payloads", async () => {
  const storage = new InMemoryBinaryStorage();
  const payloadV1 = JSON.stringify({
    version: 1,
    dimension: 3,
    records: [
      { id: "a", vector: [1, 0, 0], metadata: { label: "A" } },
      { id: "b", vector: [0, 1, 0], metadata: { label: "B" } },
    ],
  });
  await storage.save(new TextEncoder().encode(payloadV1));

  const store = new JsonVectorStore({ storage, dimension: 3 });
  const rec = await store.get("b");
  assert.ok(rec);
  assert.equal(rec.metadata.label, "B");

  const hits = await store.query([0, 1, 0], 1);
  assert.equal(hits[0].id, "b");
});

test("JsonVectorStore v2 payloads are typically smaller than v1 payloads", async () => {
  const storage = new InMemoryBinaryStorage();

  const dimension = 153;
  const store = new JsonVectorStore({ storage, dimension });
  await store.upsert([{ id: "a", vector: new Array(dimension).fill(0.1), metadata: { label: "A" } }]);

  const v2 = await storage.load();
  assert.ok(v2);
  const v2Payload = new TextDecoder().decode(v2);

  const records = await store.list();
  const v1Payload = JSON.stringify({
    version: 1,
    dimension,
    records: records.map((r) => ({ id: r.id, vector: Array.from(r.vector), metadata: r.metadata })),
  });

  assert.ok(v2Payload.length < v1Payload.length, `expected v2 < v1, got ${v2Payload.length} >= ${v1Payload.length}`);
});

