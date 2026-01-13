import assert from "node:assert/strict";
import test from "node:test";

import { JsonVectorStore } from "../src/store/jsonVectorStore.js";

class TestBinaryStorage {
  /**
   * @param {Uint8Array | null} data
   */
  constructor(data) {
    /** @type {Uint8Array | null} */
    this.data = data ? new Uint8Array(data) : null;
    this.removed = false;
  }

  async load() {
    return this.data ? new Uint8Array(this.data) : null;
  }

  /**
   * @param {Uint8Array} data
   */
  async save(data) {
    this.data = new Uint8Array(data);
  }

  async remove() {
    this.removed = true;
    this.data = null;
  }
}

test("JsonVectorStore resetOnCorrupt clears storage and loads empty on invalid JSON", async () => {
  const storage = new TestBinaryStorage(new TextEncoder().encode("{not json"));
  const store = new JsonVectorStore({ storage, dimension: 3, resetOnCorrupt: true });

  const loaded = await store.list();
  assert.deepEqual(loaded, []);
  assert.equal(storage.removed, true);
  assert.equal(storage.data, null);

  // Store should still be usable after reset.
  await store.upsert([{ id: "a", vector: [1, 0, 0], metadata: { label: "A" } }]);
  const rec = await store.get("a");
  assert.ok(rec);
  assert.equal(rec.metadata.label, "A");
  assert.ok(storage.data, "expected store to persist after upsert");
});

test("JsonVectorStore resetOnCorrupt overwrites corrupted payloads when storage lacks remove()", async () => {
  const bad = new TextEncoder().encode("{not json");
  /** @type {Uint8Array | null} */
  let stored = new Uint8Array(bad);
  let saves = 0;

  const storage = {
    async load() {
      return stored ? new Uint8Array(stored) : null;
    },
    async save(data) {
      saves += 1;
      stored = new Uint8Array(data);
    },
  };

  const store = new JsonVectorStore({ storage, dimension: 3, resetOnCorrupt: true, autoSave: false });
  const list1 = await store.list();
  assert.deepEqual(list1, []);
  assert.ok(saves >= 1, "expected JsonVectorStore to overwrite corrupted payload with empty snapshot");

  // A second instance should now be able to parse the persisted data.
  const store2 = new JsonVectorStore({ storage, dimension: 3, resetOnCorrupt: true, autoSave: false });
  const list2 = await store2.list();
  assert.deepEqual(list2, []);
});

test("JsonVectorStore resetOnCorrupt=false leaves corrupted payload in place", async () => {
  const bad = new TextEncoder().encode("{not json");
  const storage = new TestBinaryStorage(bad);
  const store = new JsonVectorStore({ storage, dimension: 3, resetOnCorrupt: false });

  const loaded = await store.list();
  assert.deepEqual(loaded, []);

  assert.equal(storage.removed, false);
  assert.ok(storage.data, "expected corrupted bytes to remain when resetOnCorrupt=false");
  assert.deepEqual(Array.from(storage.data ?? []), Array.from(bad));
});
