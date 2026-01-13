import assert from "node:assert/strict";
import test from "node:test";

import { SqliteVectorStore } from "../src/store/sqliteVectorStore.js";

let sqlJsAvailable = true;
try {
  await import("sql.js");
} catch {
  sqlJsAvailable = false;
}

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

test(
  "SqliteVectorStore resetOnCorrupt ignores invalid persisted bytes",
  { skip: !sqlJsAvailable },
  async () => {
    // Not a real SQLite file header; should fail to open as an existing DB.
    const storage = new TestBinaryStorage(new TextEncoder().encode("not a sqlite database"));
    const store = await SqliteVectorStore.create({ storage, dimension: 3, autoSave: false, resetOnCorrupt: true });

    await store.upsert([
      { id: "a", vector: [1, 0, 0], metadata: { workbookId: "wb", label: "A" } },
      { id: "b", vector: [0, 1, 0], metadata: { workbookId: "wb", label: "B" } },
    ]);

    const rec = await store.get("a");
    assert.ok(rec);
    assert.equal(rec.metadata.label, "A");

    const hits = await store.query([1, 0, 0], 1, { workbookId: "wb" });
    assert.equal(hits[0].id, "a");
    await store.close();
  }
);

