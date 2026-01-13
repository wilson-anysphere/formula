import assert from "node:assert/strict";
import test from "node:test";

let sqlJsAvailable = true;
try {
  // Keep this as a computed dynamic import (no literal bare specifier) so
  // `scripts/run-node-tests.mjs` can still execute this file when `node_modules/`
  // is missing.
  const sqlJsModuleName = "sql" + ".js";
  await import(sqlJsModuleName);
} catch {
  sqlJsAvailable = false;
}

/** @type {Promise<any> | null} */
let sqliteVectorStorePromise = null;

async function getSqliteVectorStore() {
  if (!sqliteVectorStorePromise) {
    const modulePath = "../src/store/" + "sqliteVectorStore.js";
    sqliteVectorStorePromise = import(modulePath).then((mod) => mod.SqliteVectorStore);
  }
  return sqliteVectorStorePromise;
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
    const SqliteVectorStore = await getSqliteVectorStore();
    const store = await SqliteVectorStore.create({ storage, dimension: 3, autoSave: false, resetOnCorrupt: true });
    assert.equal(storage.removed, true, "expected SqliteVectorStore to clear persisted bytes on corruption");

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

test(
  "SqliteVectorStore resetOnCorrupt overwrites invalid persisted bytes when storage lacks remove()",
  { skip: !sqlJsAvailable },
  async () => {
    // Definitely not a SQLite database (header is wrong).
    const badBytes = new Uint8Array([0, 1, 2, 3, 4, 5, 6, 7, 8, 9]);

    /** @type {Uint8Array | null} */
    let stored = new Uint8Array(badBytes);
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

    const SqliteVectorStore = await getSqliteVectorStore();
    const store = await SqliteVectorStore.create({ storage, dimension: 3, autoSave: false, resetOnCorrupt: true });
    await store.close();

    assert.ok(saves >= 1, "expected SqliteVectorStore to overwrite corrupted bytes via save()");
    assert.ok(stored && stored.length >= 16, "expected persisted DB bytes to exist after reset");
    // SQLite database files begin with the 16-byte header: "SQLite format 3\\0".
    const expectedHeader = [
      83, 81, 76, 105, 116, 101, 32, 102, 111, 114, 109, 97, 116, 32, 51, 0,
    ];
    assert.deepEqual(Array.from(stored.slice(0, 16)), expectedHeader);
  }
);

test(
  "SqliteVectorStore resetOnCorrupt overwrites when storage.load throws (no remove)",
  { skip: !sqlJsAvailable },
  async () => {
    let throws = true;
    /** @type {Uint8Array | null} */
    let stored = null;
    let saves = 0;

    const storage = {
      async load() {
        if (throws) {
          throws = false;
          throw new Error("corrupt base64");
        }
        return stored ? new Uint8Array(stored) : null;
      },
      async save(data) {
        saves += 1;
        stored = new Uint8Array(data);
      },
    };

    const SqliteVectorStore = await getSqliteVectorStore();
    await SqliteVectorStore.create({ storage, dimension: 3, autoSave: false, resetOnCorrupt: true });
    assert.ok(saves >= 1, "expected SqliteVectorStore to overwrite corrupted payload via save()");
    assert.ok(stored && stored.length >= 16);
    const expectedHeader = [
      83, 81, 76, 105, 116, 101, 32, 102, 111, 114, 109, 97, 116, 32, 51, 0,
    ];
    assert.deepEqual(Array.from(stored.slice(0, 16)), expectedHeader);
  }
);

test(
  "SqliteVectorStore resetOnCorrupt=false throws on invalid persisted bytes",
  { skip: !sqlJsAvailable },
  async () => {
    // Not a real SQLite file header; should fail to open as an existing DB.
    const storage = new TestBinaryStorage(new TextEncoder().encode("not a sqlite database"));
    const SqliteVectorStore = await getSqliteVectorStore();
    await assert.rejects(SqliteVectorStore.create({ storage, dimension: 3, autoSave: false, resetOnCorrupt: false }));
    assert.equal(storage.removed, false);
  }
);
