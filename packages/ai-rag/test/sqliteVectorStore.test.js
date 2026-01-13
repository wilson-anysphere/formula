import assert from "node:assert/strict";
import { mkdir, mkdtemp, rm } from "node:fs/promises";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { createSqliteFileVectorStore } from "../src/store/sqliteFileVectorStore.js";
import { InMemoryVectorStore } from "../src/store/inMemoryVectorStore.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

let sqlJsAvailable = true;
try {
  await import("sql.js");
} catch {
  sqlJsAvailable = false;
}

test("SqliteVectorStore persists vectors and can reload them", { skip: !sqlJsAvailable }, async () => {
  const tmpRoot = path.join(__dirname, ".tmp");
  await mkdir(tmpRoot, { recursive: true });
  const tmpDir = await mkdtemp(path.join(tmpRoot, "sqlite-store-"));
  const filePath = path.join(tmpDir, "vectors.sqlite");

  try {
    const store1 = await createSqliteFileVectorStore({ filePath, dimension: 3, autoSave: true });
    await store1.upsert([
      { id: "a", vector: [1, 0, 0], metadata: { workbookId: "wb", label: "A" } },
      { id: "b", vector: [0, 1, 0], metadata: { workbookId: "wb", label: "B" } },
    ]);
    await store1.close();

    const store2 = await createSqliteFileVectorStore({ filePath, dimension: 3, autoSave: false });
    const rec = await store2.get("a");
    assert.ok(rec);
    assert.equal(rec.metadata.label, "A");

    const hits = await store2.query([1, 0, 0], 1, { workbookId: "wb" });
    assert.equal(hits[0].id, "a");
    await store2.close();
  } finally {
    await rm(tmpDir, { recursive: true, force: true });
  }
});

test("SqliteVectorStore.list respects AbortSignal", { skip: !sqlJsAvailable }, async () => {
  const tmpRoot = path.join(__dirname, ".tmp");
  await mkdir(tmpRoot, { recursive: true });
  const tmpDir = await mkdtemp(path.join(tmpRoot, "sqlite-store-abort-"));
  const filePath = path.join(tmpDir, "vectors.sqlite");

  try {
    const store = await createSqliteFileVectorStore({ filePath, dimension: 3, autoSave: false });

    const abortController = new AbortController();
    abortController.abort();

    await assert.rejects(store.list({ signal: abortController.signal }), { name: "AbortError" });
    await store.close();
  } finally {
    await rm(tmpDir, { recursive: true, force: true });
  }
});

test("SqliteVectorStore.query returns topK matching results after filtering", { skip: !sqlJsAvailable }, async () => {
  const tmpRoot = path.join(__dirname, ".tmp");
  await mkdir(tmpRoot, { recursive: true });
  const tmpDir = await mkdtemp(path.join(tmpRoot, "sqlite-store-filter-"));
  const filePath = path.join(tmpDir, "vectors.sqlite");

  try {
    const sqliteStore = await createSqliteFileVectorStore({ filePath, dimension: 3, autoSave: false });
    const memoryStore = new InMemoryVectorStore({ dimension: 3 });

    /** @type {{ id: string, vector: number[], metadata: any }[]} */
    const records = [];
    // Add many "high scoring" vectors that will be excluded by the filter.
    for (let i = 0; i < 200; i += 1) {
      records.push({
        id: `x${i}`,
        vector: [1, 0, 0],
        metadata: { workbookId: "wb", ok: false, i },
      });
    }
    // Add enough lower scoring vectors that should pass the filter. Use distinct
    // scores so ordering is deterministic across stores.
    for (let i = 1; i <= 30; i += 1) {
      records.push({
        id: `a${i}`,
        vector: [1, i, 0],
        metadata: { workbookId: "wb", ok: true, i },
      });
    }

    await sqliteStore.upsert(records);
    await memoryStore.upsert(records);

    const topK = 10;
    const filter = (metadata) => metadata.ok === true;
    const sqliteHits = await sqliteStore.query([1, 0, 0], topK, { workbookId: "wb", filter });
    const memoryHits = await memoryStore.query([1, 0, 0], topK, { workbookId: "wb", filter });

    assert.equal(sqliteHits.length, topK);
    assert.equal(memoryHits.length, topK);
    assert.deepEqual(
      sqliteHits.map((h) => h.id),
      memoryHits.map((h) => h.id)
    );
    for (const hit of sqliteHits) {
      assert.equal(hit.metadata.ok, true);
    }

    await sqliteStore.close();
  } finally {
    await rm(tmpDir, { recursive: true, force: true });
  }
});

test("SqliteVectorStore.query respects AbortSignal", { skip: !sqlJsAvailable }, async () => {
  const tmpRoot = path.join(__dirname, ".tmp");
  await mkdir(tmpRoot, { recursive: true });
  const tmpDir = await mkdtemp(path.join(tmpRoot, "sqlite-store-query-abort-"));
  const filePath = path.join(tmpDir, "vectors.sqlite");

  try {
    const store = await createSqliteFileVectorStore({ filePath, dimension: 3, autoSave: false });

    const abortController = new AbortController();
    abortController.abort();

    await assert.rejects(store.query([1, 0, 0], 1, { signal: abortController.signal }), { name: "AbortError" });
    await store.close();
  } finally {
    await rm(tmpDir, { recursive: true, force: true });
  }
});
