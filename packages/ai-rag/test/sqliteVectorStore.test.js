import assert from "node:assert/strict";
import { mkdir, mkdtemp, rm } from "node:fs/promises";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { createSqliteFileVectorStore } from "../src/store/sqliteFileVectorStore.js";

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
