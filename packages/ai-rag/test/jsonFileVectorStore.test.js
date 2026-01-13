import assert from "node:assert/strict";
import { mkdir, mkdtemp, rm } from "node:fs/promises";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { JsonFileVectorStore } from "../src/store/jsonFileVectorStore.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("JsonFileVectorStore persists vectors and can reload them", async () => {
  const tmpRoot = path.join(__dirname, ".tmp");
  await mkdir(tmpRoot, { recursive: true });
  const tmpDir = await mkdtemp(path.join(tmpRoot, "json-store-"));
  const filePath = path.join(tmpDir, "vectors.json");

  try {
    const store1 = new JsonFileVectorStore({ filePath, dimension: 3 });
    await store1.upsert([
      { id: "a", vector: [1, 0, 0], metadata: { label: "A" } },
      { id: "b", vector: [0, 1, 0], metadata: { label: "B" } },
    ]);

    // New instance loads from disk.
    const store2 = new JsonFileVectorStore({ filePath, dimension: 3 });
    const rec = await store2.get("a");
    assert.ok(rec);
    assert.equal(rec.metadata.label, "A");

    const hits = await store2.query([1, 0, 0], 1);
    assert.equal(hits[0].id, "a");
  } finally {
    await rm(tmpDir, { recursive: true, force: true });
  }
});

test("JsonFileVectorStore.query throws on query vector dimension mismatch", async () => {
  const tmpRoot = path.join(__dirname, ".tmp");
  await mkdir(tmpRoot, { recursive: true });
  const tmpDir = await mkdtemp(path.join(tmpRoot, "json-store-dim-"));
  const filePath = path.join(tmpDir, "vectors.json");

  try {
    const store = new JsonFileVectorStore({ filePath, dimension: 3 });
    await assert.rejects(store.query([1, 0], 1), /expected 3/);
  } finally {
    await rm(tmpDir, { recursive: true, force: true });
  }
});
