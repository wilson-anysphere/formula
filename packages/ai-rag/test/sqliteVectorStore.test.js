import assert from "node:assert/strict";
import { mkdir, mkdtemp, rm, stat, writeFile } from "node:fs/promises";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { createSqliteFileVectorStore } from "../src/store/sqliteFileVectorStore.js";
import { InMemoryVectorStore } from "../src/store/inMemoryVectorStore.js";
import { SqliteVectorStore } from "../src/store/sqliteVectorStore.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

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

test(
  "SqliteVectorStore can reset persisted DB on dimension mismatch (file storage)",
  { skip: !sqlJsAvailable },
  async () => {
    const tmpRoot = path.join(__dirname, ".tmp");
    await mkdir(tmpRoot, { recursive: true });
    const tmpDir = await mkdtemp(path.join(tmpRoot, "sqlite-store-dim-mismatch-"));
    const filePath = path.join(tmpDir, "vectors.sqlite");

    try {
      const store1 = await createSqliteFileVectorStore({
        filePath,
        dimension: 3,
        autoSave: true,
      });
      await store1.upsert([{ id: "a", vector: [1, 0, 0], metadata: { workbookId: "wb" } }]);
      await store1.close();

      const store2 = await createSqliteFileVectorStore({
        filePath,
        dimension: 4,
        autoSave: true,
        resetOnDimensionMismatch: true,
      });

      const list = await store2.list();
      assert.deepEqual(list, []);

      await store2.upsert([{ id: "c", vector: [1, 0, 0, 0], metadata: { workbookId: "wb" } }]);
      const hits = await store2.query([1, 0, 0, 0], 1, { workbookId: "wb" });
      assert.equal(hits[0].id, "c");
      await store2.close();
    } finally {
      await rm(tmpDir, { recursive: true, force: true });
    }
  }
);

test(
  "SqliteVectorStore throws typed error on dimension mismatch when reset is disabled (file storage)",
  { skip: !sqlJsAvailable },
  async () => {
    const tmpRoot = path.join(__dirname, ".tmp");
    await mkdir(tmpRoot, { recursive: true });
    const tmpDir = await mkdtemp(path.join(tmpRoot, "sqlite-store-dim-mismatch-throw-"));
    const filePath = path.join(tmpDir, "vectors.sqlite");

    try {
      const store1 = await createSqliteFileVectorStore({
        filePath,
        dimension: 3,
        autoSave: true,
      });
      await store1.upsert([{ id: "a", vector: [1, 0, 0], metadata: { workbookId: "wb" } }]);
      await store1.close();

      await assert.rejects(
        createSqliteFileVectorStore({
          filePath,
          dimension: 4,
          autoSave: false,
          resetOnDimensionMismatch: false,
        }),
        (err) => {
          assert.equal(err?.name, "SqliteVectorStoreDimensionMismatchError");
          assert.equal(err?.dbDimension, 3);
          assert.equal(err?.requestedDimension, 4);
          return true;
        }
      );
    } finally {
      await rm(tmpDir, { recursive: true, force: true });
    }
  }
);

test("SqliteVectorStore.vacuum() VACUUMs and persists a smaller DB (even with autoSave:false)", { skip: !sqlJsAvailable }, async () => {
  const tmpRoot = path.join(__dirname, ".tmp");
  await mkdir(tmpRoot, { recursive: true });
  const tmpDir = await mkdtemp(path.join(tmpRoot, "sqlite-store-compact-"));
  const filePath = path.join(tmpDir, "vectors.sqlite");

  try {
    // Create a large DB snapshot.
    const store1 = await createSqliteFileVectorStore({ filePath, dimension: 3, autoSave: false });

    const payload = "x".repeat(4096);
    const total = 300;
    const records = Array.from({ length: total }, (_, i) => ({
      id: `rec-${i}`,
      vector: i % 2 === 0 ? [1, 0, 0] : [0, 1, 0],
      metadata: { workbookId: "wb", i, payload },
    }));

    await store1.upsert(records);
    // Persist the initial snapshot. (`autoSave:false` skips persisting on upsert.)
    await store1.close();

    // Reopen, delete most records, and then vacuum to reclaim space.
    const store2 = await createSqliteFileVectorStore({ filePath, dimension: 3, autoSave: false });
    const remaining = 10;
    const deleteIds = Array.from({ length: total - remaining }, (_, i) => `rec-${i}`);
    await store2.delete(deleteIds);

    const before = (await stat(filePath)).size;

    await store2.vacuum();

    const after = (await stat(filePath)).size;
    assert.ok(after > 0);
    assert.ok(after < before, `Expected compacted DB to shrink: before=${before}, after=${after}`);

    // Ensure store remains usable after VACUUM (custom dot() function still registered).
    const remainingId = `rec-${total - 1}`;
    const rec = await store2.get(remainingId);
    assert.ok(rec);
    assert.equal(rec.metadata.i, total - 1);

    const hits = await store2.query([1, 0, 0], 5, { workbookId: "wb" });
    assert.ok(hits.length > 0);

    await store2.close();

    const store3 = await createSqliteFileVectorStore({ filePath, dimension: 3, autoSave: false });
    const rec2 = await store3.get(remainingId);
    assert.ok(rec2);
    assert.equal(rec2.metadata.i, total - 1);
    const hits2 = await store3.query([1, 0, 0], 5, { workbookId: "wb" });
    assert.ok(hits2.length > 0);
    await store3.close();
  } finally {
    await rm(tmpDir, { recursive: true, force: true });
  }
});

test(
  "SqliteVectorStore.compact() waits for in-flight persistence before running VACUUM",
  { skip: !sqlJsAvailable },
  async () => {
    class BlockingStorage {
      saveCalls = 0;
      #data = null;
      #startedResolve = null;
      started = new Promise((resolve) => {
        this.#startedResolve = resolve;
      });
      release = () => {};

      async load() {
        return this.#data ? new Uint8Array(this.#data) : null;
      }

      async save(data) {
        this.saveCalls += 1;
        this.#data = new Uint8Array(data);
        if (this.saveCalls === 1) {
          this.#startedResolve?.();
          await new Promise((resolve) => {
            this.release = resolve;
          });
        }
      }
    }

    const storage = new BlockingStorage();
    const store = await SqliteVectorStore.create({ storage, dimension: 3, autoSave: true });

    let sawVacuum = false;
    const originalRun = store._db.run.bind(store._db);
    store._db.run = (sql) => {
      if (typeof sql === "string" && sql.trim().toUpperCase() === "VACUUM;") {
        sawVacuum = true;
      }
      return originalRun(sql);
    };

    const upsertPromise = store.upsert([{ id: "a", vector: [1, 0, 0], metadata: { workbookId: "wb" } }]);

    // Wait until the first persist (triggered by upsert) reaches storage.save and blocks.
    await storage.started;

    const compactPromise = store.compact();

    // Let compact enqueue/work; it should not run VACUUM while the save is blocked.
    await Promise.resolve();
    assert.equal(sawVacuum, false);

    // Unblock the save so compaction can proceed.
    storage.release();

    await Promise.all([upsertPromise, compactPromise]);
    assert.equal(sawVacuum, true);
    await store.close();
  }
);

test(
  "SqliteVectorStore migrates v1 schema to v2 (structured metadata columns)",
  { skip: !sqlJsAvailable },
  async () => {
    const tmpRoot = path.join(__dirname, ".tmp");
    await mkdir(tmpRoot, { recursive: true });
    const tmpDir = await mkdtemp(path.join(tmpRoot, "sqlite-store-migrate-"));
    const filePath = path.join(tmpDir, "vectors.sqlite");

    try {
      // Create a v1 database (no schema_version, vectors table only has metadata_json).
      function locateSqlJsFile(file, prefix = "") {
        try {
          if (typeof import.meta.resolve === "function") {
            const resolved = import.meta.resolve(`sql.js/dist/${file}`);
            if (resolved) {
              if (resolved.startsWith("file://")) {
                let pathname = decodeURIComponent(new URL(resolved).pathname);
                if (/^\/[A-Za-z]:\//.test(pathname)) pathname = pathname.slice(1);
                return pathname;
              }
              return resolved;
            }
          }
        } catch {
          // ignore
        }
        return prefix ? `${prefix}${file}` : file;
      }

      const sqlMod = await import("sql.js");
      const initSqlJs = sqlMod.default ?? sqlMod;
      const SQL = await initSqlJs({ locateFile: locateSqlJsFile });
      const db = new SQL.Database();

      db.run(`
        CREATE TABLE IF NOT EXISTS vector_store_meta (
          key TEXT PRIMARY KEY,
          value TEXT NOT NULL
        );

        INSERT INTO vector_store_meta (key, value) VALUES ('dimension', '3');

        CREATE TABLE IF NOT EXISTS vectors (
          id TEXT PRIMARY KEY,
          workbook_id TEXT,
          vector BLOB NOT NULL,
          metadata_json TEXT NOT NULL
        );

        CREATE INDEX IF NOT EXISTS idx_vectors_workbook_id ON vectors(workbook_id);
      `);

      function float32ToBlob(vec) {
        const v = vec instanceof Float32Array ? vec : Float32Array.from(vec);
        return new Uint8Array(v.buffer, v.byteOffset, v.byteLength);
      }

      function normalizeL2(vec) {
        const v = vec instanceof Float32Array ? new Float32Array(vec) : Float32Array.from(vec);
        let sum = 0;
        for (let i = 0; i < v.length; i += 1) sum += v[i] * v[i];
        const len = Math.sqrt(sum);
        if (len > 0) {
          for (let i = 0; i < v.length; i += 1) v[i] /= len;
        }
        return v;
      }

      const insert = db.prepare(
        "INSERT INTO vectors (id, workbook_id, vector, metadata_json) VALUES (?, ?, ?, ?);"
      );
      insert.run([
        "a",
        "wb",
        float32ToBlob(normalizeL2([1, 0, 0])),
        JSON.stringify({
          workbookId: "wb",
          sheetName: "Sheet1",
          kind: "table",
          title: "T1",
          rect: { r0: 0, c0: 0, r1: 1, c1: 1 },
          text: "hello",
          contentHash: "hash-a",
          metadataHash: "meta-a",
          tokenCount: 5,
          label: "A",
        }),
      ]);
      insert.run([
        "b",
        "wb",
        float32ToBlob(normalizeL2([0, 1, 0])),
        JSON.stringify({
          workbookId: "wb",
          sheetName: "Sheet1",
          kind: "table",
          title: "T2",
          rect: { r0: 2, c0: 2, r1: 3, c1: 3 },
          text: "world",
          contentHash: "hash-b",
          tokenCount: 5,
          label: "B",
        }),
      ]);
      insert.free();

      const data = db.export();
      db.close();
      await writeFile(filePath, data);

      // Reopen with the new store implementation; it should migrate on open.
      const store = await createSqliteFileVectorStore({ filePath, dimension: 3, autoSave: false });

      const rec = await store.get("a");
      assert.ok(rec);
      assert.equal(rec.metadata.workbookId, "wb");
      assert.equal(rec.metadata.sheetName, "Sheet1");
      assert.equal(rec.metadata.kind, "table");
      assert.equal(rec.metadata.title, "T1");
      assert.deepEqual(rec.metadata.rect, { r0: 0, c0: 0, r1: 1, c1: 1 });
      assert.equal(rec.metadata.text, "hello");
      assert.equal(rec.metadata.contentHash, "hash-a");
      assert.equal(rec.metadata.metadataHash, "meta-a");
      assert.equal(rec.metadata.tokenCount, 5);
      assert.equal(rec.metadata.label, "A");

      const hits = await store.query([1, 0, 0], 1, { workbookId: "wb" });
      assert.equal(hits[0].id, "a");

      // metadata_json should now only contain the extra keys.
      const stmt = store._db.prepare(
        "SELECT sheet_name, content_hash, metadata_hash, metadata_json FROM vectors WHERE id = ?;"
      );
      stmt.bind(["a"]);
      assert.ok(stmt.step());
      const migratedRow = stmt.get();
      stmt.free();
      assert.equal(migratedRow[0], "Sheet1");
      assert.equal(migratedRow[1], "hash-a");
      assert.equal(migratedRow[2], "meta-a");
      assert.deepEqual(JSON.parse(migratedRow[3]), { label: "A" });

      await store.close();

      // Ensure schema_version was persisted.
      const store2 = await createSqliteFileVectorStore({ filePath, dimension: 3, autoSave: false });
      const metaStmt = store2._db.prepare(
        "SELECT value FROM vector_store_meta WHERE key = 'schema_version' LIMIT 1;"
      );
      assert.ok(metaStmt.step());
      assert.equal(metaStmt.get()[0], "2");
      metaStmt.free();
      await store2.close();
    } finally {
      await rm(tmpDir, { recursive: true, force: true });
    }
  }
);

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

test(
  "SqliteVectorStore ensures covering index for workbook hash scans exists",
  { skip: !sqlJsAvailable },
  async () => {
    const store = await SqliteVectorStore.create({ dimension: 3, autoSave: false });
    try {
      const stmt = store._db.prepare("PRAGMA index_list(vectors);");
      /** @type {string[]} */
      const names = [];
      while (stmt.step()) {
        const row = stmt.get();
        names.push(String(row[1]));
      }
      stmt.free();
      assert.ok(names.includes("idx_vectors_workbook_hashes"));
    } finally {
      await store.close();
    }
  }
);

test(
  "SqliteVectorStore.listContentHashes does not parse metadata_json",
  { skip: !sqlJsAvailable },
  async () => {
    const store = await SqliteVectorStore.create({ dimension: 3, autoSave: false });
    try {
      const blob = new Uint8Array(new Float32Array([1, 0, 0]).buffer);
      const stmt = store._db.prepare(
        "INSERT INTO vectors (id, workbook_id, vector, content_hash, metadata_hash, metadata_json) VALUES (?, ?, ?, ?, ?, ?);"
      );
      stmt.run(["bad-json", "wb", blob, "ch", "mh", "not-json"]);
      stmt.free();

      const rows = await store.listContentHashes({ workbookId: "wb" });
      rows.sort((a, b) => a.id.localeCompare(b.id));
      assert.deepEqual(rows, [{ id: "bad-json", contentHash: "ch", metadataHash: "mh" }]);
    } finally {
      await store.close();
    }
  }
);

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

test("SqliteVectorStore.query throws on query vector dimension mismatch", { skip: !sqlJsAvailable }, async () => {
  const store = await SqliteVectorStore.create({ dimension: 3, autoSave: false });
  try {
    await assert.rejects(store.query([1, 0], 1), /expected 3/);
  } finally {
    await store.close();
  }
});

test("SqliteVectorStore.get throws when stored vector blob has wrong length", { skip: !sqlJsAvailable }, async () => {
  const store = await SqliteVectorStore.create({ dimension: 3, autoSave: false });
  try {
    const badBlob = new Uint8Array(new Float32Array([1, 0]).buffer);
    const stmt = store._db.prepare("INSERT INTO vectors (id, workbook_id, vector, metadata_json) VALUES (?, ?, ?, ?);");
    stmt.run(["bad", null, badBlob, "{}"]);
    stmt.free();

    await assert.rejects(store.get("bad"), /expected 3/);
  } finally {
    await store.close();
  }
});

test("SqliteVectorStore.list throws when stored vector blob has wrong length", { skip: !sqlJsAvailable }, async () => {
  const store = await SqliteVectorStore.create({ dimension: 3, autoSave: false });
  try {
    const badBlob = new Uint8Array(new Float32Array([1, 0]).buffer);
    const stmt = store._db.prepare("INSERT INTO vectors (id, workbook_id, vector, metadata_json) VALUES (?, ?, ?, ?);");
    stmt.run(["bad", null, badBlob, "{}"]);
    stmt.free();

    await assert.rejects(store.list(), /expected 3/);
    await assert.rejects(store.list({ includeVector: false }), /expected 12/);
  } finally {
    await store.close();
  }
});

test("SqliteVectorStore.query throws when stored vector blob has wrong length (dot() validation)", { skip: !sqlJsAvailable }, async () => {
  const store = await SqliteVectorStore.create({ dimension: 3, autoSave: false });
  try {
    const badBlob = new Uint8Array(new Float32Array([1, 0]).buffer);
    const stmt = store._db.prepare("INSERT INTO vectors (id, workbook_id, vector, metadata_json) VALUES (?, ?, ?, ?);");
    stmt.run(["bad", null, badBlob, "{}"]);
    stmt.free();

    await assert.rejects(store.query([1, 0, 0], 1), /expected 3/);
  } finally {
    await store.close();
  }
});

test(
  "SqliteVectorStore throws when stored vector blob byte length is not a multiple of 4",
  { skip: !sqlJsAvailable },
  async () => {
    const store = await SqliteVectorStore.create({ dimension: 3, autoSave: false });
    try {
      const badBlob = new Uint8Array([1, 2, 3, 4, 5]);
      const stmt = store._db.prepare(
        "INSERT INTO vectors (id, workbook_id, vector, metadata_json) VALUES (?, ?, ?, ?);"
      );
      stmt.run(["bad-bytes", null, badBlob, "{}"]);
      stmt.free();

      await assert.rejects(store.get("bad-bytes"), /failed to decode stored vector blob: Invalid vector blob length/);
      await assert.rejects(store.list(), /failed to decode stored vector blob for id=.*Invalid vector blob length/);
      await assert.rejects(store.list({ includeVector: false }), /invalid vector blob length/i);
      await assert.rejects(store.listContentHashes(), /invalid vector blob length/i);
      await assert.rejects(store.query([1, 0, 0], 1), /dot\\(\\) failed to decode arg0 vector blob: Invalid vector blob length/);
    } finally {
      await store.close();
    }
  }
);

test(
  "SqliteVectorStore.listContentHashes throws when stored vector blob has wrong length",
  { skip: !sqlJsAvailable },
  async () => {
    const store = await SqliteVectorStore.create({ dimension: 3, autoSave: false });
    try {
      const badBlob = new Uint8Array(new Float32Array([1, 0]).buffer);
      const stmt = store._db.prepare(
        "INSERT INTO vectors (id, workbook_id, vector, metadata_json) VALUES (?, ?, ?, ?);"
      );
      stmt.run(["bad", null, badBlob, "{}"]);
      stmt.free();

      await assert.rejects(store.listContentHashes(), /expected 12/);
    } finally {
      await store.close();
    }
  }
);

test(
  "SqliteVectorStore dot() throws when blobs have matching but non-store dimension",
  { skip: !sqlJsAvailable },
  async () => {
    const store = await SqliteVectorStore.create({ dimension: 3, autoSave: false });
    try {
      const badVec = new Uint8Array(new Float32Array([1, 0]).buffer);
      const stmt = store._db.prepare("SELECT dot(?, ?) AS score;");
      try {
        stmt.bind([badVec, badVec]);
        await assert.rejects(
          async () => {
            stmt.step();
          },
          /expected 3/
        );
      } finally {
        stmt.free();
      }
    } finally {
      await store.close();
    }
  }
);

test("SqliteVectorStore dot() handles unaligned Uint8Array blobs (copies into aligned buffer)", { skip: !sqlJsAvailable }, async () => {
  const store = await SqliteVectorStore.create({ dimension: 3, autoSave: false });
  try {
    // Create a Uint8Array view with an unaligned byteOffset.
    const raw = new Uint8Array(new Float32Array([1, 0, 0]).buffer);
    const buf = new ArrayBuffer(raw.byteLength + 1);
    const unaligned = new Uint8Array(buf, 1, raw.byteLength);
    unaligned.set(raw);

    const stmt = store._db.prepare("SELECT dot(?, ?) AS score;");
    try {
      stmt.bind([unaligned, unaligned]);
      assert.ok(stmt.step());
      const [score] = stmt.get();
      assert.equal(Number(score), 1);
    } finally {
      stmt.free();
    }
  } finally {
    await store.close();
  }
});

test("SqliteVectorStore dot() throws when blob dimensions mismatch (arg lengths differ)", { skip: !sqlJsAvailable }, async () => {
  const store = await SqliteVectorStore.create({ dimension: 3, autoSave: false });
  try {
    const vec3 = new Uint8Array(new Float32Array([1, 0, 0]).buffer);
    const vec2 = new Uint8Array(new Float32Array([1, 0]).buffer);
    const stmt = store._db.prepare("SELECT dot(?, ?) AS score;");
    try {
      stmt.bind([vec3, vec2]);
      await assert.rejects(
        async () => {
          stmt.step();
        },
        /got 3 vs 2/
      );
    } finally {
      stmt.free();
    }
  } finally {
    await store.close();
  }
});
