import assert from "node:assert/strict";
import { mkdtemp, mkdir, readFile, rm, stat, utimes, writeFile, readdir } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import test from "node:test";

import { CacheManager } from "../../src/cache/cache.js";
import { MemoryCacheStore } from "../../src/cache/memory.js";
import { QueryEngine } from "../../src/engine.js";

/**
 * @param {string} root
 * @param {{ recursive?: boolean }} [options]
 */
async function listDir(root, options = {}) {
  const recursive = options.recursive ?? false;
  /** @type {Array<{ path: string; name: string; size: number; mtimeMs: number; isDir: false }>} */
  const out = [];

  const visit = async (dir) => {
    const entries = await readdir(dir, { withFileTypes: true });
    for (const entry of entries) {
      const fullPath = path.join(dir, entry.name);
      if (entry.isDirectory()) {
        if (recursive) await visit(fullPath);
        continue;
      }
      if (!entry.isFile()) continue;
      const info = await stat(fullPath);
      out.push({ path: fullPath, name: entry.name, size: info.size, mtimeMs: info.mtimeMs, isDir: false });
    }
  };

  await visit(root);
  return out;
}

test("Folder.Files: recursive listing returns expected metadata", async () => {
  const root = await mkdtemp(path.join(os.tmpdir(), "pq-folder-"));
  try {
    await mkdir(path.join(root, "sub"), { recursive: true });
    await writeFile(path.join(root, "a.csv"), "A,B\n1,2\n");
    await writeFile(path.join(root, "sub", "b.json"), JSON.stringify({ ok: true }));

    const engine = new QueryEngine({
      cache: new CacheManager({ store: new MemoryCacheStore() }),
      fileAdapter: {
        readText: async (p) => String(await readFile(p, "utf8")),
        readBinary: async (p) => new Uint8Array(await readFile(p)),
        listDir,
        stat: async (p) => {
          const info = await stat(p);
          return { mtimeMs: info.mtimeMs, size: info.size };
        },
      },
    });

    const query = {
      id: "q_folder_recursive",
      name: "Folder",
      source: { type: "folder", path: root, options: { recursive: true } },
      steps: [],
    };

    const { table } = await engine.executeQueryWithMeta(query, {}, { cache: { validation: "none" } });
    assert.equal(table.rowCount, 2);
    assert.deepEqual(table.columns.map((c) => c.name), ["Name", "Extension", "Folder Path", "Date modified", "Size"]);

    // Easier: just inspect the grid.
    const grid = table.toGrid({ includeHeader: true });
    assert.equal(grid[0][0], "Name");
    const returnedNames = grid.slice(1).map((r) => r[0]).sort();
    assert.deepEqual(returnedNames, ["a.csv", "b.json"]);

    const extensions = grid.slice(1).map((r) => r[1]).sort();
    assert.deepEqual(extensions, [".csv", ".json"]);
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});

test("Folder.Files: includeContent reads bytes", async () => {
  const root = await mkdtemp(path.join(os.tmpdir(), "pq-folder-content-"));
  try {
    await writeFile(path.join(root, "blob.bin"), new Uint8Array([1, 2, 3]));

    const engine = new QueryEngine({
      cache: new CacheManager({ store: new MemoryCacheStore() }),
      fileAdapter: {
        readBinary: async (p) => new Uint8Array(await readFile(p)),
        listDir,
      },
    });

    const query = {
      id: "q_folder_content",
      name: "Folder Content",
      source: { type: "folder", path: root, options: { recursive: true, includeContent: true } },
      steps: [],
    };

    const table = await engine.executeQuery(query, {}, { cache: { validation: "none" } });
    assert.deepEqual(table.columns.map((c) => c.name), ["Name", "Extension", "Folder Path", "Date modified", "Size", "Content"]);
    const content = table.getCell(0, table.getColumnIndex("Content"));
    assert.ok(content instanceof Uint8Array);
    assert.deepEqual(Array.from(content), [1, 2, 3]);
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});

test("Folder.Files cache key changes when a file is modified", async () => {
  const root = await mkdtemp(path.join(os.tmpdir(), "pq-folder-cache-"));
  try {
    const filePath = path.join(root, "data.txt");
    await writeFile(filePath, "one");

    const engine = new QueryEngine({
      cache: new CacheManager({ store: new MemoryCacheStore() }),
      fileAdapter: {
        readBinary: async (p) => new Uint8Array(await readFile(p)),
        listDir,
      },
    });

    const query = {
      id: "q_folder_cache",
      name: "Folder Cache",
      source: { type: "folder", path: root, options: { recursive: true } },
      steps: [],
    };

    const key1 = await engine.getCacheKey(query, {}, {});

    await writeFile(filePath, "two-two-two");
    const now = Date.now() + 2000;
    await utimes(filePath, now / 1000, now / 1000);

    const key2 = await engine.getCacheKey(query, {}, {});
    assert.ok(key1 && key2);
    assert.notEqual(key1, key2);
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});
