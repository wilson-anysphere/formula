import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { createHash, randomUUID } from "node:crypto";
import { promises as fs } from "node:fs";

import { indexedDB, IDBKeyRange } from "fake-indexeddb";
import * as Y from "yjs";

import { createCollabSession } from "../src/index.ts";

globalThis.indexedDB = indexedDB;
globalThis.IDBKeyRange = IDBKeyRange;

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function encodeRecord(update) {
  const header = Buffer.allocUnsafe(4);
  header.writeUInt32BE(update.byteLength, 0);
  return Buffer.concat([header, Buffer.from(update)]);
}

test("CollabSession legacy `options.offline` (file) migrates old .yjslog format and clear removes it", async (t) => {
  const dir = await mkdtemp(path.join(tmpdir(), "collab-session-offline-compat-file-migrate-"));
  const legacyFilePath = path.join(dir, "doc.yjslog");

  t.after(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  // Create a legacy offline log file containing a single snapshot update.
  {
    const seed = createCollabSession({ schema: { autoInit: false } });
    await seed.setCellValue("Sheet1:0:0", "from-legacy");
    const update = Y.encodeStateAsUpdate(seed.doc);
    await fs.writeFile(legacyFilePath, encodeRecord(update));
    seed.destroy();
    seed.doc.destroy();
  }

  // New session should load from the legacy file via migration.
  {
    const session = createCollabSession({
      schema: { autoInit: false },
      offline: { mode: "file", filePath: legacyFilePath },
    });

    await session.offline?.whenLoaded();
    assert.equal((await session.getCell("Sheet1:0:0"))?.value, "from-legacy");

    await session.offline?.clear();

    session.destroy();
    session.doc.destroy();
  }

  // After clear(), the legacy file should be removed and the migrated state should not reappear.
  {
    const session = createCollabSession({
      schema: { autoInit: false },
      offline: { mode: "file", filePath: legacyFilePath },
    });

    await session.offline?.whenLoaded();
    assert.equal(await session.getCell("Sheet1:0:0"), null);

    session.destroy();
    session.doc.destroy();
  }
});

test("CollabSession legacy `options.offline` with autoLoad=false gates schema init until load", async (t) => {
  const dir = await mkdtemp(path.join(tmpdir(), "collab-session-offline-compat-file-schema-"));
  const legacyFilePath = path.join(dir, "doc.yjslog");

  t.after(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  // Seed a legacy offline log with a workbook that already has a non-default sheet id.
  {
    const doc = new Y.Doc();
    const sheets = doc.getArray("sheets");
    const sheet = new Y.Map();
    sheet.set("id", "Persisted");
    sheet.set("name", "Persisted");
    sheets.push([sheet]);
    const update = Y.encodeStateAsUpdate(doc);
    await fs.writeFile(legacyFilePath, encodeRecord(update));
    doc.destroy();
  }

  const session = createCollabSession({
    doc: new Y.Doc(),
    // Do not auto-load persistence; we want to ensure schema init waits.
    offline: { mode: "file", filePath: legacyFilePath, autoLoad: false },
  });

  t.after(() => {
    session.destroy();
    session.doc.destroy();
  });

  // No default sheet should be created before offline state is loaded.
  assert.equal(session.sheets.length, 0);

  await session.offline?.whenLoaded();

  const ids = session.sheets.toArray().map((s) => String(s.get("id") ?? ""));
  assert.deepEqual(ids, ["Persisted"]);
});

test("CollabSession legacy `options.offline` (file) is implemented via collab-persistence", async (t) => {
  const dir = await mkdtemp(path.join(tmpdir(), "collab-session-offline-compat-file-"));
  const filePath = path.join(dir, "doc.yjslog");

  t.after(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  // First run: write one persisted cell, then detach persistence and make a second edit.
  {
    const session = createCollabSession({
      schema: { autoInit: false },
      offline: { mode: "file", filePath },
    });

    await session.offline?.whenLoaded();

    await session.setCellValue("Sheet1:0:0", "persisted");
    await session.flushLocalPersistence();

    // Detach persistence and verify subsequent edits do *not* persist.
    session.offline?.destroy();

    await session.setCellValue("Sheet1:0:1", "not-persisted");
    await session.flushLocalPersistence();

    session.destroy();
    session.doc.destroy();

    await session.flushLocalPersistence();
  }

  // Second run: only the first edit should be present.
  {
    const session = createCollabSession({
      schema: { autoInit: false },
      offline: { mode: "file", filePath },
    });

    await session.offline?.whenLoaded();

    assert.equal((await session.getCell("Sheet1:0:0"))?.value, "persisted");
    assert.equal(await session.getCell("Sheet1:0:1"), null);

    // Clearing legacy offline storage should remove the persisted state.
    await session.offline?.clear();

    session.destroy();
    session.doc.destroy();
    await session.flushLocalPersistence();
  }

  // Third run: persistence was cleared, so we should start from an empty doc.
  {
    const session = createCollabSession({
      schema: { autoInit: false },
      offline: { mode: "file", filePath },
    });

    await session.offline?.whenLoaded();

    assert.equal(await session.getCell("Sheet1:0:0"), null);

    session.destroy();
    session.doc.destroy();
    await session.flushLocalPersistence();
  }
});

test("CollabSession legacy `options.offline` (indexeddb) is implemented via collab-persistence", async () => {
  const docId = `doc-${randomUUID()}`;
  const keyA = `key-${randomUUID()}`;
  const keyB = `key-${randomUUID()}`;

  {
    const session = createCollabSession({
      schema: { autoInit: false },
      docId,
      offline: { mode: "indexeddb", key: keyA },
    });

    await session.offline?.whenLoaded();
    await session.setCellValue("Sheet1:0:0", 123);

    // Allow the IndexedDB transaction to commit.
    await sleep(10);

    session.destroy();
    session.doc.destroy();
  }

  {
    const session = createCollabSession({
      schema: { autoInit: false },
      docId,
      offline: { mode: "indexeddb", key: keyB },
    });

    await session.offline?.whenLoaded();
    // `offline.key` should override `docId`, so switching keys should isolate state.
    assert.equal(await session.getCell("Sheet1:0:0"), null);

    await session.offline?.clear();
    session.destroy();
    session.doc.destroy();
  }

  {
    const session = createCollabSession({
      schema: { autoInit: false },
      docId,
      offline: { mode: "indexeddb", key: keyA },
    });

    await session.offline?.whenLoaded();
    assert.equal((await session.getCell("Sheet1:0:0"))?.value, 123);

    // Best-effort cleanup.
    await session.offline?.clear();
    session.destroy();
    session.doc.destroy();
  }

  {
    const session = createCollabSession({
      schema: { autoInit: false },
      docId,
      offline: { mode: "indexeddb", key: keyA },
    });

    await session.offline?.whenLoaded();
    assert.equal(await session.getCell("Sheet1:0:0"), null);

    await session.offline?.clear();
    session.destroy();
    session.doc.destroy();
  }
});

test("CollabSession legacy `options.offline` binds before load so edits during load are persisted (file)", async (t) => {
  const dir = await mkdtemp(path.join(tmpdir(), "collab-session-offline-compat-during-load-"));
  const filePath = path.join(dir, "doc.yjslog");

  t.after(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  // Artificially slow down FileCollabPersistence.load() so we can reliably make
  // edits while load is still in flight.
  const docHash = createHash("sha256").update(filePath).digest("hex");
  const hashedPath = path.join(dir, `${docHash}.yjs`);
  const realReadFile = fs.readFile;
  fs.readFile = async (...args) => {
    const candidate = args[0];
    if (candidate && String(candidate) === hashedPath) {
      await sleep(50);
    }
    return await realReadFile(...args);
  };
  t.after(() => {
    fs.readFile = realReadFile;
  });

  {
    const session = createCollabSession({
      schema: { autoInit: false },
      offline: { mode: "file", filePath },
    });

    // Let the persistence start/bind; load is still pending due to delayed readFile.
    await sleep(0);

    await session.setCellValue("Sheet1:0:0", "during-load");
    await session.offline?.whenLoaded();
    await session.flushLocalPersistence();

    session.destroy();
    session.doc.destroy();
  }

  {
    const session = createCollabSession({
      schema: { autoInit: false },
      offline: { mode: "file", filePath },
    });

    await session.offline?.whenLoaded();
    assert.equal((await session.getCell("Sheet1:0:0"))?.value, "during-load");

    await session.offline?.clear();
    session.destroy();
    session.doc.destroy();
  }
});
