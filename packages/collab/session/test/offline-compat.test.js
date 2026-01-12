import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { randomUUID } from "node:crypto";

import { indexedDB, IDBKeyRange } from "fake-indexeddb";

import { createCollabSession } from "../src/index.ts";

globalThis.indexedDB = indexedDB;
globalThis.IDBKeyRange = IDBKeyRange;

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

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
  const key = `doc-${randomUUID()}`;

  {
    const session = createCollabSession({
      schema: { autoInit: false },
      offline: { mode: "indexeddb", key },
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
      offline: { mode: "indexeddb", key },
    });

    await session.offline?.whenLoaded();
    assert.equal((await session.getCell("Sheet1:0:0"))?.value, 123);

    await session.offline?.clear();
    session.destroy();
    session.doc.destroy();
  }

  {
    const session = createCollabSession({
      schema: { autoInit: false },
      offline: { mode: "indexeddb", key },
    });

    await session.offline?.whenLoaded();
    assert.equal(await session.getCell("Sheet1:0:0"), null);

    // Best-effort cleanup (DB is already empty, but ensure the temp key is removed).
    await session.offline?.clear();
    session.destroy();
    session.doc.destroy();
  }
});

