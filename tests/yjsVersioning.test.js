import test from "node:test";
import assert from "node:assert/strict";
import os from "node:os";
import path from "node:path";
import { promises as fs } from "node:fs";

import * as Y from "yjs";

import { SQLiteVersionStore } from "../packages/versioning/src/store/sqliteVersionStore.js";
import { VersionManager } from "../packages/versioning/src/versioning/versionManager.js";
import { createYjsSpreadsheetDocAdapter } from "../packages/versioning/src/yjs/yjsSpreadsheetDocAdapter.js";
import { diffYjsSnapshots } from "../packages/versioning/src/yjs/diffSnapshots.js";
import { diffYjsVersionAgainstCurrent } from "../packages/versioning/src/yjs/versionHistory.js";

test("Yjs + SQLite: checkpoint, edit, diff, restore (history preserved)", async () => {
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "versioning-yjs-"));
  const storePath = path.join(tmpDir, "versions.sqlite");

  const store = new SQLiteVersionStore({ filePath: storePath });

  const ydoc = new Y.Doc();
  const sheets = ydoc.getArray("sheets");
  const cells = ydoc.getMap("cells");
  const comments = ydoc.getArray("comments");

  // Minimal sheet metadata for realism.
  const sheet = new Y.Map();
  sheet.set("id", "sheet1");
  sheet.set("name", "Sheet1");
  sheets.push([sheet]);

  ydoc.transact(() => {
    const cell = new Y.Map();
    cell.set("value", "move-me");
    cell.set("formula", "=A1+B1");
    cells.set("sheet1:0:0", cell);

    const comment = new Y.Map();
    comment.set("id", "c1");
    comment.set("text", "original");
    comments.push([comment]);
  });

  const doc = createYjsSpreadsheetDocAdapter(ydoc);
  const vm = new VersionManager({
    doc,
    store,
    autoStart: false,
    user: { userId: "u1", userName: "User" },
  });

  const checkpoint = await vm.createCheckpoint({ name: "Approved", locked: true });
  assert.equal(checkpoint.kind, "checkpoint");

  // Simulate a move via cut/paste to a new location, with a semantically equivalent formula.
  ydoc.transact(() => {
    cells.delete("sheet1:0:0");
    const moved = new Y.Map();
    moved.set("value", "move-me");
    moved.set("formula", "=B1 + A1");
    cells.set("sheet1:2:3", moved);

    const comment = comments.get(0);
    if (comment) comment.set("text", "edited");
  });
  const snapshot = await vm.createSnapshot({ description: "after move" });

  const diff = diffYjsSnapshots({
    beforeSnapshot: checkpoint.snapshot,
    afterSnapshot: snapshot.snapshot,
    sheetId: "sheet1",
  });
  assert.equal(diff.moved.length, 1);
  assert.deepEqual(diff.moved[0].oldLocation, { row: 0, col: 0 });
  assert.deepEqual(diff.moved[0].newLocation, { row: 2, col: 3 });

  const diffToCurrent = await diffYjsVersionAgainstCurrent({
    versionManager: vm,
    versionId: checkpoint.id,
    sheetId: "sheet1",
  });
  assert.equal(diffToCurrent.moved.length, 1);

  // Restore should bring back the original cell location/content.
  await vm.restoreVersion(checkpoint.id);
  assert.equal(cells.has("sheet1:2:3"), false);
  const restoredCell = cells.get("sheet1:0:0");
  assert.ok(restoredCell);
  assert.equal(restoredCell.get("value"), "move-me");
  assert.equal(comments.get(0)?.get("text"), "original");

  const versions = await vm.listVersions();
  assert.equal(versions.length, 3);

  const persistedCheckpoint = versions.find((v) => v.id === checkpoint.id);
  assert.ok(persistedCheckpoint);
  assert.equal(persistedCheckpoint.kind, "checkpoint");
  assert.equal(persistedCheckpoint.checkpointLocked, true);

  store.close();
});
