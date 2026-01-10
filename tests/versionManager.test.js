import test from "node:test";
import assert from "node:assert/strict";
import { EventEmitter } from "node:events";
import os from "node:os";
import path from "node:path";
import { promises as fs } from "node:fs";

import { FileVersionStore } from "../packages/versioning/src/store/fileVersionStore.js";
import { VersionManager } from "../packages/versioning/src/versioning/versionManager.js";

class FakeDoc extends EventEmitter {
  constructor() {
    super();
    this.state = { cells: {} };
  }

  setCell(key, value) {
    this.state.cells[key] = value;
    this.emit("update");
  }

  encodeState() {
    return Buffer.from(JSON.stringify(this.state), "utf8");
  }

  applyState(snapshot) {
    this.state = JSON.parse(Buffer.from(snapshot).toString("utf8"));
    this.emit("update");
  }
}

test("VersionManager: checkpoint, edit, compare snapshots, restore creates new head without destroying history", async () => {
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "versioning-"));
  const storePath = path.join(tmpDir, "versions.json");

  const doc = new FakeDoc();
  const store = new FileVersionStore({ filePath: storePath });
  const vm = new VersionManager({ doc, store, autoStart: false, user: { userId: "u1", userName: "User" } });

  doc.setCell("r0c0", { value: 1 });
  const checkpoint = await vm.createCheckpoint({ name: "Approved", locked: true });

  doc.setCell("r0c0", { value: 2 });
  await vm.createSnapshot({ description: "edit" });

  const versionsBeforeRestore = await vm.listVersions();
  assert.equal(versionsBeforeRestore.length, 2);

  await vm.restoreVersion(checkpoint.id);
  assert.deepEqual(doc.state.cells["r0c0"], { value: 1 });

  const versionsAfterRestore = await vm.listVersions();
  assert.equal(versionsAfterRestore.length, 3);
  assert.ok(versionsAfterRestore.some((v) => v.id === checkpoint.id));

  // Restore should not mutate checkpoint data; it should create a new head version.
  const restoredHead = versionsAfterRestore.find((v) => v.kind === "restore");
  assert.ok(restoredHead);

  // Persistence: a fresh manager should see the same history.
  const doc2 = new FakeDoc();
  const vm2 = new VersionManager({ doc: doc2, store: new FileVersionStore({ filePath: storePath }), autoStart: false });
  const versionsReloaded = await vm2.listVersions();
  assert.equal(versionsReloaded.length, 3);
});

