import test from "node:test";
import assert from "node:assert/strict";
import { EventEmitter } from "node:events";
import os from "node:os";
import path from "node:path";
import { promises as fs } from "node:fs";
import crypto from "node:crypto";

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

test("VersionManager retention: maxSnapshots keeps newest N snapshots and preserves checkpoints", async () => {
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "versioning-retention-count-"));
  const storePath = path.join(tmpDir, `versions-${crypto.randomUUID()}.json`);

  let now = 0;
  const doc = new FakeDoc();
  const store = new FileVersionStore({ filePath: storePath });
  const vm = new VersionManager({
    doc,
    store,
    autoStart: false,
    nowMs: () => now,
    retention: { maxSnapshots: 2 },
  });

  /** @type {Set<string>} */
  const pruned = new Set();
  vm.on("versionsPruned", ({ deletedIds }) => {
    for (const id of deletedIds) pruned.add(id);
  });

  for (let i = 0; i < 5; i += 1) {
    doc.setCell("r0c0", { value: i });
    now += 10;
    await vm.createSnapshot({ description: `s${i}` });
  }

  now += 10;
  const checkpoint = await vm.createCheckpoint({ name: "Approved" });

  const versions = await vm.listVersions();
  const snapshots = versions.filter((v) => v.kind === "snapshot");
  const checkpoints = versions.filter((v) => v.kind === "checkpoint");

  assert.equal(snapshots.length, 2);
  assert.deepEqual(
    snapshots.map((v) => v.description),
    ["s4", "s3"],
    "expected newest snapshots to be retained"
  );

  assert.equal(checkpoints.length, 1);
  assert.equal(checkpoints[0].id, checkpoint.id);

  assert.equal(pruned.size, 3, "expected older snapshots to be pruned");
});

test("VersionManager retention: maxAgeMs deletes old snapshots while preserving checkpoints", async () => {
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "versioning-retention-age-"));
  const storePath = path.join(tmpDir, `versions-${crypto.randomUUID()}.json`);

  let now = 0;
  const doc = new FakeDoc();
  const store = new FileVersionStore({ filePath: storePath });
  const vm = new VersionManager({
    doc,
    store,
    autoStart: false,
    nowMs: () => now,
    retention: { maxAgeMs: 1000 },
  });

  doc.setCell("r0c0", { value: 1 });
  now = 0;
  await vm.createSnapshot({ description: "old1" });

  doc.setCell("r0c0", { value: 2 });
  now = 500;
  await vm.createSnapshot({ description: "old2" });

  now = 600;
  const checkpoint = await vm.createCheckpoint({ name: "Checkpoint" });

  // Advance time beyond the retention window and create a new snapshot to trigger pruning.
  doc.setCell("r0c0", { value: 3 });
  now = 2000;
  await vm.createSnapshot({ description: "new" });

  const versions = await vm.listVersions();
  assert.equal(
    versions.some((v) => v.description === "old1" || v.description === "old2"),
    false,
    "expected old snapshots to be deleted"
  );
  assert.ok(versions.some((v) => v.id === checkpoint.id), "expected checkpoint to be preserved");
});

test("VersionManager retention: locked checkpoints are never deleted", async () => {
  const tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "versioning-retention-locked-"));
  const storePath = path.join(tmpDir, `versions-${crypto.randomUUID()}.json`);

  let now = 0;
  const doc = new FakeDoc();
  const store = new FileVersionStore({ filePath: storePath });
  const vm = new VersionManager({
    doc,
    store,
    autoStart: false,
    nowMs: () => now,
    retention: { maxAgeMs: 10, keepCheckpoints: false },
  });

  doc.setCell("r0c0", { value: 1 });
  now = 0;
  const locked = await vm.createCheckpoint({ name: "Locked", locked: true });

  doc.setCell("r0c0", { value: 2 });
  now = 1;
  const unlocked = await vm.createCheckpoint({ name: "Unlocked", locked: false });

  doc.setCell("r0c0", { value: 3 });
  now = 100;
  await vm.createSnapshot({ description: "trigger" });

  const versions = await vm.listVersions();
  assert.ok(versions.some((v) => v.id === locked.id), "expected locked checkpoint to remain");
  assert.equal(
    versions.some((v) => v.id === unlocked.id),
    false,
    "expected unlocked checkpoint to be pruned when keepCheckpoints is false"
  );
});

