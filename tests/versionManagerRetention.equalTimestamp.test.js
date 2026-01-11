import test from "node:test";
import assert from "node:assert/strict";
import { EventEmitter } from "node:events";

import { VersionManager } from "../packages/versioning/src/versioning/versionManager.js";

class FakeDoc extends EventEmitter {
  constructor() {
    super();
    this.state = { value: 0 };
  }

  bump() {
    this.state.value += 1;
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

class InMemoryStore {
  constructor() {
    /** @type {any[]} */
    this.versions = [];
    /** @type {string[]} */
    this.deleted = [];
  }

  async saveVersion(version) {
    this.versions.push(version);
  }

  async getVersion(versionId) {
    return this.versions.find((v) => v.id === versionId) ?? null;
  }

  async listVersions() {
    // Return newest-first to match the store contract. When timestamps are
    // identical, this order is the only reliable tie-breaker.
    return [...this.versions].reverse();
  }

  async updateVersion(_versionId, _patch) {}

  async deleteVersion(versionId) {
    this.deleted.push(versionId);
    this.versions = this.versions.filter((v) => v.id !== versionId);
  }
}

test("VersionManager retention uses store order as a stable tie-breaker for equal timestamps", async () => {
  const doc = new FakeDoc();
  const store = new InMemoryStore();

  // Force identical timestamps for all snapshots.
  const vm = new VersionManager({
    doc,
    store,
    autoStart: false,
    nowMs: () => 0,
    retention: { maxSnapshots: 2 },
  });

  doc.bump();
  const s1 = await vm.createSnapshot({ description: "s1" });
  doc.bump();
  const s2 = await vm.createSnapshot({ description: "s2" });
  doc.bump();
  const s3 = await vm.createSnapshot({ description: "s3" });

  assert.equal(s1.timestampMs, 0);
  assert.equal(s2.timestampMs, 0);
  assert.equal(s3.timestampMs, 0);

  assert.deepEqual(store.deleted, [s1.id], "expected oldest snapshot to be pruned deterministically");

  const remaining = await vm.listVersions();
  const remainingSnapshots = remaining.filter((v) => v.kind === "snapshot");
  assert.deepEqual(
    remainingSnapshots.map((v) => v.description),
    ["s3", "s2"],
    "expected newest snapshots to remain based on store ordering",
  );
});

