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

class InMemoryVersionStore {
  constructor() {
    /** @type {Map<string, any>} */
    this._versions = new Map();
    /** @type {string[]} */
    this._order = [];
  }

  async saveVersion(version) {
    if (!this._versions.has(version.id)) this._order.push(version.id);
    this._versions.set(version.id, version);
  }
  async getVersion(versionId) {
    return this._versions.get(versionId) ?? null;
  }
  async listVersions() {
    return this._order.map((id) => this._versions.get(id)).filter(Boolean);
  }
  async updateVersion(versionId, patch) {
    const existing = this._versions.get(versionId);
    if (!existing) throw new Error(`Version not found: ${versionId}`);
    this._versions.set(versionId, { ...existing, ...patch });
  }
  async deleteVersion(versionId) {
    this._versions.delete(versionId);
    this._order = this._order.filter((id) => id !== versionId);
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

test("VersionManager.destroy unsubscribes from doc updates (doc.off path)", () => {
  const doc = new FakeDoc();
  const store = new InMemoryVersionStore();
  const vm = new VersionManager({ doc, store, autoStart: false });

  assert.equal(vm.dirty, false);
  doc.setCell("r0c0", { value: 1 });
  assert.equal(vm.dirty, true);

  vm.dirty = false;
  vm.destroy();

  doc.setCell("r0c1", { value: 2 });
  assert.equal(vm.dirty, false);
});

test("VersionManager.destroy unsubscribes from doc updates (doc.removeListener path)", () => {
  const doc = new FakeDoc();
  // Simulate a doc that only implements `removeListener` (older EventEmitter API).
  // Node's EventEmitter provides `off` as an alias, but some event emitter shims do not.
  // Shadow the prototype method so `typeof doc.off !== 'function'`.
  // @ts-expect-error - intentionally overriding
  doc.off = undefined;

  const store = new InMemoryVersionStore();
  const vm = new VersionManager({ doc, store, autoStart: false });

  assert.equal(vm.dirty, false);
  doc.setCell("r0c0", { value: 1 });
  assert.equal(vm.dirty, true);

  vm.dirty = false;
  vm.destroy();

  doc.setCell("r0c1", { value: 2 });
  assert.equal(vm.dirty, false);
});

test("VersionManager.destroy unsubscribes from doc updates (unsubscribe-returning doc.on path)", () => {
  class UnsubscribableDoc {
    constructor() {
      /** @type {Set<() => void>} */
      this.listeners = new Set();
    }
    encodeState() {
      return new Uint8Array();
    }
    applyState(_snapshot) {}
    on(event, listener) {
      if (event !== "update") throw new Error(`unsupported event: ${event}`);
      this.listeners.add(listener);
      return () => this.listeners.delete(listener);
    }
    emitUpdate() {
      for (const listener of Array.from(this.listeners)) listener();
    }
  }

  const doc = new UnsubscribableDoc();
  const store = new InMemoryVersionStore();
  const vm = new VersionManager({ doc, store, autoStart: false });

  assert.equal(doc.listeners.size, 1);
  doc.emitUpdate();
  assert.equal(vm.dirty, true);

  vm.dirty = false;
  vm.destroy();
  assert.equal(doc.listeners.size, 0);

  doc.emitUpdate();
  assert.equal(vm.dirty, false);
});

test("VersionManager.destroy cleans up an in-flight autosnapshot save", async () => {
  const doc = new FakeDoc();

  class DeferredStore extends InMemoryVersionStore {
    constructor() {
      super();
      this._deferred = null;
      this.saveStarted = false;
    }

    defer() {
      /** @type {{ promise: Promise<void>, resolve: () => void }} */
      const deferred = {};
      deferred.promise = new Promise((resolve) => {
        deferred.resolve = () => resolve();
      });
      // @ts-expect-error - assigned above
      this._deferred = deferred;
      return deferred;
    }

    async saveVersion(version) {
      this.saveStarted = true;
      if (this._deferred) {
        await this._deferred.promise;
      }
      return await super.saveVersion(version);
    }
  }

  const store = new DeferredStore();
  const deferred = store.defer();

  const vm = new VersionManager({ doc, store, autoStart: false, autoSnapshotIntervalMs: 1 });

  // Mark dirty and start an autosnapshot tick.
  doc.setCell("r0c0", { value: 1 });
  assert.equal(vm.dirty, true);

  const tickPromise = vm._autoSnapshotTick();

  // Wait until saveVersion has started (tick is now awaiting our deferred).
  while (!store.saveStarted) {
    await new Promise((resolve) => setImmediate(resolve));
  }

  // Destroy while the autosnapshot is in-flight.
  vm.destroy();
  deferred.resolve();

  await tickPromise;

  const versions = await store.listVersions();
  assert.equal(versions.length, 0);
});

test("VersionManager.destroy prevents in-flight retention pruning from mutating the store", async () => {
  const doc = new FakeDoc();

  class DeferredListStore extends InMemoryVersionStore {
    constructor() {
      super();
      this.listStarted = false;
      this.deleteCalls = 0;
      this._deferred = null;
    }

    deferList() {
      /** @type {{ promise: Promise<void>, resolve: () => void }} */
      const deferred = {};
      deferred.promise = new Promise((resolve) => {
        deferred.resolve = () => resolve();
      });
      // @ts-expect-error - assigned above
      this._deferred = deferred;
      return deferred;
    }

    async listVersions() {
      this.listStarted = true;
      if (this._deferred) await this._deferred.promise;
      return await super.listVersions();
    }

    async deleteVersion(versionId) {
      this.deleteCalls += 1;
      return await super.deleteVersion(versionId);
    }
  }

  const store = new DeferredListStore();
  const deferred = store.deferList();

  const vm = new VersionManager({
    doc,
    store,
    autoStart: false,
    // Ensure retention would try to delete the newly-created snapshot.
    retention: { maxSnapshots: 0 },
  });

  doc.setCell("r0c0", { value: 1 });

  const createPromise = vm.createSnapshot();

  while (!store.listStarted) {
    await new Promise((resolve) => setImmediate(resolve));
  }

  vm.destroy();
  deferred.resolve();
  await createPromise;

  assert.equal(store.deleteCalls, 0);
});
