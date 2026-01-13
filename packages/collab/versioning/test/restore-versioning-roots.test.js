import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { createCollabVersioning } from "../src/index.ts";

class InMemoryVersionStore {
  constructor() {
    /** @type {Map<string, any>} */
    this._versions = new Map();
    /** @type {string[]} */
    this._order = [];
  }

  /**
   * @param {any} version
   */
  async saveVersion(version) {
    if (!this._versions.has(version.id)) {
      this._order.push(version.id);
    }
    this._versions.set(version.id, version);
  }

  /**
   * @param {string} versionId
   */
  async getVersion(versionId) {
    return this._versions.get(versionId) ?? null;
  }

  async listVersions() {
    return this._order.map((id) => this._versions.get(id)).filter(Boolean);
  }

  /**
   * @param {string} versionId
   * @param {{ checkpointLocked?: boolean }} patch
   */
  async updateVersion(versionId, patch) {
    const existing = this._versions.get(versionId);
    if (!existing) throw new Error(`Version not found: ${versionId}`);
    this._versions.set(versionId, {
      ...existing,
      ...(patch.checkpointLocked == null ? {} : { checkpointLocked: patch.checkpointLocked }),
    });
  }

  /**
   * @param {string} versionId
   */
  async deleteVersion(versionId) {
    this._versions.delete(versionId);
    this._order = this._order.filter((id) => id !== versionId);
  }
}

test("CollabVersioning restoreVersion does not mutate reserved versions roots (out-of-doc store)", async (t) => {
  const doc = new Y.Doc();
  t.after(() => doc.destroy());

  const metadata = doc.getMap("metadata");
  const versions = doc.getMap("versions");
  const versionsMeta = doc.getMap("versionsMeta");

  metadata.set("title", "Before");
  versions.set("v1", "before");
  versionsMeta.set("m1", "before-meta");

  const store = new InMemoryVersionStore();
  const versioning = createCollabVersioning({
    // @ts-expect-error - minimal session stub for unit tests
    session: { doc },
    store,
    autoStart: false,
    user: { userId: "user-1", userName: "User 1" },
  });
  t.after(() => versioning.destroy());

  const checkpoint = await versioning.createCheckpoint({ name: "checkpoint-1" });

  // Mutate both user-visible state (so we know restore happened) and the reserved
  // internal versioning roots (which should be ignored by restore).
  metadata.set("title", "After");

  versions.set("v1", "after");
  versions.set("v2", "new");

  versionsMeta.delete("m1");
  versionsMeta.set("m2", "after-meta");

  await versioning.restoreVersion(checkpoint.id);

  // User-visible roots should restore.
  assert.equal(metadata.get("title"), "Before");

  // Reserved roots should remain at post-mutation values.
  assert.equal(versions.get("v1"), "after");
  assert.equal(versions.get("v2"), "new");
  assert.equal(versionsMeta.has("m1"), false);
  assert.equal(versionsMeta.get("m2"), "after-meta");
});

