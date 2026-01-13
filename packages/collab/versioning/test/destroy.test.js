import test from "node:test";
import assert from "node:assert/strict";
import * as Y from "yjs";

import { createCollabVersioning } from "../src/index.ts";

/**
 * Track `doc.on("update", ...)` registrations without relying on Yjs internals.
 *
 * @param {Y.Doc} doc
 */
function trackUpdateListeners(doc) {
  /** @type {Set<unknown>} */
  const listeners = new Set();
  const originalOn = doc.on.bind(doc);
  const originalOff = doc.off.bind(doc);

  doc.on = (eventName, listener) => {
    if (eventName === "update") listeners.add(listener);
    return originalOn(eventName, listener);
  };
  doc.off = (eventName, listener) => {
    if (eventName === "update") listeners.delete(listener);
    return originalOff(eventName, listener);
  };

  return listeners;
}

test("CollabVersioning.destroy unsubscribes from Yjs document updates", async (t) => {
  const doc = new Y.Doc();
  t.after(() => doc.destroy());
  const updateListeners = trackUpdateListeners(doc);
  const baselineListenerCount = updateListeners.size;

  // Instantiate workbook roots before VersionManager attaches listeners so the
  // test only measures the effects of *updates*.
  const cells = doc.getMap("cells");

  assert.equal(updateListeners.size, baselineListenerCount);

  const versioning = createCollabVersioning({
    // @ts-expect-error - minimal session stub for unit tests
    session: { doc },
    autoStart: false,
  });

  assert.ok(updateListeners.size > baselineListenerCount);
  assert.equal(versioning.manager.dirty, false);

  versioning.destroy();
  assert.equal(updateListeners.size, baselineListenerCount);

  // Workbook edits after destroy should not mark the old manager dirty.
  cells.set("Sheet1:0:0", "alpha");
  assert.equal(versioning.manager.dirty, false);

  // A destroyed manager should not attempt to snapshot since it never becomes dirty.
  assert.equal(await versioning.manager.maybeSnapshot(), null);
});

test("CollabVersioning create/destroy cycles do not accumulate dirty listeners", (t) => {
  const doc = new Y.Doc();
  t.after(() => doc.destroy());
  const updateListeners = trackUpdateListeners(doc);

  const cells = doc.getMap("cells");

  const baselineListenerCount = updateListeners.size;
  assert.equal(updateListeners.size, baselineListenerCount);

  const versioning1 = createCollabVersioning({
    // @ts-expect-error - minimal session stub for unit tests
    session: { doc },
    autoStart: false,
  });
  const manager1 = versioning1.manager;
  assert.ok(updateListeners.size > baselineListenerCount);
  const listenerDelta = updateListeners.size - baselineListenerCount;
  versioning1.destroy();
  assert.equal(updateListeners.size, baselineListenerCount);

  const versioning2 = createCollabVersioning({
    // @ts-expect-error - minimal session stub for unit tests
    session: { doc },
    autoStart: false,
  });
  const manager2 = versioning2.manager;
  assert.equal(updateListeners.size, baselineListenerCount + listenerDelta);

  cells.set("Sheet1:0:0", "alpha");

  assert.equal(manager1.dirty, false);
  assert.equal(manager2.dirty, true);

  versioning2.destroy();
  assert.equal(updateListeners.size, baselineListenerCount);
});

test("CollabVersioning.destroy prevents autosnapshot ticks from creating new versions", async (t) => {
  const doc = new Y.Doc();
  t.after(() => doc.destroy());

  const cells = doc.getMap("cells");

  const versioning = createCollabVersioning({
    // @ts-expect-error - minimal session stub for unit tests
    session: { doc },
    autoStart: false,
    // Make manual tick testing deterministic.
    autoSnapshotIntervalMs: 1,
  });
  t.after(() => versioning.destroy());

  // Mark dirty, then destroy before any snapshot can be created.
  cells.set("Sheet1:0:0", "alpha");
  assert.equal(versioning.manager.dirty, true);

  versioning.destroy();

  // Even if an autosnapshot tick were to run after teardown (race between
  // interval callback and destroy), it should be a no-op.
  await versioning.manager._autoSnapshotTick();

  const versions = await versioning.listVersions();
  assert.equal(versions.length, 0);
});

test("CollabVersioning.destroy prevents retention work from running on post-destroy autosnapshot ticks", async (t) => {
  const doc = new Y.Doc();
  t.after(() => doc.destroy());

  const versioning = createCollabVersioning({
    // @ts-expect-error - minimal session stub for unit tests
    session: { doc },
    autoStart: false,
    // Configure retention so autosnapshot ticks would normally call `listVersions`.
    retention: { maxSnapshots: 1 },
  });
  t.after(() => versioning.destroy());

  // Seed a snapshot so retention has meaningful work (and to avoid relying on YjsVersionStore internals).
  await versioning.createSnapshot({ description: "seed" });

  const store = versioning.manager.store;
  const originalListVersions = store.listVersions.bind(store);
  let listVersionsCalls = 0;
  store.listVersions = async (...args) => {
    listVersionsCalls += 1;
    return await originalListVersions(...args);
  };

  // Reset after the seed snapshot's own retention work.
  listVersionsCalls = 0;

  versioning.destroy();
  await versioning.manager._autoSnapshotTick();

  assert.equal(listVersionsCalls, 0);
});
