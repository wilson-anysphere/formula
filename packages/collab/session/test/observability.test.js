import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { createCollabSession } from "../src/index.ts";

function createMockProvider() {
  /** @type {Map<string, Set<(...args: any[]) => void>>} */
  const listeners = new Map();
  return {
    synced: false,
    on(event, cb) {
      let set = listeners.get(event);
      if (!set) {
        set = new Set();
        listeners.set(event, set);
      }
      set.add(cb);
    },
    off(event, cb) {
      const set = listeners.get(event);
      if (!set) return;
      set.delete(cb);
      if (set.size === 0) listeners.delete(event);
    },
    emit(event, ...args) {
      for (const cb of listeners.get(event) ?? []) cb(...args);
    },
    destroy() {},
  };
}

test("CollabSession observability: sync state + status subscriptions", () => {
  const doc = new Y.Doc();
  const provider = createMockProvider();
  const session = createCollabSession({ doc, provider, schema: { autoInit: false } });

  assert.deepEqual(session.getSyncState(), { connected: false, synced: false });

  /** @type {Array<{ connected: boolean; synced: boolean }>} */
  const seen = [];
  const unsubscribe = session.onStatusChange((state) => seen.push(state));

  provider.emit("status", { status: "connected" });
  assert.deepEqual(session.getSyncState(), { connected: true, synced: false });
  assert.deepEqual(seen.at(-1), { connected: true, synced: false });

  provider.emit("sync", true);
  assert.deepEqual(session.getSyncState(), { connected: true, synced: true });
  assert.deepEqual(seen.at(-1), { connected: true, synced: true });

  provider.emit("status", { status: "disconnected" });
  assert.deepEqual(session.getSyncState(), { connected: false, synced: false });
  assert.deepEqual(seen.at(-1), { connected: false, synced: false });

  unsubscribe();
  provider.emit("status", { status: "connected" });
  assert.equal(seen.length, 3);

  session.destroy();
  doc.destroy();
});

test("CollabSession observability: update size stats track local updates only", () => {
  const doc = new Y.Doc();
  const session = createCollabSession({ doc, schema: { autoInit: false } });

  session.transactLocal(() => {
    session.metadata.set("a", 1);
  });
  const stats1 = session.getUpdateStats();
  assert.ok(stats1.lastUpdateBytes > 0);
  assert.equal(stats1.lastUpdateBytes, stats1.maxRecentBytes);
  assert.equal(stats1.avgRecentBytes, stats1.lastUpdateBytes);

  session.transactLocal(() => {
    session.metadata.set("b", 2);
  });
  const stats2 = session.getUpdateStats();
  assert.ok(stats2.lastUpdateBytes > 0);
  assert.equal(stats2.maxRecentBytes, Math.max(stats1.lastUpdateBytes, stats2.lastUpdateBytes));
  assert.equal(stats2.avgRecentBytes, (stats1.lastUpdateBytes + stats2.lastUpdateBytes) / 2);

  const beforeRemote = session.getUpdateStats();
  const remoteDoc = new Y.Doc();
  remoteDoc.getMap("remote").set("c", 3);
  Y.applyUpdate(session.doc, Y.encodeStateAsUpdate(remoteDoc), Symbol("remote"));
  assert.deepEqual(session.getUpdateStats(), beforeRemote);

  session.destroy();
  doc.destroy();
  remoteDoc.destroy();
});

test("CollabSession observability: local persistence status exposes load + flush", async () => {
  const doc = new Y.Doc();
  /** @type {string[]} */
  const flushes = [];

  const persistence = {
    async load() {},
    bind() {
      return { destroy: async () => {} };
    },
    async flush(docId) {
      flushes.push(String(docId));
    },
  };

  const session = createCollabSession({ doc, docId: "doc-1", persistence, schema: { autoInit: false } });
  await session.whenLocalPersistenceLoaded();

  const state1 = session.getLocalPersistenceState();
  assert.deepEqual(
    { enabled: state1.enabled, loaded: state1.loaded, lastFlushedAt: state1.lastFlushedAt },
    { enabled: true, loaded: true, lastFlushedAt: null }
  );

  await session.flushLocalPersistence();
  assert.deepEqual(flushes, ["doc-1"]);
  const state2 = session.getLocalPersistenceState();
  assert.equal(typeof state2.lastFlushedAt, "number");
  assert.ok((state2.lastFlushedAt ?? 0) > 0);

  session.destroy();
  doc.destroy();
});

