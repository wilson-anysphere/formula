import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { createCollabVersioning } from "../src/index.ts";
import { YjsVersionStore } from "../../../versioning/src/store/yjsVersionStore.js";

test("CollabVersioning default store uses YjsVersionStore streaming mode for large snapshots", async (t) => {
  const doc = new Y.Doc();
  t.after(() => doc.destroy());

  const versioning = createCollabVersioning({
    // @ts-expect-error - minimal session stub for unit tests
    session: { doc },
    autoStart: false,
  });
  t.after(() => versioning.destroy());

  const store = versioning.manager.store;
  assert.ok(store instanceof YjsVersionStore);
  assert.equal(store.writeMode, "stream");
  // We keep the default chunk size conservative so each Yjs update stays small.
  assert.equal(store.chunkSize, 64 * 1024);
  // 64KiB chunks with an 8-chunk transaction cap yields ~512KiB snapshot payload
  // per update, comfortably below typical websocket message limits.
  assert.equal(store.maxChunksPerTransaction, 8);
});

test("CollabVersioning default YjsVersionStore options can be overridden without constructing a store", async (t) => {
  const doc = new Y.Doc();
  t.after(() => doc.destroy());

  const versioning = createCollabVersioning({
    // @ts-expect-error - minimal session stub for unit tests
    session: { doc },
    autoStart: false,
    yjsStoreOptions: {
      chunkSize: 32 * 1024,
      maxChunksPerTransaction: 2,
    },
  });
  t.after(() => versioning.destroy());

  const store = versioning.manager.store;
  assert.ok(store instanceof YjsVersionStore);
  assert.equal(store.writeMode, "stream");
  assert.equal(store.chunkSize, 32 * 1024);
  assert.equal(store.maxChunksPerTransaction, 2);
});

