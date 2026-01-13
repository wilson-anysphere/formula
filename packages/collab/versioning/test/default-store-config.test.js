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

test("CollabVersioning ignores accidental yjsStoreOptions.doc overrides", async (t) => {
  const doc = new Y.Doc();
  const otherDoc = new Y.Doc();
  t.after(() => doc.destroy());
  t.after(() => otherDoc.destroy());

  const versioning = createCollabVersioning({
    // @ts-expect-error - minimal session stub for unit tests
    session: { doc },
    autoStart: false,
    // `yjsStoreOptions` should never be able to override the session doc binding,
    // even if a JS caller passes a `doc` field (not part of the public type).
    // @ts-expect-error - invalid option is intentional for test coverage
    yjsStoreOptions: { doc: otherDoc },
  });
  t.after(() => versioning.destroy());

  const store = versioning.manager.store;
  assert.ok(store instanceof YjsVersionStore);
  assert.equal(store.doc, doc);
});

test("CollabVersioning derives maxChunksPerTransaction from chunkSize when not provided", async (t) => {
  {
    const doc = new Y.Doc();
    t.after(() => doc.destroy());

    const versioning = createCollabVersioning({
      // @ts-expect-error - minimal session stub for unit tests
      session: { doc },
      autoStart: false,
      yjsStoreOptions: { chunkSize: 128 * 1024 },
    });
    t.after(() => versioning.destroy());

    const store = versioning.manager.store;
    assert.ok(store instanceof YjsVersionStore);
    assert.equal(store.writeMode, "stream");
    assert.equal(store.chunkSize, 128 * 1024);
    // Default target is ~512KiB per streamed update.
    assert.equal(store.maxChunksPerTransaction, 4);
  }

  {
    const doc = new Y.Doc();
    t.after(() => doc.destroy());

    const versioning = createCollabVersioning({
      // @ts-expect-error - minimal session stub for unit tests
      session: { doc },
      autoStart: false,
      yjsStoreOptions: { chunkSize: 32 * 1024 },
    });
    t.after(() => versioning.destroy());

    const store = versioning.manager.store;
    assert.ok(store instanceof YjsVersionStore);
    assert.equal(store.writeMode, "stream");
    assert.equal(store.chunkSize, 32 * 1024);
    // Cap defaults at 16 chunks/update to avoid overly chatty streams.
    assert.equal(store.maxChunksPerTransaction, 16);
  }
});

test("CollabVersioning ignores yjsStoreOptions when an explicit store is provided", async (t) => {
  const doc = new Y.Doc();
  t.after(() => doc.destroy());

  const store = new YjsVersionStore({ doc, writeMode: "single" });
  const versioning = createCollabVersioning({
    // @ts-expect-error - minimal session stub for unit tests
    session: { doc },
    store,
    autoStart: false,
    yjsStoreOptions: {
      writeMode: "stream",
      chunkSize: 32 * 1024,
      maxChunksPerTransaction: 2,
    },
  });
  t.after(() => versioning.destroy());

  assert.equal(versioning.manager.store, store);
  assert.equal(store.writeMode, "single");
});

test("CollabVersioning allows opting out of streaming via yjsStoreOptions.writeMode", async (t) => {
  const doc = new Y.Doc();
  t.after(() => doc.destroy());

  const versioning = createCollabVersioning({
    // @ts-expect-error - minimal session stub for unit tests
    session: { doc },
    autoStart: false,
    yjsStoreOptions: { writeMode: "single" },
  });
  t.after(() => versioning.destroy());

  const store = versioning.manager.store;
  assert.ok(store instanceof YjsVersionStore);
  assert.equal(store.writeMode, "single");
});
