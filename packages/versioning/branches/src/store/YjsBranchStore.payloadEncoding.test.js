import assert from "node:assert/strict";
import test from "node:test";

import * as Y from "yjs";
import { requireYjsCjs } from "../../../../collab/yjs-utils/test/require-yjs-cjs.js";

import { diffDocumentStates } from "../patch.js";
import { emptyDocumentState, normalizeDocumentState } from "../state.js";
import { YjsBranchStore } from "./YjsBranchStore.js";

/**
 * Deterministic pseudo-random bytes for tests (avoid crypto + avoid large runs of
 * the same character, which gzip would compress too well).
 *
 * @param {number} length
 */
function pseudoRandomBytes(length) {
  const out = new Uint8Array(length);
  let x = 0x12345678;
  for (let i = 0; i < length; i += 1) {
    // LCG-ish
    x = (1103515245 * x + 12345) >>> 0;
    out[i] = x & 0xff;
  }
  return out;
}

/**
 * @param {number} byteLength
 */
function pseudoRandomBase64(byteLength) {
  // eslint-disable-next-line no-undef
  return Buffer.from(pseudoRandomBytes(byteLength)).toString("base64");
}

function makeInitialState() {
  const state = emptyDocumentState();
  state.sheets.order = ["Sheet1"];
  state.sheets.metaById = {
    Sheet1: { id: "Sheet1", name: "Sheet1", view: { frozenRows: 0, frozenCols: 0 } },
  };
  state.cells = { Sheet1: {} };
  return state;
}

test("YjsBranchStore: payloadEncoding=gzip-chunks round-trips patch + snapshot", async () => {
  const doc = new Y.Doc();
  const store = new YjsBranchStore({
    ydoc: doc,
    payloadEncoding: "gzip-chunks",
    chunkSize: 1024,
    maxChunksPerTransaction: 2,
    snapshotEveryNCommits: 1,
  });

  const actor = { userId: "u1", role: "owner" };
  const docId = "doc-1";

  await store.ensureDocument(docId, actor, makeInitialState());

  const main = await store.getBranch(docId, "main");
  assert.ok(main);

  const current = await store.getDocumentStateAtCommit(main.headCommitId);
  const next = structuredClone(current);
  next.metadata.big = pseudoRandomBase64(80 * 1024);

  const patch = diffDocumentStates(current, next);
  const commit = await store.createCommit({
    docId,
    parentCommitId: main.headCommitId,
    mergeParentCommitId: null,
    createdBy: actor.userId,
    createdAt: 123,
    message: "big",
    patch,
    nextState: next,
  });

  const loaded = await store.getCommit(commit.id);
  assert.ok(loaded);
  assert.deepEqual(loaded.patch, patch);

  const loadedState = await store.getDocumentStateAtCommit(commit.id);
  assert.deepEqual(loadedState, normalizeDocumentState(next));

  const commitsMap = doc.getMap("branching:commits");
  const rawCommit = commitsMap.get(commit.id);
  assert.ok(rawCommit instanceof Y.Map);
  assert.equal(rawCommit.get("patch"), undefined);
  assert.equal(rawCommit.get("patchEncoding"), "gzip-chunks");
  assert.equal(rawCommit.get("commitComplete"), true);

  const patchChunks = rawCommit.get("patchChunks");
  assert.ok(patchChunks instanceof Y.Array);
  assert.ok(patchChunks.length > 1);

  const snapshotChunks = rawCommit.get("snapshotChunks");
  assert.ok(snapshotChunks instanceof Y.Array);
  assert.ok(snapshotChunks.length > 1);
});

test("YjsBranchStore: can read legacy json commits when configured for gzip-chunks", async () => {
  const doc = new Y.Doc();
  const docId = "doc-2";
  const actor = { userId: "u1", role: "owner" };

  const legacyStore = new YjsBranchStore({
    ydoc: doc,
    payloadEncoding: "json",
    snapshotEveryNCommits: 1,
  });

  await legacyStore.ensureDocument(docId, actor, makeInitialState());
  const main = await legacyStore.getBranch(docId, "main");
  assert.ok(main);

  const current = await legacyStore.getDocumentStateAtCommit(main.headCommitId);
  const next = structuredClone(current);
  next.metadata.small = "hello";
  const patch = diffDocumentStates(current, next);

  const commit = await legacyStore.createCommit({
    docId,
    parentCommitId: main.headCommitId,
    mergeParentCommitId: null,
    createdBy: actor.userId,
    createdAt: 123,
    message: "small",
    patch,
    nextState: next,
  });

  const reader = new YjsBranchStore({
    ydoc: doc,
    payloadEncoding: "gzip-chunks",
    chunkSize: 1024,
    maxChunksPerTransaction: 1,
  });

  const loaded = await reader.getCommit(commit.id);
  assert.ok(loaded);
  assert.deepEqual(loaded.patch, patch);

  const loadedState = await reader.getDocumentStateAtCommit(commit.id);
  assert.deepEqual(loadedState, normalizeDocumentState(next));
});

test("YjsBranchStore initializes when roots were created by a different Yjs instance (CJS getMap)", async () => {
  const Ycjs = requireYjsCjs();

  const doc = new Y.Doc();

  // Simulate a mixed module loader environment where another Yjs instance eagerly
  // instantiates the BranchStore roots before YjsBranchStore is constructed.
  Ycjs.Doc.prototype.getMap.call(doc, "branching:branches");
  Ycjs.Doc.prototype.getMap.call(doc, "branching:commits");
  Ycjs.Doc.prototype.getMap.call(doc, "branching:meta");

  assert.throws(() => doc.getMap("branching:branches"), /different constructor/);
  assert.throws(() => doc.getMap("branching:commits"), /different constructor/);
  assert.throws(() => doc.getMap("branching:meta"), /different constructor/);

  const store = new YjsBranchStore({
    ydoc: doc,
    payloadEncoding: "gzip-chunks",
    chunkSize: 1024,
    maxChunksPerTransaction: 2,
    snapshotEveryNCommits: 1,
  });

  // Root normalization should re-wrap foreign roots into the local Yjs instance.
  assert.ok(doc.share.get("branching:branches") instanceof Y.Map);
  assert.ok(doc.share.get("branching:commits") instanceof Y.Map);
  assert.ok(doc.share.get("branching:meta") instanceof Y.Map);

  assert.ok(doc.getMap("branching:branches") instanceof Y.Map);
  assert.ok(doc.getMap("branching:commits") instanceof Y.Map);
  assert.ok(doc.getMap("branching:meta") instanceof Y.Map);

  const actor = { userId: "u1", role: "owner" };
  const docId = "doc-foreign-branching";

  await store.ensureDocument(docId, actor, makeInitialState());
  const main = await store.getBranch(docId, "main");
  assert.ok(main);

  const current = await store.getDocumentStateAtCommit(main.headCommitId);
  const next = structuredClone(current);
  next.metadata.foo = "bar";

  const patch = diffDocumentStates(current, next);
  const commit = await store.createCommit({
    docId,
    parentCommitId: main.headCommitId,
    mergeParentCommitId: null,
    createdBy: actor.userId,
    createdAt: 123,
    message: "test",
    patch,
    nextState: next,
  });

  const loaded = await store.getCommit(commit.id);
  assert.ok(loaded);
  assert.deepEqual(loaded.patch, patch);

  const rawCommit = doc.getMap("branching:commits").get(commit.id);
  assert.ok(rawCommit instanceof Y.Map);
  assert.equal(rawCommit.get("patch"), undefined);
  assert.equal(rawCommit.get("patchEncoding"), "gzip-chunks");

  const patchChunks = rawCommit.get("patchChunks");
  assert.ok(patchChunks instanceof Y.Array);
});
