import test from "node:test";
import assert from "node:assert/strict";
import * as Y from "yjs";

import { YjsBranchStore } from "../packages/versioning/branches/src/store/YjsBranchStore.js";

test("YjsBranchStore.ensureDocument ignores incomplete gzip-chunks commits when inferring root/head", async () => {
  const ydoc = new Y.Doc();
  const store = new YjsBranchStore({
    ydoc,
    payloadEncoding: "gzip-chunks",
    // Force multi-transaction writes for root commits (patch + snapshot)
    maxChunksPerTransaction: 1,
  });
  const docId = "doc1";
  const actor = { userId: "u1", role: "owner" };

  await store.ensureDocument(docId, actor, { sheets: {} });

  const meta = ydoc.getMap("branching:meta");
  const branches = ydoc.getMap("branching:branches");
  const commits = ydoc.getMap("branching:commits");

  const originalRoot = meta.get("rootCommitId");
  assert.ok(typeof originalRoot === "string" && originalRoot.length > 0);

  const rootCommitMap = commits.get(originalRoot);
  assert.ok(rootCommitMap instanceof Y.Map);
  const baseCreatedAt = Number(rootCommitMap.get("createdAt") ?? Date.now());

  const incompleteRootId = "incomplete-root";
  const incompleteNewerId = "incomplete-newer";

  ydoc.transact(() => {
    // An incomplete *root* commit (parentCommitId=null) that is older than the real one.
    // This exercises root inference when `meta.rootCommitId` is missing.
    const incompleteRoot = new Y.Map();
    incompleteRoot.set("id", incompleteRootId);
    incompleteRoot.set("docId", docId);
    incompleteRoot.set("parentCommitId", null);
    incompleteRoot.set("mergeParentCommitId", null);
    incompleteRoot.set("createdBy", actor.userId);
    incompleteRoot.set("createdAt", baseCreatedAt - 1_000);
    incompleteRoot.set("message", "incomplete-root");
    incompleteRoot.set("patchEncoding", "gzip-chunks");
    incompleteRoot.set("commitComplete", false);
    incompleteRoot.set("patchChunks", new Y.Array());
    commits.set(incompleteRootId, incompleteRoot);

    // A newer incomplete commit that would previously be selected as the latest/head.
    const incompleteNewer = new Y.Map();
    incompleteNewer.set("id", incompleteNewerId);
    incompleteNewer.set("docId", docId);
    incompleteNewer.set("parentCommitId", originalRoot);
    incompleteNewer.set("mergeParentCommitId", null);
    incompleteNewer.set("createdBy", actor.userId);
    incompleteNewer.set("createdAt", baseCreatedAt + 1_000);
    incompleteNewer.set("message", "incomplete-newer");
    incompleteNewer.set("patchEncoding", "gzip-chunks");
    incompleteNewer.set("commitComplete", false);
    incompleteNewer.set("patchChunks", new Y.Array());
    commits.set(incompleteNewerId, incompleteNewer);

    // Corrupt metadata to force root/head inference.
    meta.delete("rootCommitId");
    branches.delete("main");
    meta.set("currentBranchName", "ghost");
  });

  await store.ensureDocument(docId, actor, { sheets: {} });

  assert.equal(meta.get("rootCommitId"), originalRoot);
  const main = branches.get("main");
  assert.ok(main instanceof Y.Map);
  assert.equal(main.get("headCommitId"), originalRoot);
  assert.equal(meta.get("currentBranchName"), "main");
});

test("YjsBranchStore.ensureDocument repairs meta.rootCommitId pointing at an incomplete commit", async () => {
  const ydoc = new Y.Doc();
  const store = new YjsBranchStore({
    ydoc,
    payloadEncoding: "gzip-chunks",
    maxChunksPerTransaction: 1,
  });
  const docId = "doc1";
  const actor = { userId: "u1", role: "owner" };

  await store.ensureDocument(docId, actor, { sheets: {} });

  const meta = ydoc.getMap("branching:meta");
  const commits = ydoc.getMap("branching:commits");

  const originalRoot = meta.get("rootCommitId");
  assert.ok(typeof originalRoot === "string" && originalRoot.length > 0);

  const rootCommitMap = commits.get(originalRoot);
  assert.ok(rootCommitMap instanceof Y.Map);
  const baseCreatedAt = Number(rootCommitMap.get("createdAt") ?? Date.now());

  const incompleteRootId = "incomplete-root-bad-meta";

  ydoc.transact(() => {
    const incompleteRoot = new Y.Map();
    incompleteRoot.set("id", incompleteRootId);
    incompleteRoot.set("docId", docId);
    incompleteRoot.set("parentCommitId", null);
    incompleteRoot.set("mergeParentCommitId", null);
    incompleteRoot.set("createdBy", actor.userId);
    incompleteRoot.set("createdAt", baseCreatedAt - 1_000);
    incompleteRoot.set("message", "incomplete-root");
    incompleteRoot.set("patchEncoding", "gzip-chunks");
    incompleteRoot.set("commitComplete", false);
    incompleteRoot.set("patchChunks", new Y.Array());
    commits.set(incompleteRootId, incompleteRoot);

    // Corrupt the rootCommitId pointer to simulate an interrupted write.
    meta.set("rootCommitId", incompleteRootId);
  });

  await store.ensureDocument(docId, actor, { sheets: {} });
  assert.equal(meta.get("rootCommitId"), originalRoot);
});

test("YjsBranchStore.ensureDocument can recover when root commit is mid snapshot migration (commitComplete=false but patch is inline)", async () => {
  const ydoc = new Y.Doc();
  const docId = "doc1";
  const actor = { userId: "u1", role: "owner" };

  // Seed history using JSON payloads so the root commit contains an inline patch.
  const seedStore = new YjsBranchStore({ ydoc, payloadEncoding: "json" });
  await seedStore.ensureDocument(docId, actor, { sheets: {} });

  const meta = ydoc.getMap("branching:meta");
  const commits = ydoc.getMap("branching:commits");
  const rootId = meta.get("rootCommitId");
  assert.ok(typeof rootId === "string" && rootId.length > 0);

  const rootCommit = commits.get(rootId);
  assert.ok(rootCommit instanceof Y.Map);
  assert.ok(rootCommit.get("patch") !== undefined);

  // Simulate an interrupted migration that tried to write a gzip-chunks snapshot
  // for the root commit but crashed mid-write.
  ydoc.transact(() => {
    rootCommit.delete("snapshot");
    rootCommit.set("commitComplete", false);
    rootCommit.set("snapshotEncoding", "gzip-chunks");
    const chunks = new Y.Array();
    chunks.push([new Uint8Array([1, 2, 3])]);
    rootCommit.set("snapshotChunks", chunks);
  });

  const repairStore = new YjsBranchStore({
    ydoc,
    payloadEncoding: "gzip-chunks",
    maxChunksPerTransaction: 1,
  });

  const stateBeforeRepair = await repairStore.getDocumentStateAtCommit(rootId);
  assert.equal(stateBeforeRepair.schemaVersion, 1);

  await repairStore.ensureDocument(docId, actor, { sheets: {} });

  const repaired = commits.get(rootId);
  assert.ok(repaired instanceof Y.Map);
  assert.equal(repaired.get("commitComplete"), true);
 });

test("YjsBranchStore.ensureDocument can infer rootCommitId even if root is mid snapshot migration (commitComplete=false but patch is inline)", async () => {
  const ydoc = new Y.Doc();
  const docId = "doc1";
  const actor = { userId: "u1", role: "owner" };

  // Seed history using JSON payloads so the root commit contains an inline patch.
  const seedStore = new YjsBranchStore({ ydoc, payloadEncoding: "json" });
  await seedStore.ensureDocument(docId, actor, { sheets: {} });

  const meta = ydoc.getMap("branching:meta");
  const commits = ydoc.getMap("branching:commits");
  const rootId = meta.get("rootCommitId");
  assert.ok(typeof rootId === "string" && rootId.length > 0);

  const rootCommit = commits.get(rootId);
  assert.ok(rootCommit instanceof Y.Map);

  const commitsBefore = commits.size;

  ydoc.transact(() => {
    // Simulate an interrupted migration that tried to write a gzip-chunks snapshot
    // for the root commit but crashed mid-write.
    rootCommit.delete("snapshot");
    rootCommit.set("commitComplete", false);
    rootCommit.set("snapshotEncoding", "gzip-chunks");
    const chunks = new Y.Array();
    chunks.push([new Uint8Array([1, 2, 3])]);
    rootCommit.set("snapshotChunks", chunks);

    // Also corrupt the rootCommitId metadata to force inference.
    meta.delete("rootCommitId");
  });

  const repairStore = new YjsBranchStore({
    ydoc,
    payloadEncoding: "gzip-chunks",
    maxChunksPerTransaction: 1,
  });

  await repairStore.ensureDocument(docId, actor, { sheets: {} });

  assert.equal(meta.get("rootCommitId"), rootId);
  assert.equal(commits.size, commitsBefore);

  const repaired = commits.get(rootId);
  assert.ok(repaired instanceof Y.Map);
  assert.equal(repaired.get("commitComplete"), true);
});

test("YjsBranchStore.ensureDocument deletes stale unreachable incomplete gzip-chunks commits", async () => {
  const ydoc = new Y.Doc();
  const store = new YjsBranchStore({
    ydoc,
    payloadEncoding: "gzip-chunks",
    maxChunksPerTransaction: 1,
  });
  const docId = "doc1";
  const actor = { userId: "u1", role: "owner" };

  await store.ensureDocument(docId, actor, { sheets: {} });

  const meta = ydoc.getMap("branching:meta");
  const commits = ydoc.getMap("branching:commits");
  const rootId = meta.get("rootCommitId");
  assert.ok(typeof rootId === "string" && rootId.length > 0);

  const staleId = "stale-incomplete";
  const now = Date.now();

  ydoc.transact(() => {
    const commit = new Y.Map();
    commit.set("id", staleId);
    commit.set("docId", docId);
    commit.set("parentCommitId", rootId);
    commit.set("mergeParentCommitId", null);
    commit.set("createdBy", actor.userId);
    commit.set("createdAt", now);
    commit.set("writeStartedAt", now - 2 * 60 * 60 * 1000);
    commit.set("message", "stale");
    commit.set("patchEncoding", "gzip-chunks");
    commit.set("commitComplete", false);
    commit.set("patchChunks", new Y.Array());
    commits.set(staleId, commit);
  });

  assert.ok(commits.has(staleId));

  await store.ensureDocument(docId, actor, { sheets: {} });

  assert.equal(commits.has(staleId), false);
});

test("YjsBranchStore.ensureDocument repairs main branch head when it points at an incomplete commit", async () => {
  const ydoc = new Y.Doc();
  const store = new YjsBranchStore({
    ydoc,
    payloadEncoding: "gzip-chunks",
    maxChunksPerTransaction: 1,
  });
  const docId = "doc1";
  const actor = { userId: "u1", role: "owner" };

  await store.ensureDocument(docId, actor, { sheets: {} });

  const meta = ydoc.getMap("branching:meta");
  const branches = ydoc.getMap("branching:branches");
  const commits = ydoc.getMap("branching:commits");
  const rootId = meta.get("rootCommitId");
  assert.ok(typeof rootId === "string" && rootId.length > 0);

  const rootCommit = commits.get(rootId);
  assert.ok(rootCommit instanceof Y.Map);
  const baseCreatedAt = Number(rootCommit.get("createdAt") ?? Date.now());

  const incompleteHeadId = "incomplete-head";

  ydoc.transact(() => {
    const incomplete = new Y.Map();
    incomplete.set("id", incompleteHeadId);
    incomplete.set("docId", docId);
    incomplete.set("parentCommitId", rootId);
    incomplete.set("mergeParentCommitId", null);
    incomplete.set("createdBy", actor.userId);
    incomplete.set("createdAt", baseCreatedAt + 1_000);
    incomplete.set("writeStartedAt", Date.now());
    incomplete.set("message", "incomplete");
    incomplete.set("patchEncoding", "gzip-chunks");
    incomplete.set("commitComplete", false);
    incomplete.set("patchChunks", new Y.Array());
    commits.set(incompleteHeadId, incomplete);

    const main = branches.get("main");
    assert.ok(main instanceof Y.Map);
    main.set("headCommitId", incompleteHeadId);
  });

  await store.ensureDocument(docId, actor, { sheets: {} });

  const main = branches.get("main");
  assert.ok(main instanceof Y.Map);
  assert.equal(main.get("headCommitId"), rootId);
});
