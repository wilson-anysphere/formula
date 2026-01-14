import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";
import { randomUUID } from "node:crypto";
import { createRequire } from "node:module";

import { createCollabSession } from "../packages/collab/session/src/index.ts";
import { CollabBranchingWorkflow } from "../packages/collab/branching/index.js";
import { getYMap } from "../packages/collab/yjs-utils/src/index.ts";
import { BranchService, YjsBranchStore } from "../packages/versioning/branches/src/index.js";
import {
  getAvailablePort,
  startSyncServer,
  waitForCondition,
} from "../services/sync-server/test/test-helpers.ts";

function readYMapOrObject(value, key) {
  const map = getYMap(value);
  if (map) return map.get(key);
  if (!value || typeof value !== "object" || Array.isArray(value)) return undefined;
  return value[key];
}

function sheetNameFromDoc(doc, sheetId) {
  const sheets = doc.getArray("sheets");
  for (const entry of sheets.toArray()) {
    const id = readYMapOrObject(entry, "id");
    if (id !== sheetId) continue;
    return readYMapOrObject(entry, "name") ?? null;
  }
  return null;
}

function commentContentFromDoc(doc, commentId) {
  const comments = doc.getMap("comments");
  const value = comments.get(commentId);
  const map = getYMap(value);
  if (map) return map.get("content") ?? null;
  if (!value || typeof value !== "object") return null;
  return value.content ?? null;
}

/**
 * @param {import("../packages/collab/session/src/index.ts").CollabSession} session
 */
function makeDestroySession(session) {
  let destroyed = false;
  return () => {
    if (destroyed) return;
    destroyed = true;
    session.destroy();
    session.doc.destroy();
  };
}

test("sync-server + collab branching: Yjs-backed branches/commits + checkout/merge + persistence", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-branching-e2e-"));
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const port = await getAvailablePort();
  const requireFromSyncServer = createRequire(
    new URL("../services/sync-server/package.json", import.meta.url)
  );
  const WebSocket = requireFromSyncServer("ws");

  /** @type {Awaited<ReturnType<typeof startSyncServer>> | null} */
  let server = await startSyncServer({
    port,
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
  });
  t.after(async () => {
    await server?.stop();
  });

  const docId = `branching-e2e-${randomUUID()}`;
  const wsUrl = server.wsUrl;

  const sessionA = createCollabSession({
    connection: {
      wsUrl,
      docId,
      token: "test-token",
      WebSocketPolyfill: WebSocket,
      disableBc: true,
    },
    defaultSheetId: "Sheet1",
  });
  const sessionB = createCollabSession({
    connection: {
      wsUrl,
      docId,
      token: "test-token",
      WebSocketPolyfill: WebSocket,
      disableBc: true,
    },
    defaultSheetId: "Sheet1",
  });

  const destroyA = makeDestroySession(sessionA);
  const destroyB = makeDestroySession(sessionB);

  t.after(() => {
    destroyA();
    destroyB();
  });

  await Promise.all([sessionA.whenSynced(), sessionB.whenSynced()]);

  const owner = { userId: "u-a", role: "owner" };

  const store = new YjsBranchStore({ ydoc: sessionA.doc });
  const branchService = new BranchService({ docId, store });
  const workflow = new CollabBranchingWorkflow({ session: sessionA, branchService });

  const storeB = new YjsBranchStore({ ydoc: sessionB.doc });
  const branchServiceB = new BranchService({ docId, store: storeB });
  const workflowB = new CollabBranchingWorkflow({ session: sessionB, branchService: branchServiceB });

  // Init creates the root commit + main branch in the shared Y.Doc.
  await branchService.init(owner, { sheets: {} });

  // --- Commit initial state ---
  await sessionA.setCellValue("Sheet1:0:0", "base");
  await sessionA.setCellFormula("Sheet1:0:1", "=1+1");
  await workflow.commitCurrentState(owner, "initial");

  // --- Branch + divergent edits ---
  await branchService.createBranch(owner, { name: "feature" });
  await workflow.checkoutBranch(owner, { name: "feature" });

  // Editor can commit on the globally checked-out branch without being allowed to checkout.
  const editor = { userId: "u-b", role: "editor" };
  await waitForCondition(
    () => sessionB.doc.getMap("branching:meta").get("currentBranchName") === "feature",
    10_000
  );
  await sessionB.setCellValue("Sheet1:0:2", 99);
  const editorCommit = await workflowB.commitCurrentState(editor, "editor: add C1");

  // Ensure client A observes the branch head moving before making further commits.
  await waitForCondition(() => {
    const branches = sessionA.doc.getMap("branching:branches");
    const feature = getYMap(branches.get("feature"));
    return feature?.get("headCommitId") === editorCommit.id;
  }, 10_000);
  await waitForCondition(async () => (await sessionA.getCell("Sheet1:0:2"))?.value === 99, 10_000);

  await sessionA.setCellValue("Sheet1:0:0", "feature");
  sessionA.doc.transact(() => {
    const cell = getYMap(sessionA.cells.get("Sheet1:0:1"));
    if (cell) cell.set("format", { numberFormat: "percent" });
    const sheets = sessionA.doc.getArray("sheets");
    if (sheets.length > 0) sheets.delete(0, sheets.length);
    sheets.push([{ id: "Sheet1", name: "FeatureName" }]);
    sessionA.doc.getMap("metadata").set("scenario", "feature");
    sessionA.doc.getMap("namedRanges").set("NR1", {
      sheetId: "Sheet1",
      rect: { r0: 0, c0: 0, r1: 0, c1: 0 },
    });
    sessionA.doc.getMap("comments").set("c1", {
      id: "c1",
      cellRef: "A1",
      content: "feature comment",
      resolved: false,
      replies: [],
    });
  }, sessionA.origin);
  await workflow.commitCurrentState(owner, "feature edit");

  await workflow.checkoutBranch(owner, { name: "main" });
  await sessionA.setCellValue("Sheet1:0:0", "main");
  sessionA.doc.transact(() => {
    const cell = getYMap(sessionA.cells.get("Sheet1:0:1"));
    if (cell) cell.set("format", { numberFormat: "accounting" });
    const sheets = sessionA.doc.getArray("sheets");
    if (sheets.length > 0) sheets.delete(0, sheets.length);
    sheets.push([{ id: "Sheet1", name: "MainName" }]);
    sessionA.doc.getMap("metadata").set("scenario", "main");
    sessionA.doc.getMap("namedRanges").set("NR1", {
      sheetId: "Sheet1",
      rect: { r0: 0, c0: 1, r1: 0, c1: 1 },
    });
    sessionA.doc.getMap("comments").set("c1", {
      id: "c1",
      cellRef: "A1",
      content: "main comment",
      resolved: false,
      replies: [],
    });
  }, sessionA.origin);
  await workflow.commitCurrentState(owner, "main edit");

  const preview = await workflow.previewMerge(owner, { sourceBranch: "feature" });
  assert.equal(preview.conflicts.length, 6);
  const a1Idx = preview.conflicts.findIndex(
    (c) => c.type === "cell" && c.sheetId === "Sheet1" && c.cell === "A1"
  );
  const b1Idx = preview.conflicts.findIndex(
    (c) => c.type === "cell" && c.sheetId === "Sheet1" && c.cell === "B1"
  );
  const sheetIdx = preview.conflicts.findIndex(
    (c) => c.type === "sheet" && c.reason === "rename" && c.sheetId === "Sheet1"
  );
  const metadataIdx = preview.conflicts.findIndex((c) => c.type === "metadata" && c.key === "scenario");
  const namedRangeIdx = preview.conflicts.findIndex((c) => c.type === "namedRange" && c.key === "NR1");
  const commentIdx = preview.conflicts.findIndex((c) => c.type === "comment" && c.id === "c1");
  assert.ok(a1Idx >= 0);
  assert.ok(b1Idx >= 0);
  assert.ok(sheetIdx >= 0);
  assert.ok(metadataIdx >= 0);
  assert.ok(namedRangeIdx >= 0);
  assert.ok(commentIdx >= 0);
  assert.equal(preview.conflicts[a1Idx]?.reason, "content");
  assert.equal(preview.conflicts[b1Idx]?.reason, "format");
  assert.equal(preview.conflicts[sheetIdx]?.reason, "rename");

  await workflow.merge(owner, {
    sourceBranch: "feature",
    resolutions: [
      { conflictIndex: a1Idx, choice: "theirs" },
      { conflictIndex: b1Idx, choice: "theirs" },
      { conflictIndex: sheetIdx, choice: "theirs" },
      { conflictIndex: metadataIdx, choice: "theirs" },
      { conflictIndex: namedRangeIdx, choice: "theirs" },
      { conflictIndex: commentIdx, choice: "theirs" },
    ],
    message: "merge feature into main",
  });

  // --- Workbook state propagates to other collaborators ---
  await waitForCondition(async () => (await sessionB.getCell("Sheet1:0:0"))?.value === "feature", 10_000);
  assert.equal((await sessionB.getCell("Sheet1:0:0"))?.value, "feature");
  await waitForCondition(async () => (await sessionB.getCell("Sheet1:0:1"))?.formula === "=1+1", 10_000);
  assert.equal((await sessionB.getCell("Sheet1:0:1"))?.formula, "=1+1");
  await waitForCondition(async () => (await sessionB.getCell("Sheet1:0:2"))?.value === 99, 10_000);
  assert.equal((await sessionB.getCell("Sheet1:0:2"))?.value, 99);
  await waitForCondition(() => {
    const cell = getYMap(sessionB.cells.get("Sheet1:0:1"));
    return cell?.get("format")?.numberFormat === "percent";
  }, 10_000);
  assert.deepEqual(getYMap(sessionB.cells.get("Sheet1:0:1"))?.get("format"), {
    numberFormat: "percent",
  });
  await waitForCondition(() => sheetNameFromDoc(sessionB.doc, "Sheet1") === "FeatureName", 10_000);
  assert.equal(sheetNameFromDoc(sessionB.doc, "Sheet1"), "FeatureName");
  await waitForCondition(() => {
    const nr = sessionB.doc.getMap("namedRanges").get("NR1");
    return nr?.rect?.c0 === 0;
  }, 10_000);
  assert.deepEqual(sessionB.doc.getMap("namedRanges").get("NR1"), {
    rect: { c0: 0, c1: 0, r0: 0, r1: 0 },
    sheetId: "Sheet1",
  });
  await waitForCondition(() => commentContentFromDoc(sessionB.doc, "c1") === "feature comment", 10_000);
  assert.equal(commentContentFromDoc(sessionB.doc, "c1"), "feature comment");
  await waitForCondition(() => sessionB.doc.getMap("metadata").get("scenario") === "feature", 10_000);
  assert.equal(sessionB.doc.getMap("metadata").get("scenario"), "feature");

  // --- Branch metadata propagates too (stored in the same Y.Doc) ---
  await waitForCondition(() => {
    const branches = sessionB.doc.getMap("branching:branches");
    const commits = sessionB.doc.getMap("branching:commits");
    return branches.has("main") && branches.has("feature") && commits.size >= 6;
  }, 10_000);

  // Validate metadata via BranchService on another client.
  const branchesB = await branchServiceB.listBranches();
  assert.deepEqual(
    branchesB.map((b) => b.name).sort(),
    ["feature", "main"]
  );
  await waitForCondition(async () => {
    const main = await storeB.getBranch(docId, "main");
    if (!main) return false;
    const head = await storeB.getCommit(main.headCommitId);
    return head?.message === "merge feature into main";
  }, 10_000);
  const mainB = await storeB.getBranch(docId, "main");
  assert.ok(mainB);
  const headCommitB = await storeB.getCommit(mainB.headCommitId);
  assert.equal(headCommitB?.message, "merge feature into main");

  // Tear down sessions and restart sync-server with the same data dir.
  destroyA();
  destroyB();

  // Allow time for persistence after disconnect.
  await new Promise((r) => setTimeout(r, 500));

  await server.stop();
  server = await startSyncServer({
    port,
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
  });

  const sessionC = createCollabSession({
    connection: {
      wsUrl,
      docId,
      token: "test-token",
      WebSocketPolyfill: WebSocket,
      disableBc: true,
    },
    defaultSheetId: "Sheet1",
  });
  const destroyC = makeDestroySession(sessionC);
  t.after(() => {
    destroyC();
  });

  await sessionC.whenSynced();

  // --- Workbook state persisted ---
  // Note: sync-server persistence is loaded asynchronously on the server (y-websocket
  // does not await `bindState()`), so `whenSynced()` only guarantees the client
  // finished the initial sync protocol. Persisted updates may arrive shortly after.
  await waitForCondition(async () => (await sessionC.getCell("Sheet1:0:0"))?.value === "feature", 10_000);
  assert.equal((await sessionC.getCell("Sheet1:0:0"))?.value, "feature");
  await waitForCondition(async () => (await sessionC.getCell("Sheet1:0:1"))?.formula === "=1+1", 10_000);
  assert.equal((await sessionC.getCell("Sheet1:0:1"))?.formula, "=1+1");
  await waitForCondition(async () => (await sessionC.getCell("Sheet1:0:2"))?.value === 99, 10_000);
  assert.equal((await sessionC.getCell("Sheet1:0:2"))?.value, 99);
  assert.deepEqual(getYMap(sessionC.cells.get("Sheet1:0:1"))?.get("format"), {
    numberFormat: "percent",
  });
  assert.equal(sheetNameFromDoc(sessionC.doc, "Sheet1"), "FeatureName");
  assert.deepEqual(sessionC.doc.getMap("namedRanges").get("NR1"), {
    rect: { c0: 0, c1: 0, r0: 0, r1: 0 },
    sheetId: "Sheet1",
  });
  assert.equal(commentContentFromDoc(sessionC.doc, "c1"), "feature comment");
  await waitForCondition(() => sessionC.doc.getMap("metadata").get("scenario") === "feature", 10_000);
  assert.equal(sessionC.doc.getMap("metadata").get("scenario"), "feature");

  // --- Branch metadata persisted inside the Y.Doc ---
  const branches = sessionC.doc.getMap("branching:branches");
  const commits = sessionC.doc.getMap("branching:commits");
  const meta = sessionC.doc.getMap("branching:meta");

  assert.ok(branches.has("main"));
  assert.ok(branches.has("feature"));
  assert.ok(typeof meta.get("rootCommitId") === "string");
  assert.equal(meta.get("currentBranchName"), "main");
  assert.ok(commits.size >= 6);

  // Validate persisted metadata via store API too.
  const storeC = new YjsBranchStore({ ydoc: sessionC.doc });
  const branchServiceC = new BranchService({ docId, store: storeC });
  const branchesC = await branchServiceC.listBranches();
  assert.deepEqual(
    branchesC.map((b) => b.name).sort(),
    ["feature", "main"]
  );
});
