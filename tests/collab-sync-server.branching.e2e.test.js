import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";
import { randomUUID } from "node:crypto";
import { createRequire } from "node:module";

import { createCollabSession } from "../packages/collab/session/src/index.ts";
import { CollabBranchingWorkflow } from "../packages/collab/branching/index.js";
import { BranchService, YjsBranchStore } from "../packages/versioning/branches/src/index.js";
import {
  getAvailablePort,
  startSyncServer,
  waitForCondition,
} from "../services/sync-server/test/test-helpers.ts";

/**
 * @param {unknown} value
 */
function getYMap(value) {
  if (!value || typeof value !== "object") return null;
  const maybe = value;
  if (maybe.constructor?.name !== "YMap") return null;
  if (typeof maybe.get !== "function") return null;
  if (typeof maybe.set !== "function") return null;
  if (typeof maybe.delete !== "function") return null;
  return maybe;
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
  sessionA.setCellValue("Sheet1:0:0", "base");
  sessionA.setCellFormula("Sheet1:0:1", "=1+1");
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
  sessionB.setCellValue("Sheet1:0:2", 99);
  const editorCommit = await workflowB.commitCurrentState(editor, "editor: add C1");

  // Ensure client A observes the branch head moving before making further commits.
  await waitForCondition(() => {
    const branches = sessionA.doc.getMap("branching:branches");
    const feature = getYMap(branches.get("feature"));
    return feature?.get("headCommitId") === editorCommit.id;
  }, 10_000);
  await waitForCondition(() => sessionA.getCell("Sheet1:0:2")?.value === 99, 10_000);

  sessionA.setCellValue("Sheet1:0:0", "feature");
  sessionA.doc.transact(() => {
    const cell = getYMap(sessionA.cells.get("Sheet1:0:1"));
    if (cell) cell.set("format", { numberFormat: "percent" });
  }, sessionA.origin);
  await workflow.commitCurrentState(owner, "feature edit");

  await workflow.checkoutBranch(owner, { name: "main" });
  sessionA.setCellValue("Sheet1:0:0", "main");
  sessionA.doc.transact(() => {
    const cell = getYMap(sessionA.cells.get("Sheet1:0:1"));
    if (cell) cell.set("format", { numberFormat: "accounting" });
  }, sessionA.origin);
  await workflow.commitCurrentState(owner, "main edit");

  const preview = await workflow.previewMerge(owner, { sourceBranch: "feature" });
  assert.equal(preview.conflicts.length, 2);
  const a1Idx = preview.conflicts.findIndex(
    (c) => c.type === "cell" && c.sheetId === "Sheet1" && c.cell === "A1"
  );
  const b1Idx = preview.conflicts.findIndex(
    (c) => c.type === "cell" && c.sheetId === "Sheet1" && c.cell === "B1"
  );
  assert.ok(a1Idx >= 0);
  assert.ok(b1Idx >= 0);
  assert.equal(preview.conflicts[a1Idx]?.reason, "content");
  assert.equal(preview.conflicts[b1Idx]?.reason, "format");

  await workflow.merge(owner, {
    sourceBranch: "feature",
    resolutions: [
      { conflictIndex: a1Idx, choice: "theirs" },
      { conflictIndex: b1Idx, choice: "theirs" },
    ],
    message: "merge feature into main",
  });

  // --- Workbook state propagates to other collaborators ---
  await waitForCondition(() => sessionB.getCell("Sheet1:0:0")?.value === "feature", 10_000);
  assert.equal(sessionB.getCell("Sheet1:0:0")?.value, "feature");
  await waitForCondition(() => sessionB.getCell("Sheet1:0:1")?.formula === "=1+1", 10_000);
  assert.equal(sessionB.getCell("Sheet1:0:1")?.formula, "=1+1");
  await waitForCondition(() => sessionB.getCell("Sheet1:0:2")?.value === 99, 10_000);
  assert.equal(sessionB.getCell("Sheet1:0:2")?.value, 99);
  await waitForCondition(() => {
    const cell = getYMap(sessionB.cells.get("Sheet1:0:1"));
    return cell?.get("format")?.numberFormat === "percent";
  }, 10_000);
  assert.deepEqual(getYMap(sessionB.cells.get("Sheet1:0:1"))?.get("format"), {
    numberFormat: "percent",
  });

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
  assert.equal(sessionC.getCell("Sheet1:0:0")?.value, "feature");
  assert.equal(sessionC.getCell("Sheet1:0:1")?.formula, "=1+1");
  assert.equal(sessionC.getCell("Sheet1:0:2")?.value, 99);
  assert.deepEqual(getYMap(sessionC.cells.get("Sheet1:0:1"))?.get("format"), {
    numberFormat: "percent",
  });

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
