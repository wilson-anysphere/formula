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

  // Init creates the root commit + main branch in the shared Y.Doc.
  await branchService.init(owner, { sheets: {} });

  // --- Commit initial state ---
  sessionA.setCellValue("Sheet1:0:0", "base");
  sessionA.setCellFormula("Sheet1:0:1", "=1+1");
  await workflow.commitCurrentState(owner, "initial");

  // --- Branch + divergent edits ---
  await branchService.createBranch(owner, { name: "feature" });
  await workflow.checkoutBranch(owner, { name: "feature" });
  sessionA.setCellValue("Sheet1:0:0", "feature");
  await workflow.commitCurrentState(owner, "feature edit");

  await workflow.checkoutBranch(owner, { name: "main" });
  sessionA.setCellValue("Sheet1:0:0", "main");
  await workflow.commitCurrentState(owner, "main edit");

  const preview = await workflow.previewMerge(owner, { sourceBranch: "feature" });
  assert.equal(preview.conflicts.length, 1);
  assert.equal(preview.conflicts[0].type, "cell");
  assert.equal(preview.conflicts[0].sheetId, "Sheet1");
  assert.equal(preview.conflicts[0].cell, "A1");

  await workflow.merge(owner, {
    sourceBranch: "feature",
    resolutions: [{ conflictIndex: 0, choice: "theirs" }],
    message: "merge feature into main",
  });

  // --- Workbook state propagates to other collaborators ---
  await waitForCondition(() => sessionB.getCell("Sheet1:0:0")?.value === "feature", 10_000);
  assert.equal(sessionB.getCell("Sheet1:0:0")?.value, "feature");

  // --- Branch metadata propagates too (stored in the same Y.Doc) ---
  await waitForCondition(() => {
    const branches = sessionB.doc.getMap("branching:branches");
    const commits = sessionB.doc.getMap("branching:commits");
    return branches.has("main") && branches.has("feature") && commits.size >= 5;
  }, 10_000);

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

  // --- Branch metadata persisted inside the Y.Doc ---
  const branches = sessionC.doc.getMap("branching:branches");
  const commits = sessionC.doc.getMap("branching:commits");
  const meta = sessionC.doc.getMap("branching:meta");

  assert.ok(branches.has("main"));
  assert.ok(branches.has("feature"));
  assert.ok(typeof meta.get("rootCommitId") === "string");
  assert.ok(commits.size >= 5);
});
