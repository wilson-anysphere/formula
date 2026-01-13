import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";
import { randomUUID } from "node:crypto";
import { createRequire } from "node:module";

import * as Y from "yjs";

import { BranchService } from "../packages/versioning/branches/src/BranchService.js";
import { InMemoryBranchStore } from "../packages/versioning/branches/src/store/InMemoryBranchStore.js";
import {
  applyBranchStateToYjsDoc,
  branchStateFromYjsDoc,
} from "../packages/versioning/branches/src/browser.js";
import { createCollabSession } from "../packages/collab/session/src/index.ts";
import {
  getAvailablePort,
  startSyncServer,
  waitForCondition,
} from "../services/sync-server/test/test-helpers.ts";

test("sync-server + BranchService (Yjs): merge preserves sheet metadata + namedRanges + metadata (+ comments) and survives restart", async (t) => {
  // Sync-server-backed integration tests can be slower/flakier under heavy load.
  // Prefer a longer timeout here so we don't fail spuriously when CI hosts are busy.
  const TIMEOUT_MS = 60_000;
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-branching-"));
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const port = await getAvailablePort();
  // Resolve deps from the sync-server package so this test doesn't need to add them to the root package.json.
  const requireFromSyncServer = createRequire(new URL("../services/sync-server/package.json", import.meta.url));
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

  const createClient = ({ wsUrl: wsBaseUrl, docId, token }) => {
    const session = createCollabSession({
      connection: {
        wsUrl: wsBaseUrl,
        docId,
        token,
        WebSocketPolyfill: WebSocket,
        disableBc: true,
      },
      defaultSheetId: "Sheet1",
      // This test asserts BranchService + branchStateAdapter behavior. CollabSession's
      // workbook schema auto-init can race with large sheet list replacements
      // (delete + reinsert) and create spurious default sheets under load, making
      // this test flaky. Disable it here so the Yjs document contents are only
      // mutated by the adapter under test.
      schema: { autoInit: false },
    });

    const ydoc = session.doc;

    let destroyed = false;
    const destroy = () => {
      if (destroyed) return;
      destroyed = true;
      session.destroy();
      ydoc.destroy();
    };

    return { session, ydoc, destroy };
  };

  const docId = `e2e-branching-${randomUUID()}`;
  const wsUrl = server.wsUrl;

  const clientA = createClient({ wsUrl, docId, token: "test-token" });
  const clientB = createClient({ wsUrl, docId, token: "test-token" });
  t.after(() => {
    clientA.destroy();
    clientB.destroy();
  });

  await Promise.all([clientA.session.whenSynced(), clientB.session.whenSynced()]);

  // Initialize a minimal workbook.
  clientA.ydoc.transact(() => {
    const sheets = clientA.ydoc.getArray("sheets");
    if (sheets.length > 0) sheets.delete(0, sheets.length);
    const sheet1 = new Y.Map();
    sheet1.set("id", "Sheet1");
    sheet1.set("name", "Sheet1");
    sheets.push([sheet1]);
  }, clientA.session.origin);

  await waitForCondition(() => {
    const stateB = branchStateFromYjsDoc(clientB.ydoc);
    return stateB.sheets.order.length === 1 && stateB.sheets.metaById.Sheet1?.name === "Sheet1";
  }, TIMEOUT_MS);

  const actor = { userId: "u1", role: "owner" };
  const store = new InMemoryBranchStore();
  const branchService = new BranchService({ docId, store });
  await branchService.init(actor, branchStateFromYjsDoc(clientA.ydoc));

  await branchService.createBranch(actor, { name: "feature" });

  // --- Feature branch edits (offline/state-only) ---
  const featureBase = await branchService.checkoutBranch(actor, { name: "feature" });
  const featureNext = structuredClone(featureBase);
  featureNext.sheets.metaById.Sheet1.name = "FeatureName";
  featureNext.sheets.metaById.Sheet1.visibility = "hidden";
  featureNext.sheets.metaById.Sheet1.tabColor = "FF00FF00";
  featureNext.sheets.metaById.Sheet2 = { id: "Sheet2", name: "AddedSheet" };
  featureNext.sheets.metaById.Sheet2.visibility = "veryHidden";
  featureNext.sheets.metaById.Sheet2.tabColor = "FFFF0000";
  featureNext.cells.Sheet2 = {};
  featureNext.sheets.order = ["Sheet1", "Sheet2"];
  featureNext.metadata.title = "Budget";
  featureNext.namedRanges.NR1 = { sheetId: "Sheet1", rect: { r0: 0, c0: 0, r1: 0, c1: 0 } };
  featureNext.comments.c1 = { id: "c1", cellRef: "A1", content: "hello", resolved: false, replies: [] };
  await branchService.commit(actor, { nextState: featureNext, message: "feature edits" });

  // --- Main branch edits ---
  const mainBase = await branchService.checkoutBranch(actor, { name: "main" });
  const mainNext = structuredClone(mainBase);
  mainNext.sheets.metaById.Sheet1.name = "MainName";
  await branchService.commit(actor, { nextState: mainNext, message: "main rename" });

  const preview = await branchService.previewMerge(actor, { sourceBranch: "feature" });
  assert.equal(preview.conflicts.length, 1);
  assert.deepEqual(preview.conflicts[0], {
    type: "sheet",
    reason: "rename",
    sheetId: "Sheet1",
    base: "Sheet1",
    ours: "MainName",
    theirs: "FeatureName",
  });

  const merge = await branchService.merge(actor, {
    sourceBranch: "feature",
    resolutions: [{ conflictIndex: 0, choice: "theirs" }],
    message: "merge feature",
  });

  // Apply merge result back into the shared Yjs document.
  //
  // Use the dedicated "branching-apply" origin so this bulk rewrite is not
  // treated as a normal local edit by collaborative undo/conflict tracking.
  applyBranchStateToYjsDoc(clientA.ydoc, merge.state, { origin: "branching-apply" });

  await waitForCondition(() => {
    const stateB = branchStateFromYjsDoc(clientB.ydoc);
    return (
      stateB.sheets.order.join(",") === "Sheet1,Sheet2" &&
      stateB.sheets.metaById.Sheet1?.name === "FeatureName" &&
      stateB.sheets.metaById.Sheet1?.visibility === "hidden" &&
      stateB.sheets.metaById.Sheet1?.tabColor === "FF00FF00" &&
      stateB.sheets.metaById.Sheet2?.name === "AddedSheet" &&
      stateB.sheets.metaById.Sheet2?.visibility === "veryHidden" &&
      stateB.sheets.metaById.Sheet2?.tabColor === "FFFF0000" &&
      stateB.metadata.title === "Budget" &&
      stateB.namedRanges.NR1?.sheetId === "Sheet1" &&
      stateB.comments.c1?.content === "hello"
    );
  }, TIMEOUT_MS);

  // Tear down clients and restart the server, keeping the same data directory.
  clientA.destroy();
  clientB.destroy();

  // Give the server a moment to persist state after the last client disconnects.
  await new Promise((r) => setTimeout(r, 500));

  await server.stop();
  server = await startSyncServer({
    port,
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
  });

  const clientC = createClient({ wsUrl, docId, token: "test-token" });
  t.after(() => clientC.destroy());
  await clientC.session.whenSynced();

  // Note: sync-server persistence is loaded asynchronously on the server (y-websocket
  // does not await `bindState()`), so `whenSynced()` only guarantees the client
  // finished the initial sync protocol. Persisted updates may arrive shortly after.
  await waitForCondition(() => {
    const stateC = branchStateFromYjsDoc(clientC.ydoc);
    return (
      stateC.sheets.order.join(",") === "Sheet1,Sheet2" &&
      stateC.sheets.metaById.Sheet1?.name === "FeatureName" &&
      stateC.sheets.metaById.Sheet1?.visibility === "hidden" &&
      stateC.sheets.metaById.Sheet1?.tabColor === "FF00FF00" &&
      stateC.sheets.metaById.Sheet2?.name === "AddedSheet" &&
      stateC.sheets.metaById.Sheet2?.visibility === "veryHidden" &&
      stateC.sheets.metaById.Sheet2?.tabColor === "FFFF0000" &&
      stateC.metadata.title === "Budget" &&
      stateC.namedRanges.NR1?.sheetId === "Sheet1" &&
      stateC.comments.c1?.content === "hello"
    );
  }, TIMEOUT_MS);
});
