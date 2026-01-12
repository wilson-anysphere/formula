import test from "node:test";
import assert from "node:assert/strict";
import { randomUUID } from "node:crypto";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import WebSocket from "ws";

import {
  createMetadataManagerForSession,
  createNamedRangeManagerForSession,
  createSheetManagerForSession,
} from "@formula/collab-workbook";

import { createCollabSession } from "../src/index.ts";
import { createCollabVersioning } from "../../versioning/src/index.ts";
import { SQLiteVersionStore } from "../../../versioning/src/store/sqliteVersionStore.js";
import {
  getAvailablePort,
  startSyncServer,
  waitForCondition,
} from "../../../../services/sync-server/test/test-helpers.ts";

/**
 * @param {import("../src/index.ts").CollabSession} session
 */
function snapshotSheets(session) {
  return session.sheets.toArray().map((sheet) => ({
    id: String(sheet.get("id") ?? ""),
    name: sheet.get("name") == null ? null : String(sheet.get("name")),
    visibility: sheet.get("visibility") == null ? "visible" : String(sheet.get("visibility")),
    tabColor: sheet.get("tabColor") == null ? null : String(sheet.get("tabColor")),
  }));
}

test("CollabVersioning integration: restore syncs + persists (sync-server)", async (t) => {
  // Sync-server-backed integration tests can be slower/flakier under heavy CI
  // load; use a longer timeout than in-memory unit tests.
  const TIMEOUT_MS = 120_000;
  const dataDir = await mkdtemp(path.join(tmpdir(), "collab-versioning-sync-server-"));
  const versionDbPath = path.join(dataDir, "versions.sqlite");

  /** @type {import("../src/index.ts").CollabSession | null} */
  let sessionA = null;
  /** @type {import("../src/index.ts").CollabSession | null} */
  let sessionB = null;
  /** @type {import("../../../versioning/src/store/sqliteVersionStore.js").SQLiteVersionStore | null} */
  let store = null;
  /** @type {import("../../versioning/src/index.ts").CollabVersioning | null} */
  let versioning = null;
  /** @type {Awaited<ReturnType<typeof startSyncServer>> | null} */
  let server = null;

  t.after(async () => {
    versioning?.destroy();
    store?.close?.();
    sessionA?.destroy();
    sessionB?.destroy();
    sessionA?.doc.destroy();
    sessionB?.doc.destroy();
    await server?.stop();
    await rm(dataDir, { recursive: true, force: true });
  });

  const port = await getAvailablePort();
  server = await startSyncServer({
    port,
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: { SYNC_SERVER_PERSISTENCE_BACKEND: "file" },
  });

  // Use a unique doc id to avoid cross-test collisions when running multiple
  // suites against the same sync-server data dir (or when reusing tmp dirs in
  // local debugging).
  const docId = `collab-versioning-test-doc-${randomUUID()}`;
  const wsUrl = server.wsUrl;

  // Create sessions sequentially. This test is about versioning + restore behavior,
  // not concurrent schema initialization (covered separately in
  // `workbook-schema.concurrent-init.sync-server.test.js`).
  //
  // Starting both sessions at the same time can lead to rare flakiness under heavy
  // load where the default-sheet schema converges more slowly, which then causes
  // downstream waits to time out.
  sessionA = createCollabSession({
    connection: {
      wsUrl,
      docId,
      token: "test-token",
      WebSocketPolyfill: WebSocket,
      disableBc: true,
    },
  });
  await sessionA.whenSynced();

  sessionB = createCollabSession({
    connection: {
      wsUrl,
      docId,
      token: "test-token",
      WebSocketPolyfill: WebSocket,
      disableBc: true,
    },
  });
  await sessionB.whenSynced();

  // Ensure both sessions have converged to a stable schema before applying edits.
  await waitForCondition(() => snapshotSheets(sessionA).length === 1, TIMEOUT_MS);
  await waitForCondition(() => snapshotSheets(sessionB).length === 1, TIMEOUT_MS);

  store = new SQLiteVersionStore({ filePath: versionDbPath });
  versioning = createCollabVersioning({
    session: sessionA,
    store,
    user: { userId: "user-a", userName: "User A" },
    autoStart: false,
  });

  const sheetsA = createSheetManagerForSession(sessionA);
  const namedRangesA = createNamedRangeManagerForSession(sessionA);
  const metadataA = createMetadataManagerForSession(sessionA);

  // Workbook metadata that should be included in checkpoints/restores.
  sheetsA.addSheet({ id: "Sheet2", name: "Budget" });
  sheetsA.moveSheet("Sheet1", 1);
  sheetsA.setTabColor("Sheet2", "ff00ff00");
  sheetsA.setVisibility("Sheet2", "hidden");
  namedRangesA.set("MyRange", { sheetId: "Sheet2", range: "A1:B2" });
  metadataA.set("title", "Quarterly Budget");

  await waitForCondition(() => {
    const sheets = snapshotSheets(sessionB);
    if (sheets.length !== 2) return false;
    if (sheets[0]?.id !== "Sheet2" || sheets[0]?.name !== "Budget") return false;
    if (sheets[0]?.visibility !== "hidden") return false;
    if (sheets[0]?.tabColor !== "FF00FF00") return false;
    if (sheets[1]?.id !== "Sheet1") return false;
    const nr = sessionB.namedRanges.get("MyRange");
    if (nr?.sheetId !== "Sheet2" || nr?.range !== "A1:B2") return false;
    return sessionB.metadata.get("title") === "Quarterly Budget";
  }, TIMEOUT_MS);

  // Initial edits + checkpoint.
  await sessionA.setCellValue("Sheet1:0:0", "alpha");
  await sessionA.setCellValue("Sheet1:0:1", 123);
  await waitForCondition(async () => (await sessionB.getCell("Sheet1:0:0"))?.value === "alpha", TIMEOUT_MS);
  await waitForCondition(async () => (await sessionB.getCell("Sheet1:0:1"))?.value === 123, TIMEOUT_MS);

  const checkpoint = await versioning.createCheckpoint({ name: "checkpoint-1" });

  // More edits (including a new cell that should be deleted on restore).
  await sessionA.setCellValue("Sheet1:0:0", "beta");
  await sessionA.setCellValue("Sheet1:2:0", "extra");
  sheetsA.renameSheet("Sheet2", "Budget Updated");
  sheetsA.removeSheet("Sheet2");
  namedRangesA.delete("MyRange");
  metadataA.set("title", "After");
  await waitForCondition(async () => (await sessionB.getCell("Sheet1:0:0"))?.value === "beta", TIMEOUT_MS);
  await waitForCondition(async () => (await sessionB.getCell("Sheet1:2:0"))?.value === "extra", TIMEOUT_MS);
  await waitForCondition(() => snapshotSheets(sessionB).length === 1, TIMEOUT_MS);
  await waitForCondition(() => sessionB.namedRanges.has("MyRange") === false, TIMEOUT_MS);
  await waitForCondition(() => sessionB.metadata.get("title") === "After", TIMEOUT_MS);

  // Restore the checkpoint and ensure the other collaborator converges.
  await versioning.restoreVersion(checkpoint.id);

  await waitForCondition(async () => (await sessionB.getCell("Sheet1:0:0"))?.value === "alpha", TIMEOUT_MS);
  await waitForCondition(async () => (await sessionB.getCell("Sheet1:0:1"))?.value === 123, TIMEOUT_MS);
  await waitForCondition(async () => (await sessionB.getCell("Sheet1:2:0")) == null, TIMEOUT_MS);
  await waitForCondition(() => {
     const sheets = snapshotSheets(sessionB);
     if (sheets.length !== 2) return false;
     if (sheets[0]?.id !== "Sheet2" || sheets[0]?.name !== "Budget") return false;
     if (sheets[0]?.visibility !== "hidden") return false;
     if (sheets[0]?.tabColor !== "FF00FF00") return false;
    if (sheets[1]?.id !== "Sheet1") return false;
    const nr = sessionB.namedRanges.get("MyRange");
    if (nr?.sheetId !== "Sheet2" || nr?.range !== "A1:B2") return false;
    return sessionB.metadata.get("title") === "Quarterly Budget";
  }, TIMEOUT_MS);

  assert.equal((await sessionB.getCell("Sheet1:0:0"))?.value, "alpha");
  assert.equal((await sessionB.getCell("Sheet1:0:1"))?.value, 123);
  assert.equal(await sessionB.getCell("Sheet1:2:0"), null);

  const versions = await versioning.listVersions();
  assert.ok(versions.some((v) => v.kind === "checkpoint" && v.id === checkpoint.id));
  assert.ok(versions.some((v) => v.kind === "restore"));

  // Store persistence: re-open and list versions.
  store.close();
  const reopened = new SQLiteVersionStore({ filePath: versionDbPath });
  const persisted = await reopened.listVersions();
  assert.ok(persisted.some((v) => v.kind === "checkpoint" && v.id === checkpoint.id));
  assert.ok(persisted.some((v) => v.kind === "restore"));
  reopened.close();
});
