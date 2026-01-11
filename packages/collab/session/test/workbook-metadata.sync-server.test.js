import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { randomUUID } from "node:crypto";

import WebSocket from "ws";

import {
  createMetadataManagerForSession,
  createNamedRangeManagerForSession,
  createSheetManagerForSession,
} from "@formula/collab-workbook";

import { createCollabSession } from "../src/index.ts";
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
  }));
}

test("CollabSession workbook metadata persists via sync-server (sheets + namedRanges + metadata)", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "collab-session-workbook-metadata-"));

  /** @type {import("../src/index.ts").CollabSession | null} */
  let sessionA = null;
  /** @type {import("../src/index.ts").CollabSession | null} */
  let sessionB = null;
  /** @type {import("../src/index.ts").CollabSession | null} */
  let sessionC = null;
  /** @type {Awaited<ReturnType<typeof startSyncServer>> | null} */
  let server = null;

  t.after(async () => {
    sessionA?.destroy();
    sessionB?.destroy();
    sessionC?.destroy();
    sessionA?.doc.destroy();
    sessionB?.doc.destroy();
    sessionC?.doc.destroy();
    await server?.stop();
    await rm(dataDir, { recursive: true, force: true });
  });

  const port = await getAvailablePort();
  server = await startSyncServer({
    port,
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
  });

  const docId = `collab-session-workbook-metadata-${randomUUID()}`;
  const wsUrl = server.wsUrl;

  // Create A first so schema initialization happens exactly once for brand new docs.
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

  const sheetsA = createSheetManagerForSession(sessionA);
  const namedRangesA = createNamedRangeManagerForSession(sessionA);
  const metadataA = createMetadataManagerForSession(sessionA);

  sheetsA.addSheet({ id: "Sheet2", name: "Budget" });
  sheetsA.moveSheet("Sheet1", 1);
  namedRangesA.set("MyRange", { sheetId: "Sheet2", range: "A1:B2" });
  metadataA.set("title", "Quarterly Budget");

  await waitForCondition(() => {
    const bSheets = snapshotSheets(sessionB);
    if (bSheets.length !== 2) return false;
    if (bSheets[0]?.id !== "Sheet2" || bSheets[0]?.name !== "Budget") return false;
    if (bSheets[1]?.id !== "Sheet1") return false;
    const nr = sessionB.namedRanges.get("MyRange");
    const title = sessionB.metadata.get("title");
    return nr?.sheetId === "Sheet2" && nr?.range === "A1:B2" && title === "Quarterly Budget";
  }, 10_000);

  assert.deepEqual(snapshotSheets(sessionB).map((s) => s.id), ["Sheet2", "Sheet1"]);
  assert.deepEqual(sessionB.namedRanges.get("MyRange"), { sheetId: "Sheet2", range: "A1:B2" });
  assert.equal(sessionB.metadata.get("title"), "Quarterly Budget");

  // Tear down clients and restart the server, keeping the same data directory.
  sessionA.destroy();
  sessionB.destroy();
  sessionA.doc.destroy();
  sessionB.doc.destroy();
  sessionA = null;
  sessionB = null;

  // Give the server a moment to persist state after the last client disconnects.
  await new Promise((r) => setTimeout(r, 500));

  await server.stop();
  server = await startSyncServer({
    port,
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
  });

  sessionC = createCollabSession({
    connection: {
      wsUrl,
      docId,
      token: "test-token",
      WebSocketPolyfill: WebSocket,
      disableBc: true,
    },
  });

  await sessionC.whenSynced();

  await waitForCondition(() => {
    const sheets = snapshotSheets(sessionC);
    if (sheets.length !== 2) return false;
    if (sheets[0]?.id !== "Sheet2" || sheets[0]?.name !== "Budget") return false;
    if (sheets[1]?.id !== "Sheet1") return false;
    const nr = sessionC.namedRanges.get("MyRange");
    const title = sessionC.metadata.get("title");
    return nr?.sheetId === "Sheet2" && nr?.range === "A1:B2" && title === "Quarterly Budget";
  }, 10_000);

  assert.deepEqual(snapshotSheets(sessionC).map((s) => s.id), ["Sheet2", "Sheet1"]);
  assert.deepEqual(sessionC.namedRanges.get("MyRange"), { sheetId: "Sheet2", range: "A1:B2" });
  assert.equal(sessionC.metadata.get("title"), "Quarterly Budget");
});
