import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import WebSocket from "ws";

import { createCollabSession } from "../src/index.ts";
import {
  getAvailablePort,
  startSyncServer,
  waitForCondition,
} from "../../../../services/sync-server/test/test-helpers.ts";

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

test("CollabSession offline persistence survives restart and merges on reconnect (sync-server)", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "collab-session-offline-sync-server-data-"));
  const offlineDir = await mkdtemp(path.join(tmpdir(), "collab-session-offline-sync-server-client-"));
  const offlineFilePath = path.join(offlineDir, "doc.yjslog");

  const docId = `collab-session-offline-test-doc-${crypto.randomUUID()}`;
  const port = await getAvailablePort();

  /** @type {Awaited<ReturnType<typeof startSyncServer>> | null} */
  let server = null;
  /** @type {import("../src/index.ts").CollabSession | null} */
  let sessionA = null;
  /** @type {import("../src/index.ts").CollabSession | null} */
  let sessionARestarted = null;
  /** @type {import("../src/index.ts").CollabSession | null} */
  let sessionB = null;

  t.after(async () => {
    sessionA?.destroy();
    sessionARestarted?.destroy();
    sessionB?.destroy();
    sessionA?.doc.destroy();
    sessionARestarted?.doc.destroy();
    sessionB?.doc.destroy();
    await server?.stop();
    await rm(dataDir, { recursive: true, force: true });
    await rm(offlineDir, { recursive: true, force: true });
  });

  server = await startSyncServer({
    port,
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: { SYNC_SERVER_PERSISTENCE_BACKEND: "file" },
  });

  const wsUrl = server.wsUrl;

  // Session A starts online with offline persistence enabled.
  sessionA = createCollabSession({
    connection: {
      wsUrl,
      docId,
      token: "test-token",
      WebSocketPolyfill: WebSocket,
      disableBc: true,
    },
    offline: { mode: "file", filePath: offlineFilePath },
  });

  await sessionA.offline?.whenLoaded();
  await sessionA.whenSynced();

  await sessionA.setCellValue("Sheet1:0:0", "online");
  await waitForCondition(async () => (await sessionA.getCell("Sheet1:0:0"))?.value === "online", 10_000);

  // Go offline.
  sessionA.disconnect();
  // Tear down the websocket provider before stopping the server so the child
  // process can exit promptly (avoids flaky shutdown delays in CI).
  sessionA.provider?.destroy?.();
  await server.stop();
  server = null;

  // Make edits while offline and "restart" the process.
  await sessionA.setCellValue("Sheet1:0:1", "offline");
  await sessionA.setCellFormula("Sheet1:0:2", "=1+2");

  sessionA.destroy();
  sessionA.doc.destroy();
  sessionA = null;

  // Restart the client while the server is still down. Ensure we load from the offline log.
  sessionARestarted = createCollabSession({
    connection: {
      wsUrl,
      docId,
      token: "test-token",
      WebSocketPolyfill: WebSocket,
      disableBc: true,
    },
    offline: { mode: "file", filePath: offlineFilePath },
  });

  await sessionARestarted.offline?.whenLoaded();

  assert.equal((await sessionARestarted.getCell("Sheet1:0:0"))?.value, "online");
  assert.equal((await sessionARestarted.getCell("Sheet1:0:1"))?.value, "offline");
  assert.equal((await sessionARestarted.getCell("Sheet1:0:2"))?.formula, "=1+2");

  // Bring the server back and ensure the restarted client merges its offline edits.
  await sleep(25);
  server = await startSyncServer({
    port,
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: { SYNC_SERVER_PERSISTENCE_BACKEND: "file" },
  });

  sessionARestarted.connect();
  await sessionARestarted.whenSynced();

  sessionB = createCollabSession({
    connection: {
      wsUrl: server.wsUrl,
      docId,
      token: "test-token",
      WebSocketPolyfill: WebSocket,
      disableBc: true,
    },
  });

  await sessionB.whenSynced();

  await waitForCondition(async () => (await sessionB.getCell("Sheet1:0:1"))?.value === "offline", 10_000);
  assert.equal((await sessionB.getCell("Sheet1:0:0"))?.value, "online");
  assert.equal((await sessionB.getCell("Sheet1:0:1"))?.value, "offline");
  assert.equal((await sessionB.getCell("Sheet1:0:2"))?.formula, "=1+2");
});
