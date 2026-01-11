import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { randomUUID } from "node:crypto";

import WebSocket from "ws";

import { FileCollabPersistence } from "@formula/collab-persistence/file";
import { createCollabSession } from "../src/index.ts";
import {
  getAvailablePort,
  startSyncServer,
  waitForCondition,
} from "../../../../services/sync-server/test/test-helpers.ts";

test("CollabSession integration: offline restart persists and merges on reconnect", async (t) => {
  const serverDataDir = await mkdtemp(path.join(tmpdir(), "collab-offline-merge-server-"));
  const clientDataDir = await mkdtemp(path.join(tmpdir(), "collab-offline-merge-client-"));

  /** @type {Awaited<ReturnType<typeof startSyncServer>> | null} */
  let server = null;
  /** @type {import("../src/index.ts").CollabSession | null} */
  let sessionA = null;
  /** @type {import("../src/index.ts").CollabSession | null} */
  let sessionA2 = null;
  /** @type {import("../src/index.ts").CollabSession | null} */
  let sessionB = null;

  t.after(async () => {
    sessionA?.destroy();
    sessionA?.doc.destroy();
    sessionA2?.destroy();
    sessionA2?.doc.destroy();
    sessionB?.destroy();
    sessionB?.doc.destroy();
    await server?.stop();
    await rm(serverDataDir, { recursive: true, force: true });
    await rm(clientDataDir, { recursive: true, force: true });
  });

  const port = await getAvailablePort();
  server = await startSyncServer({
    port,
    dataDir: serverDataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: { SYNC_SERVER_PERSISTENCE_BACKEND: "file" },
  });

  const wsUrl = server.wsUrl;
  const docId = `offline-merge-${randomUUID()}`;

  const persistenceA1 = new FileCollabPersistence(clientDataDir, { compactAfterUpdates: 5 });

  sessionA = createCollabSession({
    docId,
    persistence: persistenceA1,
    connection: {
      wsUrl,
      docId,
      token: "test-token",
      WebSocketPolyfill: WebSocket,
      disableBc: true,
    },
  });

  sessionB = createCollabSession({
    connection: {
      wsUrl,
      docId,
      token: "test-token",
      WebSocketPolyfill: WebSocket,
      disableBc: true,
    },
  });

  await Promise.all([sessionA.whenLocalPersistenceLoaded(), sessionA.whenSynced(), sessionB.whenSynced()]);

  // B edits while both clients are online.
  sessionB.setCellValue("Sheet1:0:0", "online");
  await waitForCondition(
    async () => (await sessionA.getCell("Sheet1:0:0"))?.value === "online",
    10_000
  );

  // A goes offline.
  sessionA.disconnect();

  // B continues editing online while A is offline.
  sessionB.setCellValue("Sheet1:0:1", "b-online");
  await waitForCondition(
    async () => (await sessionB.getCell("Sheet1:0:1"))?.value === "b-online",
    1_000
  );

  // A edits offline, then closes the app.
  sessionA.setCellValue("Sheet1:0:2", "a-offline");
  await sessionA.flushLocalPersistence();
  sessionA.destroy();
  sessionA.doc.destroy();
  await persistenceA1.flush(docId);
  sessionA = null;

  // Recreate A from the same local persistence directory and reconnect.
  const persistenceA2 = new FileCollabPersistence(clientDataDir, { compactAfterUpdates: 5 });
  sessionA2 = createCollabSession({
    docId,
    persistence: persistenceA2,
    connection: {
      wsUrl,
      docId,
      token: "test-token",
      WebSocketPolyfill: WebSocket,
      disableBc: true,
    },
  });

  await sessionA2.whenLocalPersistenceLoaded();
  assert.equal((await sessionA2.getCell("Sheet1:0:2"))?.value, "a-offline");

  await sessionA2.whenSynced();

  // B should receive A's offline changes after A reconnects.
  await waitForCondition(
    async () => (await sessionB.getCell("Sheet1:0:2"))?.value === "a-offline",
    10_000
  );
  assert.equal((await sessionB.getCell("Sheet1:0:2"))?.value, "a-offline");

  // Both should converge on B's edits made while A was offline.
  await waitForCondition(
    async () => (await sessionA2.getCell("Sheet1:0:1"))?.value === "b-online",
    10_000
  );
  assert.equal((await sessionA2.getCell("Sheet1:0:1"))?.value, "b-online");

  await persistenceA2.flush(docId);
});
