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

test("CollabSession integration: sync + presence (sync-server)", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "collab-session-sync-server-"));

  /** @type {import("../src/index.ts").CollabSession | null} */
  let sessionA = null;
  /** @type {import("../src/index.ts").CollabSession | null} */
  let sessionB = null;
  /** @type {Awaited<ReturnType<typeof startSyncServer>> | null} */
  let server = null;

  t.after(async () => {
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

  const docId = "collab-session-test-doc";
  const wsUrl = server.wsUrl;

  sessionA = createCollabSession({
    connection: {
      wsUrl,
      docId,
      token: "test-token",
      WebSocketPolyfill: WebSocket,
      disableBc: true,
    },
    presence: {
      user: { id: "user-a", name: "User A", color: "#ff0000" },
      activeSheet: "Sheet1",
      throttleMs: 0,
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
    presence: {
      user: { id: "user-b", name: "User B", color: "#00ff00" },
      activeSheet: "Sheet1",
      throttleMs: 0,
    },
  });

  await Promise.all([sessionA.whenSynced(), sessionB.whenSynced()]);

  await sessionA.setCellValue("Sheet1:0:0", 123);
  await waitForCondition(async () => (await sessionB.getCell("Sheet1:0:0"))?.value === 123, 10_000);
  assert.equal((await sessionB.getCell("Sheet1:0:0"))?.value, 123);

  await sessionA.setCellFormula("Sheet1:0:1", "=1+1");
  await waitForCondition(async () => (await sessionB.getCell("Sheet1:0:1"))?.formula === "=1+1", 10_000);
  assert.equal((await sessionB.getCell("Sheet1:0:1"))?.formula, "=1+1");

  sessionA.presence?.setCursor({ row: 5, col: 7 });
  await waitForCondition(() => {
    const remote = sessionB.presence?.getRemotePresences() ?? [];
    return remote.some((presence) => presence.cursor?.row === 5 && presence.cursor?.col === 7);
  }, 10_000);

  {
    const remote = sessionB.presence?.getRemotePresences() ?? [];
    assert.equal(remote.length, 1);
    assert.deepEqual(remote[0].cursor, { row: 5, col: 7 });
  }
});
