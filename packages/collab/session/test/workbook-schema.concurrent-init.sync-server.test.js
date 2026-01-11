import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { randomUUID } from "node:crypto";

import WebSocket from "ws";

import { createCollabSession } from "../src/index.ts";
import {
  getAvailablePort,
  startSyncServer,
  waitForCondition,
} from "../../../../services/sync-server/test/test-helpers.ts";

/**
 * @param {import("../src/index.ts").CollabSession} session
 */
function sheetIds(session) {
  return session.sheets.toArray().map((sheet) => String(sheet.get("id") ?? ""));
}

test("CollabSession schema: concurrent init converges to a single default sheet and persists (sync-server)", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "collab-session-concurrent-schema-"));

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

  const docId = `collab-session-schema-${randomUUID()}`;
  const wsUrl = server.wsUrl;

  // Create both clients before awaiting sync to increase the chance of concurrent
  // schema initialization on a brand new document.
  sessionA = createCollabSession({
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

  await Promise.all([sessionA.whenSynced(), sessionB.whenSynced()]);

  await waitForCondition(() => {
    const idsA = sheetIds(sessionA);
    const idsB = sheetIds(sessionB);
    return (
      idsA.length === 1 &&
      idsB.length === 1 &&
      idsA[0] === "Sheet1" &&
      idsB[0] === "Sheet1"
    );
  }, 10_000);

  assert.deepEqual(sheetIds(sessionA), ["Sheet1"]);
  assert.deepEqual(sheetIds(sessionB), ["Sheet1"]);

  // Restart server and ensure the converged state is what persists.
  sessionA.destroy();
  sessionB.destroy();
  sessionA.doc.destroy();
  sessionB.doc.destroy();
  sessionA = null;
  sessionB = null;

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

  await waitForCondition(() => sheetIds(sessionC).length === 1, 10_000);
  assert.deepEqual(sheetIds(sessionC), ["Sheet1"]);
});

