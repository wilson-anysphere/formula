import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import WebSocket from "ws";

import { createCollabSession } from "../src/index.ts";
import { getAvailablePort, startSyncServer, waitForCondition } from "../../../../services/sync-server/test/test-helpers.ts";

test("CollabSession E2E encryption persists through sync-server restart (unauthorized clients stay masked)", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "collab-session-encryption-sync-server-"));

  /** @type {Awaited<ReturnType<typeof startSyncServer>> | null} */
  let server = null;
  /** @type {import("../src/index.ts").CollabSession | null} */
  let sessionA = null;
  /** @type {import("../src/index.ts").CollabSession | null} */
  let sessionB = null;

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

  const docId = "collab-session-encryption-test-doc";
  const wsUrl = server.wsUrl;

  const keyBytes = new Uint8Array(32).fill(9);
  const keyForA1 = (cell) => {
    if (cell.sheetId === "Sheet1" && cell.row === 0 && cell.col === 0) {
      return { keyId: "k-range-1", keyBytes };
    }
    return null;
  };

  sessionA = createCollabSession({
    connection: {
      wsUrl,
      docId,
      token: "test-token",
      WebSocketPolyfill: WebSocket,
      disableBc: true,
    },
    encryption: { keyForCell: keyForA1 },
  });

  sessionB = createCollabSession({
    connection: {
      wsUrl,
      docId,
      token: "test-token",
      WebSocketPolyfill: WebSocket,
      disableBc: true,
    },
    // No key on B.
  });

  await Promise.all([sessionA.whenSynced(), sessionB.whenSynced()]);

  await sessionA.setCellValue("Sheet1:0:0", "server-persisted-secret");

  await waitForCondition(async () => {
    const cell = await sessionB.getCell("Sheet1:0:0");
    return cell?.encrypted === true && cell?.value === "###";
  }, 10_000);

  {
    const cell = await sessionB.getCell("Sheet1:0:0");
    assert.equal(cell?.value, "###");
    assert.equal(cell?.formula, null);
    assert.equal(cell?.encrypted, true);
  }

  // Ensure raw Yjs still doesn't contain plaintext.
  {
    const cellMap = sessionB.cells.get("Sheet1:0:0");
    assert.ok(cellMap, "expected Yjs cell map to exist");
    assert.equal(cellMap.get("value"), undefined);
    assert.equal(cellMap.get("formula"), undefined);
    assert.ok(cellMap.get("enc"), "expected encrypted payload under `enc`");
    assert.equal(JSON.stringify(cellMap.toJSON()).includes("server-persisted-secret"), false);
  }

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
    env: { SYNC_SERVER_PERSISTENCE_BACKEND: "file" },
  });

  const sessionC = createCollabSession({
    connection: {
      wsUrl,
      docId,
      token: "test-token",
      WebSocketPolyfill: WebSocket,
      disableBc: true,
    },
    // No key on C either.
  });
  t.after(() => {
    sessionC.destroy();
    sessionC.doc.destroy();
  });

  await sessionC.whenSynced();

  await waitForCondition(async () => {
    const cell = await sessionC.getCell("Sheet1:0:0");
    return cell?.encrypted === true && cell?.value === "###" && cell?.formula === null;
  }, 10_000);

  const cellC = await sessionC.getCell("Sheet1:0:0");
  assert.equal(cellC?.value, "###");
  assert.equal(cellC?.formula, null);
  assert.equal(cellC?.encrypted, true);
});
