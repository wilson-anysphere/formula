import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import WebSocket from "ws";

import { createCollabSession } from "../src/index.ts";
import { createCollabVersioning } from "../../versioning/src/index.ts";
import { SQLiteVersionStore } from "../../../versioning/src/store/sqliteVersionStore.js";
import {
  getAvailablePort,
  startSyncServer,
  waitForCondition,
} from "../../../../services/sync-server/test/test-helpers.ts";

test("CollabVersioning integration: restore syncs + persists (sync-server)", async (t) => {
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

  const docId = "collab-versioning-test-doc";
  const wsUrl = server.wsUrl;

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

  store = new SQLiteVersionStore({ filePath: versionDbPath });
  versioning = createCollabVersioning({
    session: sessionA,
    store,
    user: { userId: "user-a", userName: "User A" },
    autoStart: false,
  });

  // Initial edits + checkpoint.
  sessionA.setCellValue("Sheet1:0:0", "alpha");
  sessionA.setCellValue("Sheet1:0:1", 123);
  await waitForCondition(() => sessionB.getCell("Sheet1:0:0")?.value === "alpha", 10_000);
  await waitForCondition(() => sessionB.getCell("Sheet1:0:1")?.value === 123, 10_000);

  const checkpoint = await versioning.createCheckpoint({ name: "checkpoint-1" });

  // More edits (including a new cell that should be deleted on restore).
  sessionA.setCellValue("Sheet1:0:0", "beta");
  sessionA.setCellValue("Sheet1:2:0", "extra");
  await waitForCondition(() => sessionB.getCell("Sheet1:0:0")?.value === "beta", 10_000);
  await waitForCondition(() => sessionB.getCell("Sheet1:2:0")?.value === "extra", 10_000);

  // Restore the checkpoint and ensure the other collaborator converges.
  await versioning.restoreVersion(checkpoint.id);

  await waitForCondition(() => sessionB.getCell("Sheet1:0:0")?.value === "alpha", 10_000);
  await waitForCondition(() => sessionB.getCell("Sheet1:0:1")?.value === 123, 10_000);
  await waitForCondition(() => sessionB.getCell("Sheet1:2:0") == null, 10_000);

  assert.equal(sessionB.getCell("Sheet1:0:0")?.value, "alpha");
  assert.equal(sessionB.getCell("Sheet1:0:1")?.value, 123);
  assert.equal(sessionB.getCell("Sheet1:2:0"), null);

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

