import assert from "node:assert/strict";
import test from "node:test";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { randomUUID } from "node:crypto";
import { createRequire } from "node:module";

import { createCollabSession } from "../packages/collab/session/src/index.ts";
import { createCollabVersioning } from "../packages/collab/versioning/src/index.ts";
import { YjsVersionStore } from "../packages/versioning/src/store/yjsVersionStore.js";
import {
  getAvailablePort,
  startSyncServer,
  waitForCondition,
} from "../services/sync-server/test/test-helpers.ts";

test("sync-server e2e: YjsVersionStore shares version history + restores + persists after restart", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-versioning-e2e-"));
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

  const docId = `yjs-version-history-${randomUUID()}`;
  const wsUrl = server.wsUrl;

  const createClient = ({ userId, userName, storeCompression, explicitStore = false }) => {
    const session = createCollabSession({
      connection: {
        wsUrl,
        docId,
        token: "test-token",
        WebSocketPolyfill: WebSocket,
        disableBc: true,
      },
      defaultSheetId: "Sheet1",
    });

    const store = explicitStore
      ? new YjsVersionStore({
          doc: session.doc,
          ...(storeCompression ? { compression: storeCompression } : {}),
          chunkSize: 8 * 1024,
        })
      : null;

    const versioning = createCollabVersioning(
      store
        ? {
            session,
            store,
            user: { userId, userName },
            autoStart: false,
          }
        : {
            session,
            user: { userId, userName },
            autoStart: false,
          }
    );

    let destroyed = false;
    const destroy = () => {
      if (destroyed) return;
      destroyed = true;
      versioning.destroy();
      session.destroy();
      session.doc.destroy();
    };

    return { session, store, versioning, destroy };
  };

  const clientA = createClient({
    userId: "u-a",
    userName: "User A",
    storeCompression: "gzip",
    explicitStore: true,
  });
  // B uses the default store (YjsVersionStore) provided by CollabVersioning.
  const clientB = createClient({ userId: "u-b", userName: "User B" });

  t.after(() => {
    clientA.destroy();
    clientB.destroy();
  });

  await Promise.all([clientA.session.whenSynced(), clientB.session.whenSynced()]);

  // Seed initial workbook state and wait for propagation.
  clientA.session.setCellValue("Sheet1:0:0", 1);
  await waitForCondition(() => clientB.session.getCell("Sheet1:0:0")?.value === 1, 10_000);

  const checkpoint = await clientA.versioning.createCheckpoint({ name: "Approved", locked: true });
  assert.equal(checkpoint.kind, "checkpoint");

  clientA.session.setCellValue("Sheet1:0:0", 2);
  await waitForCondition(() => clientB.session.getCell("Sheet1:0:0")?.value === 2, 10_000);
  const snapshot = await clientA.versioning.createSnapshot({ description: "edit" });
  assert.equal(snapshot.kind, "snapshot");

  // --- Version history sync A -> B ---
  await waitForCondition(async () => {
    const versions = await clientB.versioning.listVersions();
    return versions.some((v) => v.id === checkpoint.id) && versions.some((v) => v.id === snapshot.id);
  }, 10_000);

  {
    const versions = await clientB.versioning.listVersions();
    const seenCheckpoint = versions.find((v) => v.id === checkpoint.id);
    assert.ok(seenCheckpoint);
    assert.equal(seenCheckpoint.checkpointName, "Approved");
    assert.equal(seenCheckpoint.checkpointLocked, true);
  }

  // --- checkpointLocked updates sync A -> B ---
  await clientA.versioning.setCheckpointLocked(checkpoint.id, false);
  await waitForCondition(async () => {
    const versions = await clientB.versioning.listVersions();
    const v = versions.find((row) => row.id === checkpoint.id);
    return v?.checkpointLocked === false;
  }, 10_000);

  // --- Restore propagates A -> B ---
  await clientA.versioning.restoreVersion(checkpoint.id);
  await waitForCondition(() => clientB.session.getCell("Sheet1:0:0")?.value === 1, 10_000);

  await waitForCondition(async () => {
    const versions = await clientB.versioning.listVersions();
    return versions.some((v) => v.kind === "restore");
  }, 10_000);
  {
    const versions = await clientB.versioning.listVersions();
    assert.ok(versions.some((v) => v.id === checkpoint.id), "expected checkpoint to survive restore");
    assert.ok(versions.some((v) => v.id === snapshot.id), "expected snapshot to survive restore");
  }

  // Tear down clients and restart the server (persisted doc + version history should survive).
  clientA.destroy();
  clientB.destroy();

  // Give the server a moment to flush persisted state after the last client disconnects.
  await new Promise((r) => setTimeout(r, 500));

  await server.stop();
  server = await startSyncServer({
    port,
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
  });

  // C also uses the default store (YjsVersionStore) via CollabVersioning.
  const clientC = createClient({ userId: "u-c", userName: "User C" });
  t.after(() => clientC.destroy());

  await clientC.session.whenSynced();

  // --- Hydration from persisted sync-server state ---
  await waitForCondition(() => clientC.session.getCell("Sheet1:0:0")?.value === 1, 10_000);

  const versionsC = await clientC.versioning.listVersions();
  assert.ok(versionsC.some((v) => v.id === checkpoint.id), "expected checkpoint to persist in doc");
  assert.ok(versionsC.some((v) => v.id === snapshot.id), "expected snapshot to persist in doc");
  assert.ok(versionsC.some((v) => v.kind === "restore"), "expected restore head to persist in doc");
  {
    const v = versionsC.find((row) => row.id === checkpoint.id);
    assert.equal(v?.checkpointLocked, false, "expected checkpointLocked update to persist in doc");
  }
});
