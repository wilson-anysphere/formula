import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";

import { startSyncServer } from "./test-helpers.ts";

test("startSyncServer fails fast when the sync-server process exits during startup", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-"));
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const startedAt = Date.now();
  await assert.rejects(
    startSyncServer({
      dataDir,
      auth: { mode: "opaque", token: "dev-token" },
      env: {
        // Force an early, deterministic startup failure (missing KeyRing config).
        NODE_ENV: "production",
        SYNC_SERVER_PERSISTENCE_ENCRYPTION: "keyring",
      },
    }),
    /Server failed to start|sync-server exited/i
  );

  assert.ok(Date.now() - startedAt < 5_000);
});

