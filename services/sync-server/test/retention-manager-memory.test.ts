import assert from "node:assert/strict";
import test from "node:test";

import { createLogger } from "../src/logger.js";
import { DocConnectionTracker, LeveldbRetentionManager } from "../src/retention.js";

test("LeveldbRetentionManager sweeps stale lastSeenWriteMs entries", async () => {
  const ldb = {
    getAllDocNames: async () => [],
    clearDocument: async () => {},
    getMeta: async () => undefined,
    setMeta: async () => {},
  };

  const manager = new LeveldbRetentionManager(
    ldb,
    new DocConnectionTracker(),
    createLogger("silent"),
    60_000,
    1_000
  );

  await manager.markSeen("stale", { nowMs: 0, force: true });
  await manager.markSeen("recent", { nowMs: 20_000, force: true });

  assert.equal((manager as any).lastSeenWriteMs.size, 2);

  // Trigger a sweep pass via the next markSeen call.
  await manager.markSeen("trigger", { nowMs: 31_000, force: true });

  const map: Map<string, number> = (manager as any).lastSeenWriteMs;
  assert.equal(map.has("stale"), false);
  assert.equal(map.has("recent"), true);
  assert.equal(map.has("trigger"), true);
});

