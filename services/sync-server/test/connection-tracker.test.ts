import assert from "node:assert/strict";
import test from "node:test";

import { ConnectionTracker } from "../src/limits.js";

test("ConnectionTracker treats maxConnections/maxConnectionsPerIp <= 0 as unlimited", () => {
  const tracker = new ConnectionTracker(0, 0);

  assert.deepEqual(tracker.tryRegister("ip-a"), { ok: true });
  assert.deepEqual(tracker.tryRegister("ip-a"), { ok: true });
  assert.deepEqual(tracker.tryRegister("ip-b"), { ok: true });

  assert.deepEqual(tracker.snapshot(), { total: 3, uniqueIps: 2 });
});

test("ConnectionTracker enforces maxConnectionsPerIp when set", () => {
  const tracker = new ConnectionTracker(0, 1);
  assert.deepEqual(tracker.tryRegister("ip-a"), { ok: true });
  assert.deepEqual(tracker.tryRegister("ip-a"), { ok: false, reason: "max_connections_per_ip_exceeded" });
  assert.deepEqual(tracker.tryRegister("ip-b"), { ok: true });
});

