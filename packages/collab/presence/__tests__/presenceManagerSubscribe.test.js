import assert from "node:assert/strict";
import test from "node:test";

import { InMemoryAwarenessHub, PresenceManager } from "../index.js";

test("PresenceManager.subscribe notifies on remote changes (not local cursor updates)", () => {
  const hub = new InMemoryAwarenessHub();
  const awarenessA = hub.createAwareness(1);
  const awarenessB = hub.createAwareness(2);

  const presenceA = new PresenceManager(awarenessA, {
    user: { id: "u1", name: "Ada", color: "#ff2d55" },
    activeSheet: "Sheet1",
    throttleMs: 0,
  });

  const presenceB = new PresenceManager(awarenessB, {
    user: { id: "u2", name: "Grace", color: "#4c8bf5" },
    activeSheet: "Sheet1",
    throttleMs: 0,
  });

  /** @type {any[][]} */
  const updates = [];

  const unsubscribe = presenceA.subscribe((presences) => {
    updates.push(presences);
  });

  assert.equal(updates.length, 1);
  assert.equal(updates[0].length, 1);
  assert.equal(updates[0][0].id, "u2");

  // Local cursor updates should not trigger new notifications.
  presenceA.setCursor({ row: 1, col: 1 });
  assert.equal(updates.length, 1);

  // Remote cursor update triggers a notification.
  presenceB.setCursor({ row: 2, col: 3 });
  assert.equal(updates.length, 2);
  assert.deepEqual(updates[1][0].cursor, { row: 2, col: 3 });

  // Local sheet changes should notify so the UI can re-filter remote users.
  presenceA.setActiveSheet("Sheet2");
  assert.equal(updates.length, 3);
  assert.deepEqual(updates[2], []);

  unsubscribe();

  presenceB.setCursor({ row: 3, col: 4 });
  assert.equal(updates.length, 3);
});

