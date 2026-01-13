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

test("PresenceManager.subscribe can include remote users on other sheets", () => {
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
    activeSheet: "Sheet2",
    throttleMs: 0,
  });

  /** @type {string[][]} */
  const activeSheetUpdates = [];
  /** @type {any[][]} */
  const allSheetsUpdates = [];

  const unsubscribeActiveSheet = presenceA.subscribe((presences) => {
    activeSheetUpdates.push(presences.map((presence) => presence.id));
  });

  const unsubscribeAllSheets = presenceA.subscribe(
    (presences) => {
      allSheetsUpdates.push(presences);
    },
    { includeOtherSheets: true },
  );

  // Default subscription remains active-sheet-only.
  assert.deepEqual(activeSheetUpdates, [[]]);

  // includeOtherSheets subscriptions receive presences across sheets.
  assert.equal(allSheetsUpdates.length, 1);
  assert.equal(allSheetsUpdates[0].length, 1);
  assert.equal(allSheetsUpdates[0][0].id, "u2");
  assert.equal(allSheetsUpdates[0][0].activeSheet, "Sheet2");

  presenceB.setCursor({ row: 1, col: 2 });

  assert.equal(allSheetsUpdates.length, 2);
  assert.deepEqual(allSheetsUpdates[1][0].cursor, { row: 1, col: 2 });

  unsubscribeActiveSheet();
  unsubscribeAllSheets();
});
