import assert from "node:assert/strict";
import test from "node:test";

import { InMemoryAwarenessHub, PresenceManager } from "../index.js";

test("PresenceManager filters remote presences by active sheet", () => {
  const hub = new InMemoryAwarenessHub();
  const awarenessA = hub.createAwareness(1);
  const awarenessB = hub.createAwareness(2);

  const presenceA = new PresenceManager(awarenessA, {
    user: { id: "u1", name: "Ada", color: "#ff2d55" },
    activeSheet: "Sheet1",
  });

  const presenceB = new PresenceManager(awarenessB, {
    user: { id: "u2", name: "Grace", color: "#4c8bf5" },
    activeSheet: "Sheet1",
  });

  assert.equal(presenceA.getRemotePresences().length, 1);
  assert.equal(presenceA.getRemotePresences()[0].id, "u2");

  presenceB.setActiveSheet("Sheet2");

  assert.deepEqual(presenceA.getRemotePresences(), []);
  assert.equal(presenceA.getRemotePresences({ activeSheet: "Sheet2" }).length, 1);
});

test("PresenceManager.destroy removes local awareness state", () => {
  const hub = new InMemoryAwarenessHub();
  const awarenessA = hub.createAwareness(1);
  const awarenessB = hub.createAwareness(2);

  const presenceA = new PresenceManager(awarenessA, {
    user: { id: "u1", name: "Ada", color: "#ff2d55" },
    activeSheet: "Sheet1",
  });

  const presenceB = new PresenceManager(awarenessB, {
    user: { id: "u2", name: "Grace", color: "#4c8bf5" },
    activeSheet: "Sheet1",
  });

  assert.equal(presenceA.getRemotePresences().length, 1);

  presenceB.destroy();

  assert.deepEqual(presenceA.getRemotePresences(), []);
});

