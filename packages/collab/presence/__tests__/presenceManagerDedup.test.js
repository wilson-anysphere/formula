import assert from "node:assert/strict";
import test from "node:test";

import { PresenceManager } from "../index.js";

test("PresenceManager avoids broadcasting when cursor and selections are unchanged", () => {
  const calls = [];
  const awareness = {
    clientID: 1,
    setLocalStateField(field, value) {
      calls.push({ field, value });
    },
    getStates() {
      return new Map();
    },
  };

  const presence = new PresenceManager(awareness, {
    user: { id: "u1", name: "Ada", color: "#ff2d55" },
    activeSheet: "Sheet1",
    throttleMs: 0,
  });

  calls.length = 0;

  presence.setCursor({ row: 1, col: 1 });
  presence.setCursor({ row: 1, col: 1 });
  presence.setCursor({ row: 1.9, col: 1.1 });

  assert.equal(calls.length, 1);

  presence.setSelections([{ start: { row: 0, col: 0 }, end: { row: 2, col: 2 } }]);
  presence.setSelections([{ startRow: 0, startCol: 0, endRow: 2, endCol: 2 }]);
  presence.setSelections([{ startRow: 2, startCol: 2, endRow: 0, endCol: 0 }]);

  assert.equal(calls.length, 2);
});

