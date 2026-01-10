import assert from "node:assert/strict";
import test from "node:test";

import { deserializePresenceState, serializePresenceState } from "../index.js";

test("presence serialization roundtrip", () => {
  const state = {
    id: "user-1",
    name: "Ada",
    color: "#ff00aa",
    activeSheet: "Sheet1",
    cursor: { row: 2, col: 4 },
    selections: [
      { start: { row: 10, col: 1 }, end: { row: 12, col: 5 } },
      { startRow: 3, startCol: 2, endRow: 1, endCol: 7 },
    ],
    lastActive: 12345,
  };

  const serialized = serializePresenceState(state);
  const roundTripped = deserializePresenceState(JSON.parse(JSON.stringify(serialized)));

  assert.deepEqual(roundTripped, {
    ...state,
    selections: [
      { startRow: 10, startCol: 1, endRow: 12, endCol: 5 },
      { startRow: 1, startCol: 2, endRow: 3, endCol: 7 },
    ],
  });
});

test("presence deserialization rejects invalid payloads", () => {
  assert.equal(deserializePresenceState(null), null);
  assert.equal(deserializePresenceState({ v: 1 }), null);
  assert.equal(
    deserializePresenceState({
      v: 1,
      id: "u",
      name: "n",
      color: "#fff",
      sheet: "Sheet1",
      cursor: { row: "a", col: 1 },
      selections: [],
      lastActive: 1,
    }),
    null,
  );
});
