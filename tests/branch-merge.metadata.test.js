import test from "node:test";
import assert from "node:assert/strict";

import { applyConflictResolutions, mergeDocumentStates } from "../packages/versioning/branches/src/merge.js";

/**
 * @param {Partial<import("../packages/versioning/branches/src/types.js").DocumentState>} overrides
 */
function state(overrides) {
  return {
    schemaVersion: 1,
    sheets: {
      order: ["Sheet1"],
      metaById: { Sheet1: { id: "Sheet1", name: "Sheet1" } },
    },
    cells: { Sheet1: {} },
    metadata: {},
    namedRanges: {},
    comments: {},
    ...overrides,
  };
}

test("merge: metadata key collision surfaces conflict + resolves", () => {
  const base = state({ metadata: { answer: 1 } });
  const ours = state({ metadata: { answer: 2 } });
  const theirs = state({ metadata: { answer: 3 } });

  const result = mergeDocumentStates({ base, ours, theirs });
  const idx = result.conflicts.findIndex((c) => c.type === "metadata");
  assert.ok(idx >= 0, "expected metadata conflict");
  const conflict = result.conflicts[idx];
  assert.equal(conflict.type, "metadata");
  assert.equal(conflict.key, "answer");

  // Merge defaults to ours.
  assert.equal(result.merged.schemaVersion, 1);
  assert.equal(result.merged.metadata.answer, 2);

  const resolved = applyConflictResolutions(result, [{ conflictIndex: idx, choice: "theirs" }]);
  assert.equal(resolved.schemaVersion, 1);
  assert.equal(resolved.metadata.answer, 3);
});

test("merge: named range key collision surfaces conflict", () => {
  const base = state({ namedRanges: { MyRange: { ref: "A1" } } });
  const ours = state({ namedRanges: { MyRange: { ref: "B1" } } });
  const theirs = state({ namedRanges: { MyRange: { ref: "C1" } } });

  const result = mergeDocumentStates({ base, ours, theirs });
  const conflict = result.conflicts.find((c) => c.type === "namedRange");
  assert.ok(conflict);
  assert.equal(conflict.type, "namedRange");
  assert.equal(conflict.key, "MyRange");
});

test("merge: comment id collision surfaces conflict", () => {
  const base = state({
    comments: {
      c1: { id: "c1", content: "base", replies: [] },
    },
  });
  const ours = state({
    comments: {
      c1: { id: "c1", content: "ours", replies: [] },
    },
  });
  const theirs = state({
    comments: {
      c1: { id: "c1", content: "theirs", replies: [] },
    },
  });

  const result = mergeDocumentStates({ base, ours, theirs });
  const conflict = result.conflicts.find((c) => c.type === "comment");
  assert.ok(conflict);
  assert.equal(conflict.type, "comment");
  assert.equal(conflict.id, "c1");
});

test("merge: sheet rename + order collisions surface conflicts", () => {
  const base = {
    schemaVersion: 1,
    sheets: {
      order: ["Sheet1", "Sheet2", "Sheet3"],
      metaById: {
        Sheet1: { id: "Sheet1", name: "Sheet1" },
        Sheet2: { id: "Sheet2", name: "Sheet2" },
        Sheet3: { id: "Sheet3", name: "Sheet3" },
      },
    },
    cells: { Sheet1: {}, Sheet2: {}, Sheet3: {} },
    metadata: {},
    namedRanges: {},
    comments: {},
  };
  const ours = {
    ...structuredClone(base),
    sheets: {
      order: ["Sheet2", "Sheet1", "Sheet3"],
      metaById: {
        ...base.sheets.metaById,
        Sheet1: { id: "Sheet1", name: "OursName" },
      },
    },
  };
  const theirs = {
    ...structuredClone(base),
    sheets: {
      // Move Sheet1 to a different destination than ours to force an order conflict.
      order: ["Sheet2", "Sheet3", "Sheet1"],
      metaById: {
        ...base.sheets.metaById,
        Sheet1: { id: "Sheet1", name: "TheirsName" },
      },
    },
  };

  const result = mergeDocumentStates({ base, ours, theirs });
  assert.ok(result.conflicts.some((c) => c.type === "sheet" && c.reason === "order"));
  assert.ok(
    result.conflicts.some((c) => c.type === "sheet" && c.reason === "rename" && c.sheetId === "Sheet1"),
  );
});
