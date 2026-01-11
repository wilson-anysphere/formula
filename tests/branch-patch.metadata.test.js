import test from "node:test";
import assert from "node:assert/strict";

import { applyPatch, diffDocumentStates } from "../packages/versioning/branches/src/patch.js";
import { normalizeDocumentState } from "../packages/versioning/branches/src/state.js";

test("patch: diff/apply includes workbook metadata changes", () => {
  const base = normalizeDocumentState({
    schemaVersion: 1,
    sheets: { order: ["Sheet1"], metaById: { Sheet1: { id: "Sheet1", name: "Sheet1" } } },
    cells: { Sheet1: {} },
    metadata: { a: 1, b: 2, d: 4 },
    namedRanges: {},
    comments: {},
  });

  const next = structuredClone(base);
  next.metadata.b = 3; // modify
  next.metadata.c = { nested: true }; // add
  delete next.metadata.d; // delete

  const patch = diffDocumentStates(base, next);
  assert.equal(patch.schemaVersion, 1);
  assert.deepEqual(patch.metadata, { b: 3, c: { nested: true }, d: null });

  const applied = applyPatch(base, patch);
  assert.deepEqual(applied.metadata, next.metadata);
});

