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

test("patch: legacy cell patches can still apply when combined with metadata section", () => {
  const base = normalizeDocumentState({
    schemaVersion: 1,
    sheets: { order: ["Sheet1"], metaById: { Sheet1: { id: "Sheet1", name: "Sheet1" } } },
    cells: { Sheet1: { A1: { value: 1 } } },
    metadata: { scenario: "base" },
    namedRanges: {},
    comments: {},
  });

  // Legacy patch shape: cells stored under `patch.sheets`.
  const legacyPatch = {
    sheets: { Sheet1: { A1: { value: 2 } } },
    metadata: { scenario: "next" },
  };

  // @ts-expect-error - intentionally mixing legacy+new patch shape for compat.
  const applied = applyPatch(base, legacyPatch);
  assert.equal(applied.cells.Sheet1.A1.value, 2);
  assert.equal(applied.metadata.scenario, "next");
});

test("patch: diff/apply includes sheet view changes (frozen panes)", () => {
  const base = normalizeDocumentState({
    schemaVersion: 1,
    sheets: {
      order: ["Sheet1"],
      metaById: { Sheet1: { id: "Sheet1", name: "Sheet1", view: { frozenRows: 0, frozenCols: 0 } } },
    },
    cells: { Sheet1: {} },
    metadata: {},
    namedRanges: {},
    comments: {},
  });

  const next = structuredClone(base);
  next.sheets.metaById.Sheet1.view = { frozenRows: 2, frozenCols: 1 };

  const patch = diffDocumentStates(base, next);
  assert.equal(patch.schemaVersion, 1);
  assert.deepEqual(patch.sheets?.metaById?.Sheet1?.view, { frozenRows: 2, frozenCols: 1 });

  const applied = applyPatch(base, patch);
  assert.deepEqual(applied.sheets.metaById.Sheet1.view, { frozenRows: 2, frozenCols: 1 });
});
