import assert from "node:assert/strict";
import test from "node:test";

import { mergeDocumentStates } from "./merge.js";

test("mergeDocumentStates: merges sheet visibility/tabColor and treats omissions as no-op", () => {
  const base = {
    schemaVersion: 1,
    sheets: {
      order: ["Sheet1"],
      metaById: { Sheet1: { id: "Sheet1", name: "Sheet1", visibility: "hidden", tabColor: "FF00FF00" } },
    },
    cells: { Sheet1: {} },
    metadata: {},
    namedRanges: {},
    comments: {},
  };

  // Simulate an older client that omits newer sheet metadata fields entirely.
  const ours = structuredClone(base);
  delete ours.sheets.metaById.Sheet1.visibility;
  delete ours.sheets.metaById.Sheet1.tabColor;

  // Theirs changes the fields.
  const theirs = structuredClone(base);
  theirs.sheets.metaById.Sheet1.visibility = "visible";
  theirs.sheets.metaById.Sheet1.tabColor = "FFFF0000";

  const result = mergeDocumentStates({ base, ours, theirs });
  assert.equal(result.conflicts.length, 0);
  assert.equal(result.merged.sheets.metaById.Sheet1.visibility, "visible");
  assert.equal(result.merged.sheets.metaById.Sheet1.tabColor, "FFFF0000");
});

