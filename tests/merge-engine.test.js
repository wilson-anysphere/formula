import test from "node:test";
import assert from "node:assert/strict";

import { mergeDocumentStates } from "../packages/versioning/branches/src/merge.js";

test("merge: edits to different cells merge automatically", () => {
  const base = { sheets: { Sheet1: { A1: { value: 1 } } } };
  const ours = { sheets: { Sheet1: { A1: { value: 2 } } } };
  const theirs = { sheets: { Sheet1: { A1: { value: 1 }, B1: { value: 3 } } } };

  const result = mergeDocumentStates({ base, ours, theirs });
  assert.equal(result.conflicts.length, 0);
  assert.deepEqual(result.merged.cells.Sheet1, { A1: { value: 2 }, B1: { value: 3 } });
});

test("merge: same-cell identical value auto-merges", () => {
  const base = { sheets: { Sheet1: { A1: { value: 1 } } } };
  const ours = { sheets: { Sheet1: { A1: { value: 2 } } } };
  const theirs = { sheets: { Sheet1: { A1: { value: 2 } } } };

  const result = mergeDocumentStates({ base, ours, theirs });
  assert.equal(result.conflicts.length, 0);
  assert.deepEqual(result.merged.cells.Sheet1.A1, { value: 2 });
});

test("merge: formula AST-equivalent changes auto-merge", () => {
  const base = { sheets: { Sheet1: { A1: { formula: "=SUM(A1,B1)" } } } };
  const ours = { sheets: { Sheet1: { A1: { formula: "=sum( A1 , B1 )" } } } };
  const theirs = { sheets: { Sheet1: { A1: { formula: "=SUM(A1, B1)" } } } };

  const result = mergeDocumentStates({ base, ours, theirs });
  assert.equal(result.conflicts.length, 0);
  assert.equal(result.merged.cells.Sheet1.A1.formula, "=sum( A1 , B1 )");
});

test("merge: format-only change merges with value change", () => {
  const base = { sheets: { Sheet1: { A1: { value: 1, format: { bold: false } } } } };
  const ours = { sheets: { Sheet1: { A1: { value: 1, format: { bold: true } } } } };
  const theirs = { sheets: { Sheet1: { A1: { value: 2, format: { bold: false } } } } };

  const result = mergeDocumentStates({ base, ours, theirs });
  assert.equal(result.conflicts.length, 0);
  assert.deepEqual(result.merged.cells.Sheet1.A1, { value: 2, format: { bold: true } });
});

test("merge: conflicting same-cell value edits create conflict", () => {
  const base = { sheets: { Sheet1: { A1: { value: 1 } } } };
  const ours = { sheets: { Sheet1: { A1: { value: 2 } } } };
  const theirs = { sheets: { Sheet1: { A1: { value: 3 } } } };

  const result = mergeDocumentStates({ base, ours, theirs });
  assert.equal(result.conflicts.length, 1);
  assert.equal(result.conflicts[0].type, "cell");
  assert.equal(result.conflicts[0].cell, "A1");
});

test("merge: move + edit is merged by relocating edit onto moved cell", () => {
  const base = { sheets: { Sheet1: { A1: { value: 1 } } } };
  const ours = { sheets: { Sheet1: { B1: { value: 1 } } } }; // move A1 -> B1 (A1 deleted)
  const theirs = { sheets: { Sheet1: { A1: { value: 2 } } } }; // edit A1

  const result = mergeDocumentStates({ base, ours, theirs });
  assert.equal(result.conflicts.length, 0);
  assert.deepEqual(result.merged.cells.Sheet1, { B1: { value: 2 } });
});

test("merge: move-to-different-destinations surfaces a move conflict", () => {
  const base = { sheets: { Sheet1: { A1: { value: 1 } } } };
  const ours = { sheets: { Sheet1: { B1: { value: 1 } } } };
  const theirs = { sheets: { Sheet1: { C1: { value: 1 } } } };

  const result = mergeDocumentStates({ base, ours, theirs });
  assert.ok(result.conflicts.some((c) => c.type === "move"));
});

test("merge: sheet rename conflict is detected", () => {
  const base = {
    schemaVersion: 1,
    sheets: { order: ["s1"], metaById: { s1: { id: "s1", name: "Sheet1" } } },
    cells: { s1: {} },
    namedRanges: {},
    comments: {},
  };
  const ours = structuredClone(base);
  ours.sheets.metaById.s1.name = "Ours";
  const theirs = structuredClone(base);
  theirs.sheets.metaById.s1.name = "Theirs";

  const result = mergeDocumentStates({ base, ours, theirs });
  assert.ok(result.conflicts.some((c) => c.type === "sheet" && c.reason === "rename" && c.sheetId === "s1"));
  assert.equal(result.merged.sheets.metaById.s1.name, "Ours");
});

test("merge: named range conflict is detected", () => {
  const base = {
    schemaVersion: 1,
    sheets: { order: ["Sheet1"], metaById: { Sheet1: { id: "Sheet1", name: "Sheet1" } } },
    cells: { Sheet1: {} },
    namedRanges: { NR1: { sheetId: "Sheet1", rect: { r0: 0, c0: 0, r1: 0, c1: 0 } } },
    comments: {},
  };
  const ours = structuredClone(base);
  ours.namedRanges.NR1 = { sheetId: "Sheet1", rect: { r0: 1, c0: 1, r1: 1, c1: 1 } };
  const theirs = structuredClone(base);
  theirs.namedRanges.NR1 = { sheetId: "Sheet1", rect: { r0: 2, c0: 2, r1: 2, c1: 2 } };

  const result = mergeDocumentStates({ base, ours, theirs });
  assert.ok(result.conflicts.some((c) => c.type === "namedRange" && c.key === "NR1"));
  assert.deepEqual(result.merged.namedRanges.NR1, ours.namedRanges.NR1);
});

test("merge: non-overlapping sheet reorders merge automatically", () => {
  const base = {
    schemaVersion: 1,
    sheets: {
      order: ["A", "B", "C", "D"],
      metaById: {
        A: { id: "A", name: "A" },
        B: { id: "B", name: "B" },
        C: { id: "C", name: "C" },
        D: { id: "D", name: "D" },
      },
    },
    cells: { A: {}, B: {}, C: {}, D: {} },
    namedRanges: {},
    comments: {},
  };

  const ours = structuredClone(base);
  ours.sheets.order = ["B", "A", "C", "D"]; // move A after B

  const theirs = structuredClone(base);
  theirs.sheets.order = ["A", "B", "D", "C"]; // move D before C

  const result = mergeDocumentStates({ base, ours, theirs });
  assert.equal(result.conflicts.length, 0);
  assert.deepEqual(result.merged.sheets.order, ["B", "A", "D", "C"]);
});
