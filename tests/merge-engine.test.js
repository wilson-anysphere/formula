import test from "node:test";
import assert from "node:assert/strict";

import { mergeDocumentStates } from "../packages/versioning/branches/src/merge.js";

test("merge: edits to different cells merge automatically", () => {
  const base = { sheets: { Sheet1: { A1: { value: 1 } } } };
  const ours = { sheets: { Sheet1: { A1: { value: 2 } } } };
  const theirs = { sheets: { Sheet1: { A1: { value: 1 }, B1: { value: 3 } } } };

  const result = mergeDocumentStates({ base, ours, theirs });
  assert.equal(result.conflicts.length, 0);
  assert.deepEqual(result.merged.sheets.Sheet1, { A1: { value: 2 }, B1: { value: 3 } });
});

test("merge: same-cell identical value auto-merges", () => {
  const base = { sheets: { Sheet1: { A1: { value: 1 } } } };
  const ours = { sheets: { Sheet1: { A1: { value: 2 } } } };
  const theirs = { sheets: { Sheet1: { A1: { value: 2 } } } };

  const result = mergeDocumentStates({ base, ours, theirs });
  assert.equal(result.conflicts.length, 0);
  assert.deepEqual(result.merged.sheets.Sheet1.A1, { value: 2 });
});

test("merge: formula AST-equivalent changes auto-merge", () => {
  const base = { sheets: { Sheet1: { A1: { formula: "=SUM(A1,B1)" } } } };
  const ours = { sheets: { Sheet1: { A1: { formula: "=sum( A1 , B1 )" } } } };
  const theirs = { sheets: { Sheet1: { A1: { formula: "=SUM(A1, B1)" } } } };

  const result = mergeDocumentStates({ base, ours, theirs });
  assert.equal(result.conflicts.length, 0);
  assert.equal(result.merged.sheets.Sheet1.A1.formula, "=sum( A1 , B1 )");
});

test("merge: format-only change merges with value change", () => {
  const base = { sheets: { Sheet1: { A1: { value: 1, format: { bold: false } } } } };
  const ours = { sheets: { Sheet1: { A1: { value: 1, format: { bold: true } } } } };
  const theirs = { sheets: { Sheet1: { A1: { value: 2, format: { bold: false } } } } };

  const result = mergeDocumentStates({ base, ours, theirs });
  assert.equal(result.conflicts.length, 0);
  assert.deepEqual(result.merged.sheets.Sheet1.A1, { value: 2, format: { bold: true } });
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
  assert.deepEqual(result.merged.sheets.Sheet1, { B1: { value: 2 } });
});

test("merge: move-to-different-destinations surfaces a move conflict", () => {
  const base = { sheets: { Sheet1: { A1: { value: 1 } } } };
  const ours = { sheets: { Sheet1: { B1: { value: 1 } } } };
  const theirs = { sheets: { Sheet1: { C1: { value: 1 } } } };

  const result = mergeDocumentStates({ base, ours, theirs });
  assert.ok(result.conflicts.some((c) => c.type === "move"));
});

