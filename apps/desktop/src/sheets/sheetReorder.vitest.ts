import { describe, expect, it } from "vitest";

import type { SheetMeta } from "./workbookSheetStore";
import { computeWorkbookSheetMoveIndex } from "./sheetReorder";

function makeSheet(id: string, visibility: SheetMeta["visibility"]): SheetMeta {
  return {
    id,
    name: id,
    visibility,
  };
}

function applyMove(sheets: SheetMeta[], fromSheetId: string, toIndex: number): SheetMeta[] {
  const fromIndex = sheets.findIndex((s) => s.id === fromSheetId);
  expect(fromIndex).toBeGreaterThanOrEqual(0);

  const next = sheets.slice();
  const [moved] = next.splice(fromIndex, 1);
  expect(moved).toBeTruthy();
  next.splice(toIndex, 0, moved!);
  return next;
}

describe("computeWorkbookSheetMoveIndex", () => {
  it("maps visible reorders through hidden sheets (hidden in middle)", () => {
    const sheets = [makeSheet("A", "visible"), makeSheet("B", "hidden"), makeSheet("C", "visible")];

    const toIndex = computeWorkbookSheetMoveIndex({
      sheets,
      fromSheetId: "C",
      dropTarget: { kind: "before", targetSheetId: "A" },
    });

    expect(toIndex).toBe(0);
    const next = applyMove(sheets, "C", toIndex!);
    expect(next.map((s) => `${s.id}:${s.visibility}`)).toEqual(["C:visible", "A:visible", "B:hidden"]);
  });

  it("supports drop at end of visible list (hidden in middle)", () => {
    const sheets = [makeSheet("A", "visible"), makeSheet("B", "hidden"), makeSheet("C", "visible")];

    const toIndex = computeWorkbookSheetMoveIndex({
      sheets,
      fromSheetId: "A",
      dropTarget: { kind: "end" },
    });

    // C is the last visible sheet (index 2), so moving A to "end of visible tabs"
    // becomes an insertion at absolute index 2 in the full sheet order.
    expect(toIndex).toBe(2);
    const next = applyMove(sheets, "A", toIndex!);
    expect(next.map((s) => s.id)).toEqual(["B", "C", "A"]);
  });

  it("inserts at end of visible list but before trailing hidden sheets", () => {
    const sheets = [
      makeSheet("A", "visible"),
      makeSheet("B", "hidden"),
      makeSheet("C", "visible"),
      makeSheet("D", "veryHidden"),
    ];

    const toIndex = computeWorkbookSheetMoveIndex({
      sheets,
      fromSheetId: "A",
      dropTarget: { kind: "end" },
    });

    // C is the last visible sheet; D is trailing hidden. The moved sheet should
    // land at index 2 (before D).
    expect(toIndex).toBe(2);
    const next = applyMove(sheets, "A", toIndex!);
    expect(next.map((s) => `${s.id}:${s.visibility}`)).toEqual(["B:hidden", "C:visible", "A:visible", "D:veryHidden"]);
  });

  it("supports inserting after a visible tab, crossing hidden sheets", () => {
    const sheets = [makeSheet("A", "visible"), makeSheet("B", "hidden"), makeSheet("C", "visible")];

    const toIndex = computeWorkbookSheetMoveIndex({
      sheets,
      fromSheetId: "C",
      dropTarget: { kind: "after", targetSheetId: "A" },
    });

    expect(toIndex).toBe(1);
    const next = applyMove(sheets, "C", toIndex!);
    expect(next.map((s) => `${s.id}:${s.visibility}`)).toEqual(["A:visible", "C:visible", "B:hidden"]);
  });

  it("preserves leading hidden sheets when inserting before the first visible tab", () => {
    const sheets = [makeSheet("H", "hidden"), makeSheet("A", "visible"), makeSheet("B", "visible")];

    const toIndex = computeWorkbookSheetMoveIndex({
      sheets,
      fromSheetId: "B",
      dropTarget: { kind: "before", targetSheetId: "A" },
    });

    expect(toIndex).toBe(1);
    const next = applyMove(sheets, "B", toIndex!);
    expect(next.map((s) => `${s.id}:${s.visibility}`)).toEqual(["H:hidden", "B:visible", "A:visible"]);
  });

  it("handles multiple hidden blocks when inserting within the visible strip", () => {
    const sheets = [
      makeSheet("A", "visible"),
      makeSheet("H1", "hidden"),
      makeSheet("B", "visible"),
      makeSheet("H2", "veryHidden"),
      makeSheet("C", "visible"),
    ];

    const toIndex = computeWorkbookSheetMoveIndex({
      sheets,
      fromSheetId: "C",
      dropTarget: { kind: "before", targetSheetId: "B" },
    });

    expect(toIndex).toBe(2);
    const next = applyMove(sheets, "C", toIndex!);
    expect(next.map((s) => `${s.id}:${s.visibility}`)).toEqual([
      "A:visible",
      "H1:hidden",
      "C:visible",
      "B:visible",
      "H2:veryHidden",
    ]);
  });
});

