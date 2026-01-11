import { describe, expect, it } from "vitest";

import { AuditingOverlayRenderer } from "../AuditingOverlayRenderer";
import { colToName, expandRange, nameToCol, parseCellAddress } from "../address";

function makeFakeCtx() {
  const calls: Array<[string, ...any[]]> = [];

  // A minimal CanvasRenderingContext2D shim; drawing helpers in overlay renderers
  // call a small set of APIs. We record calls so assertions stay deterministic.
  const ctx: any = {
    canvas: { width: 100, height: 100 },
    save: () => calls.push(["save"]),
    restore: () => calls.push(["restore"]),
    setTransform: (...args: any[]) => calls.push(["setTransform", ...args]),
    clearRect: (...args: any[]) => calls.push(["clearRect", ...args]),
    fillRect: (...args: any[]) => calls.push(["fillRect", ...args]),
    strokeRect: (...args: any[]) => calls.push(["strokeRect", ...args]),
    beginPath: () => calls.push(["beginPath"]),
    rect: (...args: any[]) => calls.push(["rect", ...args]),
    clip: () => calls.push(["clip"]),
  };

  Object.defineProperty(ctx, "fillStyle", {
    set(value) {
      calls.push(["fillStyle", value]);
    },
  });

  Object.defineProperty(ctx, "strokeStyle", {
    set(value) {
      calls.push(["strokeStyle", value]);
    },
  });

  Object.defineProperty(ctx, "globalAlpha", {
    set(value) {
      calls.push(["globalAlpha", value]);
    },
  });

  Object.defineProperty(ctx, "lineWidth", {
    set(value) {
      calls.push(["lineWidth", value]);
    },
  });

  return { ctx: ctx as CanvasRenderingContext2D, calls };
}

describe("auditing overlay address helpers", () => {
  it("parses and formats A1-style column names", () => {
    expect(colToName(0)).toBe("A");
    expect(colToName(25)).toBe("Z");
    expect(colToName(26)).toBe("AA");
    expect(nameToCol("A")).toBe(0);
    expect(nameToCol("aa")).toBe(26);
    expect(nameToCol("A1")).toBe(null);
  });

  it("parses sheet-qualified cell addresses", () => {
    expect(parseCellAddress("A1")).toEqual({ row: 0, col: 0 });
    expect(parseCellAddress("Sheet1!B2")).toEqual({ row: 1, col: 1 });
    expect(parseCellAddress("'My Sheet'!C3")).toEqual({ row: 2, col: 2 });
  });

  it("expands simple ranges", () => {
    expect(expandRange("A1:A3")).toEqual(["A1", "A2", "A3"]);
    expect(expandRange("Sheet1!A1:B2")).toEqual(["A1", "B1", "A2", "B2"]);
  });
});

describe("AuditingOverlayRenderer", () => {
  it("draws precedent + dependent highlights", () => {
    const renderer = new AuditingOverlayRenderer({
      precedentFill: "red",
      precedentStroke: "red",
      dependentFill: "green",
      dependentStroke: "green",
    });
    const { ctx, calls } = makeFakeCtx();

    renderer.render(
      ctx,
      { precedents: ["A1", "not-a-cell"], dependents: ["B2"] },
      {
        getCellRect: (row, col) => ({ x: col * 10, y: row * 10, width: 10, height: 10 }),
      },
    );

    const fills = calls.filter((c) => c[0] === "fillRect");
    const strokes = calls.filter((c) => c[0] === "strokeRect");
    expect(fills).toHaveLength(2);
    expect(strokes).toHaveLength(2);
  });

  it("clears with an identity transform so DPR scaling doesn't interfere", () => {
    const renderer = new AuditingOverlayRenderer({ precedentFill: "red" });
    const { ctx, calls } = makeFakeCtx();
    renderer.clear(ctx);

    expect(calls.map((c) => c[0])).toEqual(["save", "setTransform", "clearRect", "restore"]);
  });
});

