/**
 * @vitest-environment jsdom
 */

import { beforeEach, describe, expect, it, vi } from "vitest";

import { DocumentController } from "../../document/documentController.js";
import { setFillColor } from "../toolbar.js";

describe("toolbar formatting safety cap", () => {
  beforeEach(() => {
    document.body.innerHTML = "";
    const toastRoot = document.createElement("div");
    toastRoot.id = "toast-root";
    document.body.appendChild(toastRoot);
  });

  it("refuses to apply formatting when the (non-band) selection exceeds the cap", () => {
    const doc = new DocumentController();
    const spy = vi.spyOn(doc, "setRangeFormat");

    const start = performance.now();
    // Two medium rectangles + a small tail range pushes us just over the total cell cap (100k).
    setFillColor(doc, "Sheet1", ["A1:Z1923", "A3000:Z4922", "A6000:A6004"], "#FFFF0000");
    const elapsed = performance.now() - start;

    expect(elapsed).toBeLessThan(250);
    expect(spy).not.toHaveBeenCalled();

    const toast = document.querySelector('[data-testid="toast"]') as HTMLElement | null;
    expect(toast).not.toBeNull();
    expect(toast?.dataset.type).toBe("warning");
    expect(toast?.textContent).toMatch(/Selection too large to apply formatting/i);
  });

  it("allows formatting for large rectangles that use the range-run fast path", () => {
    const doc = new DocumentController();
    const spy = vi.spyOn(doc, "setRangeFormat");

    const start = performance.now();
    // 26 cols * 3846 rows = 99,996 cells (below the 100k cap, above the 50k range-run threshold).
    const applied = setFillColor(doc, "Sheet1", "A1:Z3846", "#FFFF0000");
    const elapsed = performance.now() - start;

    expect(elapsed).toBeLessThan(250);
    expect(applied).toBe(true);
    expect(spy).toHaveBeenCalledTimes(1);
    expect(document.querySelector('[data-testid="toast"]')).toBeNull();

    const style = doc.getCellFormat("Sheet1", { row: 0, col: 0 }) as any;
    expect(style?.fill?.fgColor).toBe("#FFFF0000");
  });

  it("refuses to apply formatting to an enormous rectangular selection", () => {
    const doc = new DocumentController();
    const spy = vi.spyOn(doc, "setRangeFormat");

    const start = performance.now();
    const applied = setFillColor(doc, "Sheet1", "A1:Z1000000", "#FFFF0000");
    const elapsed = performance.now() - start;

    expect(elapsed).toBeLessThan(250);
    expect(applied).toBe(false);
    expect(spy).not.toHaveBeenCalled();

    const toast = document.querySelector('[data-testid="toast"]') as HTMLElement | null;
    expect(toast).not.toBeNull();
    expect(toast?.dataset.type).toBe("warning");
    expect(toast?.textContent).toMatch(/Selection too large to apply formatting/i);

    const style = doc.getCellFormat("Sheet1", { row: 0, col: 0 }) as any;
    expect(style?.fill?.fgColor).not.toBe("#FFFF0000");
  });

  it("refuses to apply formatting to extremely large full-row selections (row band cap)", () => {
    const doc = new DocumentController();
    const spy = vi.spyOn(doc, "setRangeFormat");

    const start = performance.now();
    setFillColor(doc, "Sheet1", "A1:XFD60000", "#FFFF0000");
    const elapsed = performance.now() - start;

    expect(elapsed).toBeLessThan(250);
    expect(spy).not.toHaveBeenCalled();
    expect(document.querySelector('[data-testid="toast"]')).not.toBeNull();
  });

  it("still applies formatting for small ranges", () => {
    const doc = new DocumentController();

    setFillColor(doc, "Sheet1", "A1:B2", "#FFFF0000");

    const cell = doc.getCell("Sheet1", "A1") as any;
    const style = doc.styleTable.get(cell.styleId) as any;
    expect(style.fill?.fgColor).toBe("#FFFF0000");
    expect(document.querySelector('[data-testid="toast"]')).toBeNull();
  });
});
