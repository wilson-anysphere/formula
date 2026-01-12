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
    // Keep each rectangle below the range-run threshold so the guard counts every cell
    // (and blocks once we exceed the total enumerated cell cap).
    setFillColor(doc, "Sheet1", ["A1:Z1923", "A3000:Z4922", "A6000:A6004"], "#FFFF0000");
    const elapsed = performance.now() - start;

    expect(elapsed).toBeLessThan(250);
    expect(spy).not.toHaveBeenCalled();

    const toast = document.querySelector('[data-testid="toast"]') as HTMLElement | null;
    expect(toast).not.toBeNull();
    expect(toast?.dataset.type).toBe("warning");
    expect(toast?.textContent).toMatch(/Selection too large to apply formatting/i);
  });

  it("refuses to apply formatting to an enormous rectangular selection", () => {
    const doc = new DocumentController();
    const spy = vi.spyOn(doc, "setRangeFormat");

    const start = performance.now();
    const applied = setFillColor(doc, "Sheet1", "A1:Z1000000", "#FFFF0000");
    const elapsed = performance.now() - start;

    expect(elapsed).toBeLessThan(250);
    // Large rectangles use a compressed range-run formatting representation and should
    // be allowed (the safety cap is for per-cell enumeration across many small ranges).
    expect(applied).toBe(true);
    expect(spy).toHaveBeenCalled();
    expect(document.querySelector('[data-testid="toast"]')).toBeNull();
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
