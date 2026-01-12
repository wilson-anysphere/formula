import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { CanvasGridRenderer } from "../CanvasGridRenderer";
import type { CellProvider } from "../../model/CellProvider";

describe("CanvasGridRenderer.applyAxisSizeOverrides", () => {
  beforeEach(() => {
    // Avoid scheduling real animation frames in the unit test environment.
    vi.stubGlobal("requestAnimationFrame", () => 0);
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("applies 1000+ overrides with only one render invalidation", () => {
    const provider: CellProvider = { getCell: () => null };
    const renderer = new CanvasGridRenderer({ provider, rowCount: 2_000, colCount: 2_000 });

    const requestRenderSpy = vi.spyOn(renderer, "requestRender");

    const rows = new Map<number, number>();
    const cols = new Map<number, number>();
    for (let i = 0; i < 1_500; i += 1) rows.set(i, 30);
    for (let i = 0; i < 1_200; i += 1) cols.set(i, 110);

    renderer.applyAxisSizeOverrides({ rows, cols }, { resetUnspecified: true });

    expect(requestRenderSpy).toHaveBeenCalledTimes(1);
  });

  it("clears previous overrides when resetUnspecified is true", () => {
    const provider: CellProvider = { getCell: () => null };
    const renderer = new CanvasGridRenderer({ provider, rowCount: 10, colCount: 10, defaultRowHeight: 10, defaultColWidth: 10 });

    const firstRows = new Map<number, number>([
      [1, 25],
      [2, 30]
    ]);

    renderer.applyAxisSizeOverrides({ rows: firstRows }, { resetUnspecified: true });
    expect(renderer.getRowHeight(1)).toBe(25);
    expect(renderer.getRowHeight(2)).toBe(30);

    const nextRows = new Map<number, number>([[2, 35]]);
    renderer.applyAxisSizeOverrides({ rows: nextRows }, { resetUnspecified: true });

    // Row 1 was not specified in the new map, so it should be reset to the default.
    expect(renderer.getRowHeight(1)).toBe(10);
    expect(renderer.getRowHeight(2)).toBe(35);
  });

  it("preserves unspecified overrides when resetUnspecified is false", () => {
    const provider: CellProvider = { getCell: () => null };
    const renderer = new CanvasGridRenderer({ provider, rowCount: 10, colCount: 10, defaultRowHeight: 10, defaultColWidth: 10 });

    renderer.applyAxisSizeOverrides({ rows: new Map([[1, 25]]) }, { resetUnspecified: true });
    expect(renderer.getRowHeight(1)).toBe(25);

    renderer.applyAxisSizeOverrides({ rows: new Map([[2, 30]]) }, { resetUnspecified: false });
    expect(renderer.getRowHeight(1)).toBe(25);
    expect(renderer.getRowHeight(2)).toBe(30);
  });

  it("treats default-sized entries as a request to clear the override", () => {
    const provider: CellProvider = { getCell: () => null };
    const renderer = new CanvasGridRenderer({ provider, rowCount: 10, colCount: 10, defaultRowHeight: 10, defaultColWidth: 10 });

    renderer.applyAxisSizeOverrides({ rows: new Map([[1, 25]]) }, { resetUnspecified: true });
    expect(renderer.getRowHeight(1)).toBe(25);

    // Set row 1 back to the default (10). This should clear the override.
    renderer.applyAxisSizeOverrides({ rows: new Map([[1, 10]]) }, { resetUnspecified: false });
    expect(renderer.getRowHeight(1)).toBe(10);
  });

  it("does nothing when applying the same overrides again (no extra invalidation)", () => {
    const provider: CellProvider = { getCell: () => null };
    const renderer = new CanvasGridRenderer({ provider, rowCount: 100, colCount: 100 });

    const requestRenderSpy = vi.spyOn(renderer, "requestRender");

    const rows = new Map<number, number>([
      [1, 30],
      [5, 45]
    ]);
    const cols = new Map<number, number>([
      [2, 120],
      [4, 80]
    ]);

    renderer.applyAxisSizeOverrides({ rows, cols }, { resetUnspecified: true });
    renderer.applyAxisSizeOverrides({ rows, cols }, { resetUnspecified: true });

    // The second call should be a no-op and not schedule another frame.
    expect(requestRenderSpy).toHaveBeenCalledTimes(1);
  });

  it("keeps runtime axis overrides consistent when setting a size very close to the default", () => {
    const provider: CellProvider = { getCell: () => null };
    const renderer = new CanvasGridRenderer({ provider, rowCount: 10, colCount: 10, defaultRowHeight: 10, defaultColWidth: 10 });

    // Within the renderer's epsilon, this should be treated as "default sized" and clear the override.
    renderer.setRowHeight(2, 10 + 5e-7);
    expect(renderer.getRowHeight(2)).toBe(10);
    expect(renderer.scroll.rows.getSize(2)).toBe(10);

    renderer.setColWidth(3, 10 + 5e-7);
    expect(renderer.getColWidth(3)).toBe(10);
    expect(renderer.scroll.cols.getSize(3)).toBe(10);
  });
});
