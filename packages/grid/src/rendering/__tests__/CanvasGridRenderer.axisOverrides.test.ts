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
});

