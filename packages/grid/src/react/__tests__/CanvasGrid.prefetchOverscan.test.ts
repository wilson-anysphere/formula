// @vitest-environment jsdom

import React from "react";
import { createRoot } from "react-dom/client";
import { act } from "react-dom/test-utils";
import { describe, expect, it, vi } from "vitest";
import type { CellRange } from "../../model/CellProvider";
import { CanvasGrid } from "../CanvasGrid";
import type { GridApi } from "../CanvasGrid";

function restoreActEnvironment(previous: unknown): void {
  if (previous === undefined) {
    Reflect.deleteProperty(globalThis as any, "IS_REACT_ACT_ENVIRONMENT");
    return;
  }
  (globalThis as any).IS_REACT_ACT_ENVIRONMENT = previous;
}

describe("CanvasGrid prefetch overscan", () => {
  it("prefetches beyond the visible viewport by the configured overscan", async () => {
    const previousActEnvironment = (globalThis as any).IS_REACT_ACT_ENVIRONMENT;
    (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

    const prefetch = vi.fn<(range: CellRange) => void>();

    vi.stubGlobal(
      "ResizeObserver",
      class ResizeObserver {
        observe(_target: Element): void {}
        unobserve(_target: Element): void {}
        disconnect(): void {}
      }
    );

    // We don't need to render a frame for this test; avoid executing the full canvas renderer.
    vi.stubGlobal("requestAnimationFrame", vi.fn((_cb: FrameRequestCallback) => 0));

    const viewportWidth = 50;
    const viewportHeight = 40;

    const overscanRows = 2;
    const overscanCols = 3;

    const rowHeight = 10;
    const colWidth = 10;

    const boundingRect = vi
      .spyOn(HTMLElement.prototype, "getBoundingClientRect")
      .mockImplementation(
        () =>
          ({
            width: viewportWidth,
            height: viewportHeight,
            top: 0,
            left: 0,
            right: viewportWidth,
            bottom: viewportHeight,
            x: 0,
            y: 0,
            toJSON: () => ({})
          }) as DOMRect
      );

    const ctxStub: Partial<CanvasRenderingContext2D> = {
      setTransform: vi.fn(),
      measureText: (text: string) =>
        ({
          width: text.length * 6,
          actualBoundingBoxAscent: 8,
          actualBoundingBoxDescent: 2
        }) as TextMetrics
    };

    const getContext = vi
      .spyOn(HTMLCanvasElement.prototype, "getContext")
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      .mockImplementation(() => ctxStub as any);

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        React.createElement(CanvasGrid, {
          provider: {
            getCell: () => null,
            prefetch
          },
          rowCount: 100,
          colCount: 100,
          defaultRowHeight: rowHeight,
          defaultColWidth: colWidth,
          prefetchOverscanRows: overscanRows,
          prefetchOverscanCols: overscanCols
        })
      );
    });

    const lastCall = prefetch.mock.calls.at(-1)?.[0];
    expect(lastCall).toBeTruthy();

    // With a 50x40 viewport and 10x10 cell sizes we render:
    // - rows [0, 4)
    // - cols [0, 5)
    const visibleRange = {
      startRow: 0,
      endRow: 4,
      startCol: 0,
      endCol: 5
    };

    expect(lastCall).toEqual({
      startRow: Math.max(0, visibleRange.startRow - overscanRows),
      endRow: Math.min(100, visibleRange.endRow + overscanRows),
      startCol: Math.max(0, visibleRange.startCol - overscanCols),
      endCol: Math.min(100, visibleRange.endCol + overscanCols)
    });
    expect(prefetch).toHaveBeenCalledTimes(1);

    await act(async () => {
      root.unmount();
    });
    host.remove();

    boundingRect.mockRestore();
    getContext.mockRestore();
    vi.unstubAllGlobals();
    restoreActEnvironment(previousActEnvironment);
  });

  it("clamps overscanned prefetch range to grid bounds", async () => {
    const previousActEnvironment = (globalThis as any).IS_REACT_ACT_ENVIRONMENT;
    (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

    const prefetch = vi.fn<(range: CellRange) => void>();

    vi.stubGlobal(
      "ResizeObserver",
      class ResizeObserver {
        observe(_target: Element): void {}
        unobserve(_target: Element): void {}
        disconnect(): void {}
      }
    );

    vi.stubGlobal("requestAnimationFrame", vi.fn((_cb: FrameRequestCallback) => 0));

    const viewportWidth = 50;
    const viewportHeight = 40;

    const overscanRows = 2;
    const overscanCols = 3;

    const rowHeight = 10;
    const colWidth = 10;

    const rowCount = 10;
    const colCount = 10;

    const boundingRect = vi
      .spyOn(HTMLElement.prototype, "getBoundingClientRect")
      .mockImplementation(
        () =>
          ({
            width: viewportWidth,
            height: viewportHeight,
            top: 0,
            left: 0,
            right: viewportWidth,
            bottom: viewportHeight,
            x: 0,
            y: 0,
            toJSON: () => ({})
          }) as DOMRect
      );

    const ctxStub: Partial<CanvasRenderingContext2D> = {
      setTransform: vi.fn(),
      measureText: (text: string) =>
        ({
          width: text.length * 6,
          actualBoundingBoxAscent: 8,
          actualBoundingBoxDescent: 2
        }) as TextMetrics
    };

    const getContext = vi
      .spyOn(HTMLCanvasElement.prototype, "getContext")
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      .mockImplementation(() => ctxStub as any);

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    const apiRef = React.createRef<GridApi>();

    await act(async () => {
      root.render(
        React.createElement(CanvasGrid, {
          provider: {
            getCell: () => null,
            prefetch
          },
          rowCount,
          colCount,
          defaultRowHeight: rowHeight,
          defaultColWidth: colWidth,
          prefetchOverscanRows: overscanRows,
          prefetchOverscanCols: overscanCols,
          apiRef
        })
      );
    });

    // Scroll to the bottom-right corner (max scroll).
    apiRef.current?.scrollTo(50, 60);

    const lastCall = prefetch.mock.calls.at(-1)?.[0];
    expect(lastCall).toEqual({
      startRow: 4,
      endRow: rowCount,
      startCol: 2,
      endCol: colCount
    });

    await act(async () => {
      root.unmount();
    });
    host.remove();

    boundingRect.mockRestore();
    getContext.mockRestore();
    vi.unstubAllGlobals();
    restoreActEnvironment(previousActEnvironment);
  });

  it("prefetches frozen quadrants separately (without relying on overscan)", async () => {
    const previousActEnvironment = (globalThis as any).IS_REACT_ACT_ENVIRONMENT;
    (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

    const prefetch = vi.fn<(range: CellRange) => void>();

    vi.stubGlobal(
      "ResizeObserver",
      class ResizeObserver {
        observe(_target: Element): void {}
        unobserve(_target: Element): void {}
        disconnect(): void {}
      }
    );

    vi.stubGlobal("requestAnimationFrame", vi.fn((_cb: FrameRequestCallback) => 0));

    const viewportWidth = 150;
    const viewportHeight = 200;

    const overscanRows = 1;
    const overscanCols = 1;

    const rowHeight = 10;
    const colWidth = 10;

    const rowCount = 100;
    const colCount = 100;

    const frozenRows = 12;
    const frozenCols = 8;

    const boundingRect = vi
      .spyOn(HTMLElement.prototype, "getBoundingClientRect")
      .mockImplementation(
        () =>
          ({
            width: viewportWidth,
            height: viewportHeight,
            top: 0,
            left: 0,
            right: viewportWidth,
            bottom: viewportHeight,
            x: 0,
            y: 0,
            toJSON: () => ({})
          }) as DOMRect
      );

    const ctxStub: Partial<CanvasRenderingContext2D> = {
      setTransform: vi.fn(),
      measureText: (text: string) =>
        ({
          width: text.length * 6,
          actualBoundingBoxAscent: 8,
          actualBoundingBoxDescent: 2
        }) as TextMetrics
    };

    const getContext = vi
      .spyOn(HTMLCanvasElement.prototype, "getContext")
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      .mockImplementation(() => ctxStub as any);

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        React.createElement(CanvasGrid, {
          provider: {
            getCell: () => null,
            prefetch
          },
          rowCount,
          colCount,
          frozenRows,
          frozenCols,
          defaultRowHeight: rowHeight,
          defaultColWidth: colWidth,
          prefetchOverscanRows: overscanRows,
          prefetchOverscanCols: overscanCols
        })
      );
    });

    expect(prefetch).toHaveBeenCalledTimes(4);
    expect(prefetch.mock.calls.map((call) => call[0])).toEqual([
      // Frozen (top-left) quadrant.
      { startRow: 0, endRow: frozenRows, startCol: 0, endCol: frozenCols },
      // Frozen rows + scrollable columns (top-right) quadrant.
      { startRow: 0, endRow: frozenRows, startCol: frozenCols, endCol: 16 },
      // Scrollable rows + frozen columns (bottom-left) quadrant.
      { startRow: frozenRows, endRow: 21, startCol: 0, endCol: frozenCols },
      // Main (scrollable) quadrant.
      { startRow: frozenRows, endRow: 21, startCol: frozenCols, endCol: 16 }
    ]);

    await act(async () => {
      root.unmount();
    });
    host.remove();

    boundingRect.mockRestore();
    getContext.mockRestore();
    vi.unstubAllGlobals();
    restoreActEnvironment(previousActEnvironment);
  });

  it("dedupes prefetch calls when the prefetched range does not change", async () => {
    const previousActEnvironment = (globalThis as any).IS_REACT_ACT_ENVIRONMENT;
    (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

    const prefetch = vi.fn<(range: CellRange) => void>();

    vi.stubGlobal(
      "ResizeObserver",
      class ResizeObserver {
        observe(_target: Element): void {}
        unobserve(_target: Element): void {}
        disconnect(): void {}
      }
    );

    // Avoid executing a full render frame; we only care about prefetch calls.
    vi.stubGlobal("requestAnimationFrame", vi.fn((_cb: FrameRequestCallback) => 0));

    const viewportWidth = 50;
    const viewportHeight = 40;

    const rowHeight = 10;
    const colWidth = 10;

    const boundingRect = vi
      .spyOn(HTMLElement.prototype, "getBoundingClientRect")
      .mockImplementation(
        () =>
          ({
            width: viewportWidth,
            height: viewportHeight,
            top: 0,
            left: 0,
            right: viewportWidth,
            bottom: viewportHeight,
            x: 0,
            y: 0,
            toJSON: () => ({})
          }) as DOMRect
      );

    const ctxStub: Partial<CanvasRenderingContext2D> = {
      setTransform: vi.fn(),
      measureText: (text: string) =>
        ({
          width: text.length * 6,
          actualBoundingBoxAscent: 8,
          actualBoundingBoxDescent: 2
        }) as TextMetrics
    };

    const getContext = vi
      .spyOn(HTMLCanvasElement.prototype, "getContext")
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      .mockImplementation(() => ctxStub as any);

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    const apiRef = React.createRef<GridApi>();

    await act(async () => {
      root.render(
        React.createElement(CanvasGrid, {
          provider: {
            getCell: () => null,
            prefetch
          },
          rowCount: 100,
          colCount: 100,
          defaultRowHeight: rowHeight,
          defaultColWidth: colWidth,
          prefetchOverscanRows: 0,
          prefetchOverscanCols: 0,
          apiRef
        })
      );
    });

    const initialPrefetchCalls = prefetch.mock.calls.length;
    expect(apiRef.current).toBeTruthy();

    // Scroll by several pixels. Only the first pixel should change the visible range
    // (from rows [0,4) to [0,5)), subsequent pixels within the same row keep the same range.
    apiRef.current?.scrollBy(0, 1);
    apiRef.current?.scrollBy(0, 1);
    apiRef.current?.scrollBy(0, 1);
    apiRef.current?.scrollBy(0, 1);

    expect(prefetch.mock.calls.length).toBe(initialPrefetchCalls + 1);

    await act(async () => {
      root.unmount();
    });
    host.remove();

    boundingRect.mockRestore();
    getContext.mockRestore();
    vi.unstubAllGlobals();
    restoreActEnvironment(previousActEnvironment);
  });
});
