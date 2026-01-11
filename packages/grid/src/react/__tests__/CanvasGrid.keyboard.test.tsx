// @vitest-environment jsdom
import React, { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { CanvasGrid, type GridApi } from "../CanvasGrid";

// React 18 relies on this flag to suppress act() warnings in test runners.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

function createMock2dContext(canvas: HTMLCanvasElement): CanvasRenderingContext2D {
  const noop = () => {};
  return {
    canvas,
    fillStyle: "#000",
    strokeStyle: "#000",
    lineWidth: 1,
    font: "",
    textAlign: "left",
    textBaseline: "alphabetic",
    globalAlpha: 1,
    imageSmoothingEnabled: false,
    setTransform: noop,
    clearRect: noop,
    fillRect: noop,
    strokeRect: noop,
    beginPath: noop,
    rect: noop,
    clip: noop,
    fill: noop,
    stroke: noop,
    moveTo: noop,
    lineTo: noop,
    closePath: noop,
    save: noop,
    restore: noop,
    drawImage: noop,
    translate: noop,
    rotate: noop,
    fillText: noop,
    measureText: (text: string) =>
      ({
        width: text.length * 6,
        actualBoundingBoxAscent: 8,
        actualBoundingBoxDescent: 2
      }) as TextMetrics
  } as unknown as CanvasRenderingContext2D;
}

describe("CanvasGrid keyboard navigation", () => {
  const originalGetContext = HTMLCanvasElement.prototype.getContext;

  beforeEach(() => {
    vi.stubGlobal(
      "ResizeObserver",
      class ResizeObserver {
        observe(): void {}
        unobserve(): void {}
        disconnect(): void {}
      }
    );

    vi.stubGlobal("requestAnimationFrame", (cb: FrameRequestCallback) => {
      cb(0);
      return 0;
    });

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      return createMock2dContext(this);
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

    vi.spyOn(HTMLElement.prototype, "getBoundingClientRect").mockReturnValue({
      left: 0,
      top: 0,
      right: 400,
      bottom: 200,
      width: 400,
      height: 200,
      x: 0,
      y: 0,
      toJSON: () => {}
    } as DOMRect);
  });

  afterEach(() => {
    HTMLCanvasElement.prototype.getContext = originalGetContext;
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it("moves selection with arrow keys when the grid container is focused", async () => {
    const apiRef = React.createRef<GridApi>();
    const onSelectionChange = vi.fn();
    const onSelectionRangeChange = vi.fn();

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <CanvasGrid
          provider={{ getCell: (row, col) => ({ row, col, value: `${row},${col}` }) }}
          rowCount={10}
          colCount={10}
          apiRef={apiRef}
          onSelectionChange={onSelectionChange}
          onSelectionRangeChange={onSelectionRangeChange}
        />
      );
    });

    // Seed selection at A1 (row 0, col 0) and clear initial calls.
    await act(async () => {
      apiRef.current?.setSelection(0, 0);
    });
    onSelectionChange.mockClear();
    onSelectionRangeChange.mockClear();

    const container = host.querySelector('[data-testid="canvas-grid"]') as HTMLDivElement;
    expect(container).toBeTruthy();
    container.focus();

    await act(async () => {
      container.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowRight", bubbles: true, cancelable: true }));
    });

    expect(onSelectionChange).toHaveBeenCalledWith({ row: 0, col: 1 });
    expect(onSelectionRangeChange).toHaveBeenCalledWith({ startRow: 0, endRow: 1, startCol: 1, endCol: 2 });

    const status = host.querySelector('[data-testid="canvas-grid-a11y-status"]') as HTMLDivElement;
    expect(status.textContent).toContain("Active cell B1");
    expect(status.textContent).toContain("value 0,1");

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });
});

