// @vitest-environment jsdom
import React, { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { CanvasGrid, type GridApi } from "../CanvasGrid";

// React 18 relies on this flag to suppress act() warnings in test runners.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

describe("CanvasGrid onScroll", () => {
  beforeEach(() => {
    vi.stubGlobal(
      "ResizeObserver",
      class ResizeObserver {
        observe(): void {}
        unobserve(): void {}
        disconnect(): void {}
      }
    );

    // Avoid running full render frames; these tests only validate scroll callback behavior.
    vi.stubGlobal("requestAnimationFrame", vi.fn((_cb: FrameRequestCallback) => 0));

    vi.spyOn(HTMLElement.prototype, "getBoundingClientRect").mockReturnValue({
      left: 0,
      top: 0,
      right: 200,
      bottom: 200,
      width: 200,
      height: 200,
      x: 0,
      y: 0,
      toJSON: () => ({})
    } as unknown as DOMRect);

    const ctxStub: Partial<CanvasRenderingContext2D> = {
      setTransform: vi.fn(),
      measureText: (text: string) =>
        ({
          width: text.length * 6,
          actualBoundingBoxAscent: 8,
          actualBoundingBoxDescent: 2
        }) as TextMetrics
    };

    vi.spyOn(HTMLCanvasElement.prototype, "getContext").mockImplementation(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      () => ctxStub as any
    );
  });

  afterEach(() => {
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it("fires onScroll for imperative scroll calls and wheel events without redundant emissions", async () => {
    const apiRef = React.createRef<GridApi>();
    const onScroll = vi.fn();

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <CanvasGrid
          provider={{ getCell: () => null }}
          rowCount={100}
          colCount={10}
          defaultRowHeight={10}
          defaultColWidth={10}
          apiRef={apiRef}
          onScroll={onScroll}
        />
      );
    });

    expect(onScroll).toHaveBeenCalledTimes(0);

    await act(async () => {
      apiRef.current?.scrollBy(0, 50);
    });

    expect(onScroll).toHaveBeenCalledTimes(1);
    expect(onScroll.mock.calls[0]?.[0]).toEqual({ x: 0, y: 50 });

    await act(async () => {
      apiRef.current?.scrollBy(0, 0);
    });

    expect(onScroll).toHaveBeenCalledTimes(1);

    const container = host.querySelector('[data-testid="canvas-grid"]') as HTMLDivElement;
    expect(container).toBeTruthy();

    await act(async () => {
      container.dispatchEvent(new WheelEvent("wheel", { deltaY: 120, bubbles: true, cancelable: true }));
    });

    expect(onScroll).toHaveBeenCalledTimes(2);
    expect(onScroll.mock.calls[1]?.[0]?.y).toBeGreaterThan(50);

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });
});

