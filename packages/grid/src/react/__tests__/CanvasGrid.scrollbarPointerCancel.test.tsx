// @vitest-environment jsdom
import React, { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { CanvasGrid, type GridApi } from "../CanvasGrid";

// React 18 relies on this flag to suppress act() warnings in test runners.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

function createPointerEvent(type: string, options: { clientX: number; clientY: number; pointerId: number }): Event {
  const PointerEventCtor = (window as unknown as { PointerEvent?: typeof PointerEvent }).PointerEvent;
  if (PointerEventCtor) {
    return new PointerEventCtor(type, {
      bubbles: true,
      cancelable: true,
      clientX: options.clientX,
      clientY: options.clientY,
      buttons: 1,
      pointerId: options.pointerId
    } as PointerEventInit);
  }

  const event = new MouseEvent(type, {
    bubbles: true,
    cancelable: true,
    clientX: options.clientX,
    clientY: options.clientY,
    buttons: 1
  });
  Object.defineProperty(event, "pointerId", { value: options.pointerId });
  return event;
}

describe("CanvasGrid scrollbar thumb pointercancel", () => {
  const rect200 = {
    left: 0,
    top: 0,
    right: 200,
    bottom: 200,
    width: 200,
    height: 200,
    x: 0,
    y: 0,
    toJSON: () => ({})
  } as unknown as DOMRect;

  beforeEach(() => {
    vi.stubGlobal(
      "ResizeObserver",
      class ResizeObserver {
        observe(): void {}
        unobserve(): void {}
        disconnect(): void {}
      }
    );

    // Avoid running full render frames; these tests only validate scroll behavior.
    vi.stubGlobal("requestAnimationFrame", vi.fn((_cb: FrameRequestCallback) => 0));

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

  it("stops scrolling the vertical thumb drag after pointercancel", async () => {
    vi.spyOn(HTMLElement.prototype, "getBoundingClientRect").mockReturnValue(rect200);

    const apiRef = React.createRef<GridApi>();
    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <CanvasGrid provider={{ getCell: () => null }} rowCount={100} colCount={10} defaultRowHeight={10} defaultColWidth={10} apiRef={apiRef} />
      );
    });

    const container = host.querySelector('[data-testid="canvas-grid"]') as HTMLDivElement;
    const [vTrack] = Array.from(container.querySelectorAll<HTMLDivElement>('div[aria-hidden="true"]'));
    const vThumb = vTrack?.querySelector("div") as HTMLDivElement | null;
    expect(vThumb).toBeTruthy();

    expect(apiRef.current?.getScroll().y).toBe(0);

    await act(async () => {
      vThumb!.dispatchEvent(createPointerEvent("pointerdown", { clientX: 0, clientY: 0, pointerId: 1 }));
    });

    await act(async () => {
      window.dispatchEvent(createPointerEvent("pointermove", { clientX: 0, clientY: 50, pointerId: 1 }));
    });
    const duringDrag = apiRef.current?.getScroll().y ?? 0;
    expect(duringDrag).toBeGreaterThan(0);

    await act(async () => {
      window.dispatchEvent(createPointerEvent("pointercancel", { clientX: 0, clientY: 50, pointerId: 1 }));
    });
    const afterCancel = apiRef.current?.getScroll().y ?? 0;

    await act(async () => {
      window.dispatchEvent(createPointerEvent("pointermove", { clientX: 0, clientY: 150, pointerId: 1 }));
    });

    expect(apiRef.current?.getScroll().y).toBe(afterCancel);

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });
});

