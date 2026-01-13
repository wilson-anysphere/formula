// @vitest-environment jsdom
import React, { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { CanvasGrid, type GridApi } from "../CanvasGrid";
import * as scrollbarMath from "../../virtualization/scrollbarMath";

// React 18 relies on this flag to suppress act() warnings in test runners.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

function createPointerEvent(
  type: string,
  options: { clientX: number; clientY: number; pointerId: number }
): Event {
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

describe("CanvasGrid scrollbars", () => {
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

  it("scrolls when clicking the vertical scrollbar track and does not start a selection", async () => {
    vi.spyOn(HTMLElement.prototype, "getBoundingClientRect").mockReturnValue(rect200);

    const apiRef = React.createRef<GridApi>();

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<CanvasGrid provider={{ getCell: () => null }} rowCount={100} colCount={10} defaultRowHeight={10} defaultColWidth={10} apiRef={apiRef} />);
    });

    const container = host.querySelector('[data-testid="canvas-grid"]') as HTMLDivElement;
    const [vTrack] = Array.from(container.querySelectorAll<HTMLDivElement>('div[aria-hidden="true"]'));
    expect(vTrack).toBeTruthy();

    expect(apiRef.current?.getScroll().y).toBe(0);
    expect(apiRef.current?.getSelection()).toBeNull();

    await act(async () => {
      vTrack.dispatchEvent(createPointerEvent("pointerdown", { clientX: 0, clientY: 150, pointerId: 1 }));
    });

    expect(apiRef.current?.getScroll().y).toBeGreaterThan(0);
    expect(apiRef.current?.getSelection()).toBeNull();

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });

  it("scrolls when clicking the horizontal scrollbar track and does not start a selection", async () => {
    vi.spyOn(HTMLElement.prototype, "getBoundingClientRect").mockReturnValue(rect200);

    const apiRef = React.createRef<GridApi>();

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<CanvasGrid provider={{ getCell: () => null }} rowCount={10} colCount={100} defaultRowHeight={10} defaultColWidth={10} apiRef={apiRef} />);
    });

    const container = host.querySelector('[data-testid="canvas-grid"]') as HTMLDivElement;
    const tracks = Array.from(container.querySelectorAll<HTMLDivElement>('div[aria-hidden="true"]'));
    const hTrack = tracks[1];
    expect(hTrack).toBeTruthy();

    expect(apiRef.current?.getScroll().x).toBe(0);
    expect(apiRef.current?.getSelection()).toBeNull();

    await act(async () => {
      hTrack.dispatchEvent(createPointerEvent("pointerdown", { clientX: 150, clientY: 0, pointerId: 1 }));
    });

    expect(apiRef.current?.getScroll().x).toBeGreaterThan(0);
    expect(apiRef.current?.getSelection()).toBeNull();

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });

  it("treats shift+wheel as horizontal scroll", async () => {
    vi.spyOn(HTMLElement.prototype, "getBoundingClientRect").mockReturnValue(rect200);

    const apiRef = React.createRef<GridApi>();

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<CanvasGrid provider={{ getCell: () => null }} rowCount={10} colCount={100} defaultRowHeight={10} defaultColWidth={10} apiRef={apiRef} />);
    });

    const container = host.querySelector('[data-testid="canvas-grid"]') as HTMLDivElement;
    expect(apiRef.current?.getScroll()).toEqual({ x: 0, y: 0 });

    await act(async () => {
      container.dispatchEvent(
        new WheelEvent("wheel", { deltaX: 0, deltaY: 120, shiftKey: true, bubbles: true, cancelable: true })
      );
    });

    const scroll = apiRef.current?.getScroll();
    expect(scroll?.x).toBeGreaterThan(0);
    expect(scroll?.y).toBe(0);

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });

  it("normalizes wheel deltaMode for line and page scrolling", async () => {
    vi.spyOn(HTMLElement.prototype, "getBoundingClientRect").mockReturnValue(rect200);

    const apiRef = React.createRef<GridApi>();

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<CanvasGrid provider={{ getCell: () => null }} rowCount={100} colCount={10} defaultRowHeight={10} defaultColWidth={10} apiRef={apiRef} />);
    });

    const container = host.querySelector('[data-testid="canvas-grid"]') as HTMLDivElement;

    await act(async () => {
      container.dispatchEvent(new WheelEvent("wheel", { deltaY: 1, deltaMode: 1, bubbles: true, cancelable: true }));
    });
    expect(apiRef.current?.getScroll().y).toBe(16);

    await act(async () => {
      apiRef.current?.scrollTo(0, 0);
    });

    await act(async () => {
      container.dispatchEvent(new WheelEvent("wheel", { deltaY: 1, deltaMode: 2, bubbles: true, cancelable: true }));
    });
    expect(apiRef.current?.getScroll().y).toBe(200);

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });

  it("uses ctrl+wheel to zoom the grid", async () => {
    const rectSpy = vi.spyOn(HTMLElement.prototype, "getBoundingClientRect").mockReturnValue(rect200);

    const apiRef = React.createRef<GridApi>();

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<CanvasGrid provider={{ getCell: () => null }} rowCount={100} colCount={10} defaultRowHeight={10} defaultColWidth={10} apiRef={apiRef} />);
    });

    const container = host.querySelector('[data-testid="canvas-grid"]') as HTMLDivElement;
    expect(apiRef.current?.getZoom()).toBe(1);

    // Initial mount/resize may read layout; we only care about the ctrl+wheel zoom path.
    rectSpy.mockClear();

    await act(async () => {
      container.dispatchEvent(
        new WheelEvent("wheel", { deltaY: -100, ctrlKey: true, clientX: 100, clientY: 100, bubbles: true, cancelable: true })
      );
    });

    expect(rectSpy).not.toHaveBeenCalled();

    const zoom = apiRef.current?.getZoom() ?? 0;
    expect(zoom).toBeGreaterThan(1);
    expect(zoom).toBeCloseTo(Math.exp(0.1), 3);
    expect(apiRef.current?.getColWidth(0)).toBeCloseTo(10 * zoom, 3);

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });

  it("uses meta+wheel to zoom the grid", async () => {
    const rectSpy = vi.spyOn(HTMLElement.prototype, "getBoundingClientRect").mockReturnValue(rect200);

    const apiRef = React.createRef<GridApi>();

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<CanvasGrid provider={{ getCell: () => null }} rowCount={100} colCount={10} defaultRowHeight={10} defaultColWidth={10} apiRef={apiRef} />);
    });

    const container = host.querySelector('[data-testid="canvas-grid"]') as HTMLDivElement;
    expect(apiRef.current?.getZoom()).toBe(1);

    // Initial mount/resize may read layout; we only care about the meta+wheel zoom path.
    rectSpy.mockClear();

    await act(async () => {
      container.dispatchEvent(
        new WheelEvent("wheel", { deltaY: -100, metaKey: true, clientX: 100, clientY: 100, bubbles: true, cancelable: true })
      );
    });

    expect(rectSpy).not.toHaveBeenCalled();

    const zoom = apiRef.current?.getZoom() ?? 0;
    expect(zoom).toBeGreaterThan(1);
    expect(zoom).toBeCloseTo(Math.exp(0.1), 3);
    expect(apiRef.current?.getColWidth(0)).toBeCloseTo(10 * zoom, 3);

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });

  it("dragging the vertical scrollbar thumb respects zoom-scaled minimum thumb size", async () => {
    vi.spyOn(HTMLElement.prototype, "getBoundingClientRect").mockReturnValue(rect200);

    const apiRef = React.createRef<GridApi>();

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<CanvasGrid provider={{ getCell: () => null }} rowCount={100} colCount={10} defaultRowHeight={10} defaultColWidth={10} apiRef={apiRef} />);
    });

    await act(async () => {
      apiRef.current?.setZoom(2);
    });

    expect(apiRef.current?.getZoom()).toBe(2);

    const container = host.querySelector('[data-testid="canvas-grid"]') as HTMLDivElement;
    const tracks = Array.from(container.querySelectorAll<HTMLDivElement>('div[aria-hidden="true"]'));
    const vTrack = tracks[0];
    expect(vTrack).toBeTruthy();
    const vThumb = vTrack.querySelector("div") as HTMLDivElement;
    expect(vThumb).toBeTruthy();

    expect(apiRef.current?.getScroll().y).toBe(0);

    // With zoom=2: totalHeight = 100 * (10 * 2) = 2000, viewportHeight=200, maxScrollY=1800.
    // TrackSize is mocked to 200. rawThumb = (200/2000)*200 = 20, minThumb = 24*2=48 -> thumbSize=48
    // thumbTravel = 200 - 48 = 152. Dragging by thumbTravel should reach max scroll.
    await act(async () => {
      vThumb.dispatchEvent(createPointerEvent("pointerdown", { clientX: 0, clientY: 0, pointerId: 1 }));
      window.dispatchEvent(createPointerEvent("pointermove", { clientX: 0, clientY: 152, pointerId: 1 }));
      window.dispatchEvent(createPointerEvent("pointerup", { clientX: 0, clientY: 152, pointerId: 1 }));
    });

    expect(apiRef.current?.getScroll().y).toBe(1800);

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });

  it("syncScrollbars derives track sizes from viewport metrics (no scrollbars)", async () => {
    vi.spyOn(HTMLElement.prototype, "getBoundingClientRect").mockReturnValue(rect200);
    const thumbSpy = vi.spyOn(scrollbarMath, "computeScrollbarThumb");

    const apiRef = React.createRef<GridApi>();
    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <CanvasGrid
          provider={{ getCell: () => null }}
          rowCount={10}
          colCount={10}
          defaultRowHeight={10}
          defaultColWidth={10}
          apiRef={apiRef}
        />
      );
    });

    const viewport = apiRef.current?.getViewportState();
    expect(viewport).toBeTruthy();

    thumbSpy.mockClear();
    await act(async () => {
      // Trigger a scrollbar sync after the initial mount effects.
      apiRef.current?.scrollBy(0, 0);
    });

    const zoom = apiRef.current?.getZoom() ?? 1;
    const inset = 2 * zoom;
    const thickness = 10 * zoom;
    const gap = 4 * zoom;
    const corner = inset + thickness + gap;

    expect(thumbSpy).toHaveBeenCalledTimes(2);

    const vArgs = thumbSpy.mock.calls[0]?.[0];
    const hArgs = thumbSpy.mock.calls[1]?.[0];
    expect(vArgs).toBeTruthy();
    expect(hArgs).toBeTruthy();

    expect(vArgs?.trackSize).toBeCloseTo(viewport!.height - inset - corner, 6);
    expect(vArgs?.viewportSize).toBeCloseTo(viewport!.height - viewport!.frozenHeight, 6);
    expect(vArgs?.contentSize).toBeCloseTo(viewport!.totalHeight - viewport!.frozenHeight, 6);

    expect(hArgs?.trackSize).toBeCloseTo(viewport!.width - inset - corner, 6);
    expect(hArgs?.viewportSize).toBeCloseTo(viewport!.width - viewport!.frozenWidth, 6);
    expect(hArgs?.contentSize).toBeCloseTo(viewport!.totalWidth - viewport!.frozenWidth, 6);

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });

  it("syncScrollbars passes correct thumb inputs for vertical-only scroll", async () => {
    vi.spyOn(HTMLElement.prototype, "getBoundingClientRect").mockReturnValue(rect200);
    const thumbSpy = vi.spyOn(scrollbarMath, "computeScrollbarThumb");

    const apiRef = React.createRef<GridApi>();
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
        />
      );
    });

    const viewport = apiRef.current?.getViewportState();
    expect(viewport).toBeTruthy();
    expect(viewport?.maxScrollY).toBeGreaterThan(0);
    expect(viewport?.maxScrollX).toBe(0);

    thumbSpy.mockClear();
    await act(async () => {
      apiRef.current?.scrollBy(0, 0);
    });

    const zoom = apiRef.current?.getZoom() ?? 1;
    const inset = 2 * zoom;
    const thickness = 10 * zoom;
    const gap = 4 * zoom;
    const corner = inset + thickness + gap;

    expect(thumbSpy).toHaveBeenCalledTimes(2);
    const vArgs = thumbSpy.mock.calls[0]?.[0];
    const hArgs = thumbSpy.mock.calls[1]?.[0];

    expect(vArgs?.trackSize).toBeCloseTo(viewport!.height - inset - corner, 6);
    expect(vArgs?.viewportSize).toBeCloseTo(viewport!.height - viewport!.frozenHeight, 6);
    expect(vArgs?.contentSize).toBeCloseTo(viewport!.totalHeight - viewport!.frozenHeight, 6);

    expect(hArgs?.trackSize).toBeCloseTo(viewport!.width - inset - corner, 6);
    expect(hArgs?.viewportSize).toBeCloseTo(viewport!.width - viewport!.frozenWidth, 6);
    expect(hArgs?.contentSize).toBeCloseTo(viewport!.totalWidth - viewport!.frozenWidth, 6);

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });

  it("syncScrollbars passes correct thumb inputs for both-axis scroll", async () => {
    vi.spyOn(HTMLElement.prototype, "getBoundingClientRect").mockReturnValue(rect200);
    const thumbSpy = vi.spyOn(scrollbarMath, "computeScrollbarThumb");

    const apiRef = React.createRef<GridApi>();
    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <CanvasGrid
          provider={{ getCell: () => null }}
          rowCount={100}
          colCount={100}
          defaultRowHeight={10}
          defaultColWidth={10}
          apiRef={apiRef}
        />
      );
    });

    const viewport = apiRef.current?.getViewportState();
    expect(viewport).toBeTruthy();
    expect(viewport?.maxScrollY).toBeGreaterThan(0);
    expect(viewport?.maxScrollX).toBeGreaterThan(0);

    thumbSpy.mockClear();
    await act(async () => {
      apiRef.current?.scrollBy(0, 0);
    });

    const zoom = apiRef.current?.getZoom() ?? 1;
    const inset = 2 * zoom;
    const thickness = 10 * zoom;
    const gap = 4 * zoom;
    const corner = inset + thickness + gap;

    expect(thumbSpy).toHaveBeenCalledTimes(2);
    const vArgs = thumbSpy.mock.calls[0]?.[0];
    const hArgs = thumbSpy.mock.calls[1]?.[0];

    expect(vArgs?.trackSize).toBeCloseTo(viewport!.height - inset - corner, 6);
    expect(vArgs?.viewportSize).toBeCloseTo(viewport!.height - viewport!.frozenHeight, 6);
    expect(vArgs?.contentSize).toBeCloseTo(viewport!.totalHeight - viewport!.frozenHeight, 6);

    expect(hArgs?.trackSize).toBeCloseTo(viewport!.width - inset - corner, 6);
    expect(hArgs?.viewportSize).toBeCloseTo(viewport!.width - viewport!.frozenWidth, 6);
    expect(hArgs?.contentSize).toBeCloseTo(viewport!.totalWidth - viewport!.frozenWidth, 6);

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });

  it("syncScrollbars passes correct thumb inputs with frozen panes", async () => {
    vi.spyOn(HTMLElement.prototype, "getBoundingClientRect").mockReturnValue(rect200);
    const thumbSpy = vi.spyOn(scrollbarMath, "computeScrollbarThumb");

    const apiRef = React.createRef<GridApi>();
    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <CanvasGrid
          provider={{ getCell: () => null }}
          rowCount={100}
          colCount={100}
          frozenRows={2}
          frozenCols={2}
          defaultRowHeight={10}
          defaultColWidth={10}
          apiRef={apiRef}
        />
      );
    });

    const viewport = apiRef.current?.getViewportState();
    expect(viewport).toBeTruthy();
    expect(viewport?.frozenHeight).toBeGreaterThan(0);
    expect(viewport?.frozenWidth).toBeGreaterThan(0);

    thumbSpy.mockClear();
    await act(async () => {
      apiRef.current?.scrollBy(0, 0);
    });

    const zoom = apiRef.current?.getZoom() ?? 1;
    const inset = 2 * zoom;
    const thickness = 10 * zoom;
    const gap = 4 * zoom;
    const corner = inset + thickness + gap;

    expect(thumbSpy).toHaveBeenCalledTimes(2);
    const vArgs = thumbSpy.mock.calls[0]?.[0];
    const hArgs = thumbSpy.mock.calls[1]?.[0];

    expect(vArgs?.trackSize).toBeCloseTo(viewport!.height - inset - corner, 6);
    expect(vArgs?.viewportSize).toBeCloseTo(viewport!.height - viewport!.frozenHeight, 6);
    expect(vArgs?.contentSize).toBeCloseTo(viewport!.totalHeight - viewport!.frozenHeight, 6);

    expect(hArgs?.trackSize).toBeCloseTo(viewport!.width - inset - corner, 6);
    expect(hArgs?.viewportSize).toBeCloseTo(viewport!.width - viewport!.frozenWidth, 6);
    expect(hArgs?.contentSize).toBeCloseTo(viewport!.totalWidth - viewport!.frozenWidth, 6);

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });

  it("syncScrollbars avoids getBoundingClientRect during scroll updates", async () => {
    const rectSpy = vi.spyOn(HTMLElement.prototype, "getBoundingClientRect").mockReturnValue(rect200);

    const apiRef = React.createRef<GridApi>();
    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <CanvasGrid provider={{ getCell: () => null }} rowCount={100} colCount={100} defaultRowHeight={10} defaultColWidth={10} apiRef={apiRef} />
      );
    });

    rectSpy.mockClear();

    await act(async () => {
      apiRef.current?.scrollBy(10, 10);
    });

    expect(rectSpy).not.toHaveBeenCalled();

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });
});
