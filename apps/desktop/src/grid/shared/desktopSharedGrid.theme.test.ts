// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import type { CellProvider } from "@formula/grid";
import { DesktopSharedGrid } from "./desktopSharedGrid";

function createMock2dContext(canvas: HTMLCanvasElement): CanvasRenderingContext2D {
  const noop = () => {};
  let fillStyle: string | CanvasGradient | CanvasPattern = "#000";
  let strokeStyle: string | CanvasGradient | CanvasPattern = "#000";

  return {
    canvas,
    get fillStyle() {
      return fillStyle;
    },
    set fillStyle(value) {
      fillStyle = value;
    },
    get strokeStyle() {
      return strokeStyle;
    },
    set strokeStyle(value) {
      strokeStyle = value;
    },
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

describe("DesktopSharedGrid theme watching", () => {
  const originalGetContext = HTMLCanvasElement.prototype.getContext;
  const originalRaf = globalThis.requestAnimationFrame;

  beforeEach(() => {
    vi.stubGlobal("requestAnimationFrame", (cb: FrameRequestCallback) => {
      cb(0);
      return 0;
    });
  });

  afterEach(() => {
    HTMLCanvasElement.prototype.getContext = originalGetContext;
    if (originalRaf) {
      vi.stubGlobal("requestAnimationFrame", originalRaf);
    } else {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      delete (globalThis as any).requestAnimationFrame;
    }
    vi.unstubAllGlobals();
  });

  it("refreshes renderer theme when documentElement data-theme changes", async () => {
    const container = document.createElement("div");
    document.body.appendChild(container);

    const provider: CellProvider = {
      getCell: () => null
    };

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");

    const contexts = new Map<HTMLCanvasElement, CanvasRenderingContext2D>();
    contexts.set(gridCanvas, createMock2dContext(gridCanvas));
    contexts.set(contentCanvas, createMock2dContext(contentCanvas));
    contexts.set(selectionCanvas, createMock2dContext(selectionCanvas));

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      return contexts.get(this) ?? createMock2dContext(this);
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

    const sharedGrid = new DesktopSharedGrid({
      container,
      provider,
      rowCount: 1,
      colCount: 1,
      canvases: { grid: gridCanvas, content: contentCanvas, selection: selectionCanvas },
      scrollbars: {
        vTrack: document.createElement("div"),
        vThumb: document.createElement("div"),
        hTrack: document.createElement("div"),
        hThumb: document.createElement("div")
      },
      enableWheel: false,
      enableKeyboard: false,
      enableResize: false
    });

    try {
      const setThemeSpy = vi.spyOn(sharedGrid.renderer, "setTheme");

      document.documentElement.setAttribute("data-theme", "dark");
      await new Promise((resolve) => setTimeout(resolve, 0));

      expect(setThemeSpy).toHaveBeenCalled();

      setThemeSpy.mockClear();
      sharedGrid.destroy();

      document.documentElement.setAttribute("data-theme", "light");
      await new Promise((resolve) => setTimeout(resolve, 0));

      expect(setThemeSpy).not.toHaveBeenCalled();
    } finally {
      // Ensure cleanup even if the test fails mid-flight.
      sharedGrid.destroy();
      container.remove();
    }
  });
});

