// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { MockCellProvider } from "@formula/grid";
import { DesktopSharedGrid } from "../desktopSharedGrid";

function createMock2dContext(canvas: HTMLCanvasElement): CanvasRenderingContext2D {
  const noop = () => {};

  // CanvasGridRenderer touches a fairly wide surface area of the 2D context. For
  // a11y unit tests we only need attach/resize/selection calls to succeed, so a
  // no-op context is sufficient.
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
    setLineDash: noop,
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

describe("DesktopSharedGrid a11y", () => {
  const originalGetContext = HTMLCanvasElement.prototype.getContext;

  beforeEach(() => {
    // Some parts of the grid renderer use rAF for scheduling; execute immediately.
    vi.stubGlobal("requestAnimationFrame", (cb: FrameRequestCallback) => {
      cb(0);
      return 0;
    });
    vi.stubGlobal("cancelAnimationFrame", () => {});

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      return createMock2dContext(this);
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;
  });

  afterEach(() => {
    HTMLCanvasElement.prototype.getContext = originalGetContext;
    vi.unstubAllGlobals();
    document.body.innerHTML = "";
  });

  it("exposes the active cell via aria-activedescendant with stable row/col indices", () => {
    const container = document.createElement("div");
    document.body.appendChild(container);

    const provider = new MockCellProvider({ rowCount: 10, colCount: 10 });

    const canvases = {
      grid: document.createElement("canvas"),
      content: document.createElement("canvas"),
      selection: document.createElement("canvas")
    };

    const scrollbars = {
      vTrack: document.createElement("div"),
      vThumb: document.createElement("div"),
      hTrack: document.createElement("div"),
      hThumb: document.createElement("div")
    };

    const grid = new DesktopSharedGrid({
      container,
      provider,
      rowCount: 10,
      colCount: 10,
      canvases,
      scrollbars
    });

    grid.resize(300, 200, 1);

    grid.setSelectionRanges([{ startRow: 2, endRow: 3, startCol: 4, endCol: 5 }], { activeCell: { row: 2, col: 4 } });

    const activeId = container.getAttribute("aria-activedescendant");
    expect(activeId).toBeTruthy();

    const activeCell = container.querySelector('[data-testid="canvas-grid-a11y-active-cell"]') as HTMLDivElement | null;
    expect(activeCell).not.toBeNull();
    expect(activeCell?.id).toBe(activeId);
    expect(activeCell?.getAttribute("role")).toBe("gridcell");
    expect(activeCell?.getAttribute("aria-rowindex")).toBe("3");
    expect(activeCell?.getAttribute("aria-colindex")).toBe("5");
    expect(activeCell?.getAttribute("aria-selected")).toBe("true");

    // Update selection and ensure row/col indices change (aria-activedescendant stays stable).
    grid.setSelectionRanges([{ startRow: 0, endRow: 1, startCol: 0, endCol: 1 }], { activeCell: { row: 0, col: 0 } });
    expect(container.getAttribute("aria-activedescendant")).toBe(activeId);
    expect(activeCell?.getAttribute("aria-rowindex")).toBe("1");
    expect(activeCell?.getAttribute("aria-colindex")).toBe("1");

    // Clearing selection should remove aria-activedescendant but keep the SR-only active-cell element mounted.
    grid.setSelectionRanges(null);
    expect(container.getAttribute("aria-activedescendant")).toBeNull();
    expect(container.querySelector(`#${activeId}`)).toBe(activeCell);
    expect(activeCell?.getAttribute("aria-hidden")).toBe("true");

    grid.destroy();
  });
});

