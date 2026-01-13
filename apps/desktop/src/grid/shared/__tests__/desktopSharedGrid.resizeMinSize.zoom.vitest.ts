// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { MockCellProvider } from "@formula/grid";

import { DesktopSharedGrid } from "../desktopSharedGrid";

function createMock2dContext(canvas: HTMLCanvasElement): CanvasRenderingContext2D {
  const noop = () => {};

  // DesktopSharedGrid relies on CanvasGridRenderer, which touches a broad surface
  // area of the 2D canvas context. For resize unit tests, a no-op context is
  // sufficient as long as the used methods/properties exist.
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
        actualBoundingBoxDescent: 2,
      }) as TextMetrics,
  } as unknown as CanvasRenderingContext2D;
}

function createPointerLikeMouseEvent(
  type: string,
  options: {
    clientX: number;
    clientY: number;
    button?: number;
    pointerId?: number;
    pointerType?: string;
  },
): MouseEvent {
  const event = new MouseEvent(type, {
    bubbles: true,
    cancelable: true,
    clientX: options.clientX,
    clientY: options.clientY,
    button: options.button ?? 0,
  });
  Object.defineProperty(event, "pointerId", { configurable: true, value: options.pointerId ?? 1 });
  Object.defineProperty(event, "pointerType", { configurable: true, value: options.pointerType ?? "mouse" });
  return event;
}

describe("DesktopSharedGrid resize min size clamps respect zoom", () => {
  const originalGetContext = HTMLCanvasElement.prototype.getContext;

  beforeEach(() => {
    vi.stubGlobal("requestAnimationFrame", (cb: FrameRequestCallback) => {
      cb(0);
      return 0;
    });
    vi.stubGlobal("cancelAnimationFrame", () => {});

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      return createMock2dContext(this);
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;

    document.body.innerHTML = "";
  });

  afterEach(() => {
    HTMLCanvasElement.prototype.getContext = originalGetContext;
    vi.unstubAllGlobals();
    document.body.innerHTML = "";
  });

  it("clamps column/row resizing to MIN_* * zoom (matching CanvasGrid)", () => {
    const container = document.createElement("div");
    document.body.appendChild(container);

    const gridCanvas = document.createElement("canvas");
    const contentCanvas = document.createElement("canvas");
    const selectionCanvas = document.createElement("canvas");
    container.appendChild(gridCanvas);
    container.appendChild(contentCanvas);
    container.appendChild(selectionCanvas);

    const vTrack = document.createElement("div");
    const vThumb = document.createElement("div");
    const hTrack = document.createElement("div");
    const hThumb = document.createElement("div");
    container.appendChild(vTrack);
    container.appendChild(vThumb);
    container.appendChild(hTrack);
    container.appendChild(hThumb);

    vi.spyOn(selectionCanvas, "getBoundingClientRect").mockReturnValue({
      left: 0,
      top: 0,
      right: 400,
      bottom: 200,
      width: 400,
      height: 200,
      x: 0,
      y: 0,
      toJSON: () => {},
    } as DOMRect);

    const grid = new DesktopSharedGrid({
      container,
      provider: new MockCellProvider({ rowCount: 10, colCount: 10 }),
      rowCount: 10,
      colCount: 10,
      frozenRows: 1,
      frozenCols: 1,
      defaultRowHeight: 24,
      defaultColWidth: 100,
      canvases: { grid: gridCanvas, content: contentCanvas, selection: selectionCanvas },
      scrollbars: { vTrack, vThumb, hTrack, hThumb },
      enableResize: true,
      enableWheel: false,
      enableKeyboard: false,
    });

    grid.resize(400, 200, 1);
    grid.renderer.setColWidth(0, 48);
    grid.renderer.setRowHeight(0, 24);
    grid.setZoom(2);

    const zoom = grid.getZoom();
    expect(zoom).toBe(2);

    // ----- Column resize: attempt to shrink col 1 below min -----
    const headerColWidth = grid.renderer.getColWidth(0);
    const col1Width = grid.renderer.getColWidth(1);
    const headerRowHeight = grid.renderer.getRowHeight(0);

    const col1BoundaryX = headerColWidth + col1Width;
    const headerRowY = headerRowHeight / 2;

    selectionCanvas.dispatchEvent(
      createPointerLikeMouseEvent("pointerdown", {
        clientX: col1BoundaryX,
        clientY: headerRowY,
        pointerId: 1,
      }),
    );
    selectionCanvas.dispatchEvent(
      createPointerLikeMouseEvent("pointermove", {
        clientX: col1BoundaryX - 10_000,
        clientY: headerRowY,
        pointerId: 1,
      }),
    );
    selectionCanvas.dispatchEvent(
      createPointerLikeMouseEvent("pointerup", {
        clientX: col1BoundaryX - 10_000,
        clientY: headerRowY,
        pointerId: 1,
      }),
    );

    expect(grid.renderer.getColWidth(1)).toBe(24 * zoom);

    // ----- Row resize: attempt to shrink row 1 below min -----
    const row1Height = grid.renderer.getRowHeight(1);
    const row1BoundaryY = headerRowHeight + row1Height;
    const rowHeaderX = headerColWidth / 2;

    selectionCanvas.dispatchEvent(
      createPointerLikeMouseEvent("pointerdown", {
        clientX: rowHeaderX,
        clientY: row1BoundaryY,
        pointerId: 2,
      }),
    );
    selectionCanvas.dispatchEvent(
      createPointerLikeMouseEvent("pointermove", {
        clientX: rowHeaderX,
        clientY: row1BoundaryY - 10_000,
        pointerId: 2,
      }),
    );
    selectionCanvas.dispatchEvent(
      createPointerLikeMouseEvent("pointerup", {
        clientX: rowHeaderX,
        clientY: row1BoundaryY - 10_000,
        pointerId: 2,
      }),
    );

    expect(grid.renderer.getRowHeight(1)).toBe(16 * zoom);

    grid.destroy();
    container.remove();
  });
});

