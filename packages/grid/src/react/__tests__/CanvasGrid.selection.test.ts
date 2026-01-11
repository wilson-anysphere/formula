// @vitest-environment jsdom
import React from "react";
import { act } from "react-dom/test-utils";
import { createRoot } from "react-dom/client";
import { beforeEach, afterEach, describe, expect, it, vi } from "vitest";
import { CanvasGrid, MockCellProvider, type GridApi } from "../../index";

function createMock2dContext(canvas: HTMLCanvasElement): CanvasRenderingContext2D {
  const noop = () => {};

  // CanvasGridRenderer uses a fairly wide surface area of the 2D context for
  // rendering. The selection API tests only need attach/resize to succeed, so a
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

describe("CanvasGrid selection API", () => {
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
  });

  afterEach(() => {
    HTMLCanvasElement.prototype.getContext = originalGetContext;
    vi.unstubAllGlobals();
  });

  it("notifies onSelectionChange and exposes selection via apiRef", async () => {
    const provider = new MockCellProvider({ rowCount: 10, colCount: 10 });
    const apiRef = React.createRef<GridApi>();
    const onSelectionChange = vi.fn();

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        React.createElement(CanvasGrid, {
          provider,
          rowCount: 10,
          colCount: 10,
          apiRef,
          onSelectionChange
        })
      );
    });

    const canvases = host.querySelectorAll("canvas");
    expect(canvases.length).toBe(3);
    const selectionCanvas = canvases[2] as HTMLCanvasElement;

    vi.spyOn(selectionCanvas, "getBoundingClientRect").mockReturnValue({
      left: 0,
      top: 0,
      right: 0,
      bottom: 0,
      width: 0,
      height: 0,
      x: 0,
      y: 0,
      toJSON: () => {}
    } as DOMRect);

    const eventInit = { clientX: 5, clientY: 5, bubbles: true };
    const event =
      typeof PointerEvent !== "undefined"
        ? new PointerEvent("pointerdown", eventInit)
        : new MouseEvent("pointerdown", eventInit);

    await act(async () => {
      selectionCanvas.dispatchEvent(event);
    });

    expect(onSelectionChange).toHaveBeenCalledTimes(1);
    expect(onSelectionChange).toHaveBeenCalledWith({ row: 0, col: 0 });
    expect(apiRef.current?.getSelection()).toEqual({ row: 0, col: 0 });

    apiRef.current?.clearSelection();
    expect(onSelectionChange).toHaveBeenCalledTimes(2);
    expect(onSelectionChange).toHaveBeenLastCalledWith(null);
    expect(apiRef.current?.getSelection()).toBeNull();

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });
});

