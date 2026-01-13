// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { CellProvider } from "../../model/CellProvider";
import { CanvasGridRenderer } from "../CanvasGridRenderer";

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
    setLineDash: noop,
    measureText: (text: string) =>
      ({
        width: text.length * 6,
        actualBoundingBoxAscent: 8,
        actualBoundingBoxDescent: 2
      }) as TextMetrics
  } as unknown as CanvasRenderingContext2D;
}

describe("CanvasGridRenderer.setScroll", () => {
  const originalGetContext = HTMLCanvasElement.prototype.getContext;
  const originalRaf = globalThis.requestAnimationFrame;

  let rafSpy: ReturnType<typeof vi.fn>;

  beforeEach(() => {
    rafSpy = vi.fn((cb: FrameRequestCallback) => {
      cb(0);
      return 0;
    });
    vi.stubGlobal("requestAnimationFrame", rafSpy);

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      return createMock2dContext(this);
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;
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

  it("is a no-op when the aligned/clamped scroll position does not change", () => {
    const prefetchSpy = vi.fn();
    const provider: CellProvider = { getCell: () => null, prefetch: prefetchSpy };

    const renderer = new CanvasGridRenderer({ provider, rowCount: 100, colCount: 10 });
    // The unit test only cares about invalidation side effects; skip heavy rendering.
    (renderer as unknown as { renderFrame: () => void }).renderFrame = vi.fn();

    const grid = document.createElement("canvas");
    const content = document.createElement("canvas");
    const selection = document.createElement("canvas");

    renderer.attach({ grid, content, selection });
    renderer.resize(200, 100, 1);

    // Clear initial prefetch + render invalidation triggered by `resize()`.
    prefetchSpy.mockClear();
    rafSpy.mockClear();

    renderer.setScroll(0, 0);

    expect(prefetchSpy).not.toHaveBeenCalled();
    expect(rafSpy).not.toHaveBeenCalled();

    renderer.setScroll(0, 50);

    expect(prefetchSpy).toHaveBeenCalled();
    expect(rafSpy).toHaveBeenCalled();
  });
});

