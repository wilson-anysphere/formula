// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { CellProvider } from "../../model/CellProvider";
import { CanvasGridRenderer } from "../CanvasGridRenderer";

function createMock2dContext(canvas: HTMLCanvasElement): CanvasRenderingContext2D {
  const noop = () => {};

  // Keep in sync with other CanvasGridRenderer unit tests: a minimal no-op 2D context surface.
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

describe("CanvasGridRenderer.subscribeViewport", () => {
  const originalGetContext = HTMLCanvasElement.prototype.getContext;

  let rafCallbacks: Map<number, FrameRequestCallback>;
  let nextRafId = 1;

  const flushRaf = () => {
    const callbacks = Array.from(rafCallbacks.values());
    rafCallbacks.clear();
    for (const cb of callbacks) cb(0);
  };

  beforeEach(() => {
    rafCallbacks = new Map();
    nextRafId = 1;

    vi.stubGlobal("requestAnimationFrame", (cb: FrameRequestCallback) => {
      const id = nextRafId++;
      rafCallbacks.set(id, cb);
      return id;
    });
    vi.stubGlobal("cancelAnimationFrame", (id: number) => {
      rafCallbacks.delete(id);
    });

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      return createMock2dContext(this);
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;
  });

  afterEach(() => {
    HTMLCanvasElement.prototype.getContext = originalGetContext;
    vi.useRealTimers();
    vi.unstubAllGlobals();
  });

  function createRenderer(): CanvasGridRenderer {
    const provider: CellProvider = { getCell: () => null };
    const renderer = new CanvasGridRenderer({ provider, rowCount: 100, colCount: 100, defaultRowHeight: 10, defaultColWidth: 10 });
    // The unit tests only care about notification behavior; skip heavy rendering work.
    (renderer as unknown as { renderFrame: () => void }).renderFrame = vi.fn();

    const grid = document.createElement("canvas");
    const content = document.createElement("canvas");
    const selection = document.createElement("canvas");
    renderer.attach({ grid, content, selection });
    renderer.resize(200, 100, 1);
    flushRaf();
    return renderer;
  }

  it("fires for axis size changes but not for scroll offset updates", () => {
    const renderer = createRenderer();

    const listener = vi.fn();
    renderer.subscribeViewport(listener, { animationFrame: true });

    renderer.scrollBy(10, 10);
    flushRaf();
    expect(listener).not.toHaveBeenCalled();

    renderer.setRowHeight(0, 20);
    flushRaf();
    expect(listener).toHaveBeenCalledTimes(1);
    expect(listener.mock.calls[0]?.[0]?.reason).toBe("axisSize");
  });

  it("coalesces multiple changes into a single callback when animationFrame is enabled", () => {
    const renderer = createRenderer();

    const listener = vi.fn();
    renderer.subscribeViewport(listener, { animationFrame: true });

    renderer.setRowHeight(0, 20);
    renderer.setRowHeight(1, 25);
    expect(listener).not.toHaveBeenCalled();

    flushRaf();
    expect(listener).toHaveBeenCalledTimes(1);
  });

  it("debounces notifications when debounceMs is set", () => {
    vi.useFakeTimers();

    const renderer = createRenderer();

    const listener = vi.fn();
    renderer.subscribeViewport(listener, { debounceMs: 25 });

    renderer.setColWidth(0, 150);
    renderer.setColWidth(1, 160);
    expect(listener).not.toHaveBeenCalled();

    vi.advanceTimersByTime(24);
    expect(listener).not.toHaveBeenCalled();

    vi.advanceTimersByTime(1);
    expect(listener).toHaveBeenCalledTimes(1);
    expect(listener.mock.calls[0]?.[0]?.reason).toBe("axisSize");
  });

  it("does not fire a pending animationFrame notification after destroy()", () => {
    const renderer = createRenderer();

    const listener = vi.fn();
    renderer.subscribeViewport(listener, { animationFrame: true });

    renderer.setRowHeight(0, 20);
    renderer.destroy();
    flushRaf();

    expect(listener).not.toHaveBeenCalled();
  });
});

