// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { CellProvider, CellProviderUpdate, CellRange } from "@formula/grid";
import { DesktopSharedGrid } from "../desktopSharedGrid";

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

class SubscribableCellProvider implements CellProvider {
  private readonly listeners = new Set<(update: CellProviderUpdate) => void>();
  private value: string;

  constructor(value: string) {
    this.value = value;
  }

  setValue(next: string): void {
    this.value = next;
  }

  emit(update: CellProviderUpdate): void {
    for (const listener of this.listeners) listener(update);
  }

  getCell(row: number, col: number) {
    if (row === 0 && col === 0) return { row, col, value: this.value };
    return { row, col, value: null };
  }

  subscribe(listener: (update: CellProviderUpdate) => void): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }
}

describe("DesktopSharedGrid a11y provider updates", () => {
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
  });

  afterEach(() => {
    HTMLCanvasElement.prototype.getContext = originalGetContext;
    vi.unstubAllGlobals();
    document.body.innerHTML = "";
  });

  it("refreshes the a11y status text when the active cell value changes", () => {
    const provider = new SubscribableCellProvider("hello");

    const container = document.createElement("div");
    document.body.appendChild(container);

    const canvases = {
      grid: document.createElement("canvas"),
      content: document.createElement("canvas"),
      selection: document.createElement("canvas")
    };
    container.appendChild(canvases.grid);
    container.appendChild(canvases.content);
    container.appendChild(canvases.selection);

    const scrollbars = {
      vTrack: document.createElement("div"),
      vThumb: document.createElement("div"),
      hTrack: document.createElement("div"),
      hThumb: document.createElement("div")
    };
    scrollbars.vTrack.appendChild(scrollbars.vThumb);
    scrollbars.hTrack.appendChild(scrollbars.hThumb);
    container.appendChild(scrollbars.vTrack);
    container.appendChild(scrollbars.hTrack);

    const grid = new DesktopSharedGrid({
      container,
      provider,
      rowCount: 10,
      colCount: 10,
      canvases,
      scrollbars,
      enableWheel: false,
      enableKeyboard: false,
      enableResize: false
    });

    grid.resize(300, 200, 1);

    grid.setSelectionRanges([{ startRow: 0, endRow: 1, startCol: 0, endCol: 1 }], { activeCell: { row: 0, col: 0 } });

    const status = container.querySelector('[data-testid="canvas-grid-a11y-status"]') as HTMLDivElement | null;
    expect(status?.textContent).toContain("Active cell A1, value hello.");

    const activeCell = container.querySelector('[data-testid="canvas-grid-a11y-active-cell"]') as HTMLDivElement | null;
    expect(activeCell?.textContent).toContain("Cell A1, value hello.");

    provider.setValue("world");
    const updateRange: CellRange = { startRow: 0, endRow: 1, startCol: 0, endCol: 1 };
    provider.emit({ type: "cells", range: updateRange });

    expect(status?.textContent).toContain("Active cell A1, value world.");
    expect(activeCell?.textContent).toContain("Cell A1, value world.");

    // Updates outside the selected cell should not force an announcement.
    provider.setValue("ignored");
    provider.emit({ type: "cells", range: { startRow: 5, endRow: 6, startCol: 5, endCol: 6 } });
    expect(status?.textContent).toContain("Active cell A1, value world.");

    grid.destroy();
  });
});

