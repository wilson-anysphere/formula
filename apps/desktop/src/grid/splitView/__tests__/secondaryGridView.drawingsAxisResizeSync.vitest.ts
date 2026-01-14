/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { SecondaryGridView } from "../secondaryGridView";
import { DocumentController } from "../../../document/documentController.js";
import type { DrawingObject, ImageStore } from "../../../drawings/types";
import { pxToEmu } from "../../../drawings/overlay";

type CtxCall = { method: string; args: unknown[] };

function createMockCanvasContext(): CanvasRenderingContext2D {
  const noop = () => {};
  const gradient = { addColorStop: noop } as any;
  const context = new Proxy(
    {
      canvas: document.createElement("canvas"),
      measureText: (text: string) => ({ width: text.length * 8 }),
      createLinearGradient: () => gradient,
      createPattern: () => null,
      getImageData: () => ({ data: new Uint8ClampedArray(), width: 0, height: 0 }),
      putImageData: noop,
    },
    {
      get(target, prop) {
        if (prop in target) return (target as any)[prop];
        return noop;
      },
      set(target, prop, value) {
        (target as any)[prop] = value;
        return true;
      },
    },
  );
  return context as any;
}

function createRecordingCanvasContext(calls: CtxCall[]): CanvasRenderingContext2D {
  const noop = () => {};
  const gradient = { addColorStop: noop } as any;
  const context = new Proxy(
    {
      canvas: document.createElement("canvas"),
      measureText: (text: string) => ({ width: text.length * 8 }),
      createLinearGradient: () => gradient,
      createPattern: () => null,
      getImageData: () => ({ data: new Uint8ClampedArray(), width: 0, height: 0 }),
      putImageData: noop,
    },
    {
      get(target, prop) {
        if (prop in target) return (target as any)[prop];
        if (typeof prop === "string") {
          return (...args: unknown[]) => {
            calls.push({ method: prop, args });
          };
        }
        return noop;
      },
      set(target, prop, value) {
        (target as any)[prop] = value;
        return true;
      },
    },
  );
  return context as any;
}

describe("SecondaryGridView drawings overlay + axis resize", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  beforeEach(() => {
    document.body.innerHTML = "";

    // Prevent DesktopSharedGrid from delivering rAF-batched scroll callbacks during construction.
    Object.defineProperty(globalThis, "requestAnimationFrame", {
      configurable: true,
      value: () => 0,
    });
    Object.defineProperty(globalThis, "cancelAnimationFrame", { configurable: true, value: () => {} });

    (globalThis as any).ResizeObserver = class {
      observe() {}
      disconnect() {}
    };
  });

  it("re-renders drawings when column widths change", async () => {
    const container = document.createElement("div");
    Object.defineProperty(container, "clientWidth", { configurable: true, value: 300 });
    Object.defineProperty(container, "clientHeight", { configurable: true, value: 200 });
    document.body.appendChild(container);

    const calls: CtxCall[] = [];
    const drawingsCtx = createRecordingCanvasContext(calls);

    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: function (this: HTMLCanvasElement) {
        if (this.classList.contains("grid-canvas--drawings")) return drawingsCtx;
        return createMockCanvasContext();
      },
    });

    const doc = new DocumentController();
    const sheetId = "Sheet1";

    const images: ImageStore = { get: () => undefined, set: () => {} };

    // Anchor at B1 so its x-position depends on the width of column A.
    const objects: DrawingObject[] = [
      {
        id: 1,
        kind: { type: "image", imageId: "missing" },
        zOrder: 0,
        anchor: {
          type: "oneCell",
          from: { cell: { row: 0, col: 1 }, offset: { xEmu: 0, yEmu: 0 } }, // B1
          size: { cx: pxToEmu(10), cy: pxToEmu(10) },
        },
      },
    ];

    const gridView = new SecondaryGridView({
      container,
      document: doc,
      getSheetId: () => sheetId,
      rowCount: 20,
      colCount: 20,
      showFormulas: () => false,
      getComputedValue: () => null,
      getDrawingObjects: () => objects,
      images,
    });

    // Flush any initial render work from construction.
    calls.splice(0, calls.length);
    await (gridView as any).renderDrawings();

    const firstStroke = calls.find((call) => call.method === "strokeRect");
    expect(firstStroke).toBeTruthy();
    const x1 = Number(firstStroke!.args[0]);
    expect(Number.isFinite(x1)).toBe(true);

    const renderer = gridView.grid.renderer;
    const index = 1; // grid col 1 => doc col 0 (A)
    const prevSize = renderer.getColWidth(index);
    const nextSize = prevSize + 50;
    // Mimic interactive drag: renderer is already updated before the callback fires.
    renderer.setColWidth(index, nextSize);

    calls.splice(0, calls.length);
    const renderSpy = vi.spyOn(gridView as any, "renderDrawings");
    renderSpy.mockClear();

    (gridView as any).onAxisSizeChange({
      kind: "col",
      index,
      size: nextSize,
      previousSize: prevSize,
      defaultSize: renderer.scroll.cols.defaultSize,
      zoom: renderer.getZoom(),
      source: "resize",
    });

    expect(renderSpy).toHaveBeenCalled();
    // Wait for the async overlay render triggered by the axis-size callback.
    await ((gridView as any).drawingsRenderPromise ?? Promise.resolve());

    const secondStroke = calls.find((call) => call.method === "strokeRect");
    expect(secondStroke).toBeTruthy();
    const x2 = Number(secondStroke!.args[0]);
    expect(x2).toBeCloseTo(x1 + (nextSize - prevSize), 6);

    gridView.destroy();
    container.remove();
  });

  it("re-renders drawings when row heights change", async () => {
    const container = document.createElement("div");
    Object.defineProperty(container, "clientWidth", { configurable: true, value: 300 });
    Object.defineProperty(container, "clientHeight", { configurable: true, value: 200 });
    document.body.appendChild(container);

    const calls: CtxCall[] = [];
    const drawingsCtx = createRecordingCanvasContext(calls);

    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: function (this: HTMLCanvasElement) {
        if (this.classList.contains("grid-canvas--drawings")) return drawingsCtx;
        return createMockCanvasContext();
      },
    });

    const doc = new DocumentController();
    const sheetId = "Sheet1";

    const images: ImageStore = { get: () => undefined, set: () => {} };

    // Anchor at A2 so its y-position depends on the height of row 1 (doc row 0).
    const objects: DrawingObject[] = [
      {
        id: 1,
        kind: { type: "image", imageId: "missing" },
        zOrder: 0,
        anchor: {
          type: "oneCell",
          from: { cell: { row: 1, col: 0 }, offset: { xEmu: 0, yEmu: 0 } }, // A2
          size: { cx: pxToEmu(10), cy: pxToEmu(10) },
        },
      },
    ];

    const gridView = new SecondaryGridView({
      container,
      document: doc,
      getSheetId: () => sheetId,
      rowCount: 20,
      colCount: 20,
      showFormulas: () => false,
      getComputedValue: () => null,
      getDrawingObjects: () => objects,
      images,
    });

    calls.splice(0, calls.length);
    await (gridView as any).renderDrawings();

    const firstStroke = calls.find((call) => call.method === "strokeRect");
    expect(firstStroke).toBeTruthy();
    const y1 = Number(firstStroke!.args[1]);
    expect(Number.isFinite(y1)).toBe(true);

    const renderer = gridView.grid.renderer;
    const index = 1; // grid row 1 => doc row 0 (row 1)
    const prevSize = renderer.getRowHeight(index);
    const nextSize = prevSize + 30;
    renderer.setRowHeight(index, nextSize);

    calls.splice(0, calls.length);
    const renderSpy = vi.spyOn(gridView as any, "renderDrawings");
    renderSpy.mockClear();

    (gridView as any).onAxisSizeChange({
      kind: "row",
      index,
      size: nextSize,
      previousSize: prevSize,
      defaultSize: renderer.scroll.rows.defaultSize,
      zoom: renderer.getZoom(),
      source: "resize",
    });

    expect(renderSpy).toHaveBeenCalled();
    await ((gridView as any).drawingsRenderPromise ?? Promise.resolve());

    const secondStroke = calls.find((call) => call.method === "strokeRect");
    expect(secondStroke).toBeTruthy();
    const y2 = Number(secondStroke!.args[1]);
    expect(y2).toBeCloseTo(y1 + (nextSize - prevSize), 6);

    gridView.destroy();
    container.remove();
  });
});
