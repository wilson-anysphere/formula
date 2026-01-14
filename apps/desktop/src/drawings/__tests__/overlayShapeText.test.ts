import { describe, expect, it } from "vitest";

import { DrawingOverlay, pxToEmu, type GridGeometry, type Viewport } from "../overlay";
import type { DrawingObject, ImageStore } from "../types";

function createStubCanvasContext(): { ctx: CanvasRenderingContext2D; calls: Array<{ method: string; args: unknown[] }> } {
  const calls: Array<{ method: string; args: unknown[] }> = [];
  const ctx: any = {
    clearRect: (...args: unknown[]) => calls.push({ method: "clearRect", args }),
    drawImage: (...args: unknown[]) => calls.push({ method: "drawImage", args }),
    save: () => calls.push({ method: "save", args: [] }),
    restore: () => calls.push({ method: "restore", args: [] }),
    beginPath: () => calls.push({ method: "beginPath", args: [] }),
    rect: (...args: unknown[]) => calls.push({ method: "rect", args }),
    clip: () => calls.push({ method: "clip", args: [] }),
    setLineDash: (...args: unknown[]) => calls.push({ method: "setLineDash", args }),
    strokeRect: (...args: unknown[]) => calls.push({ method: "strokeRect", args }),
    fillText: (...args: unknown[]) => calls.push({ method: "fillText", args }),
    measureText: (text: string) => ({ width: text.length * 6 }),
    moveTo: (...args: unknown[]) => calls.push({ method: "moveTo", args }),
    lineTo: (...args: unknown[]) => calls.push({ method: "lineTo", args }),
    stroke: (...args: unknown[]) => calls.push({ method: "stroke", args }),
  };

  return { ctx: ctx as CanvasRenderingContext2D, calls };
}

function createStubCanvas(ctx: CanvasRenderingContext2D): HTMLCanvasElement {
  const canvas: any = {
    width: 0,
    height: 0,
    style: {},
    getContext: (type: string) => (type === "2d" ? ctx : null),
  };
  return canvas as HTMLCanvasElement;
}

function createShapeObject(rawXml: string): DrawingObject {
  return {
    id: 1,
    kind: { type: "shape", rawXml },
    anchor: {
      type: "absolute",
      pos: { xEmu: pxToEmu(5), yEmu: pxToEmu(7) },
      size: { cx: pxToEmu(120), cy: pxToEmu(30) },
    },
    zOrder: 0,
  };
}

const images: ImageStore = {
  get: () => undefined,
  set: () => {},
  delete: () => {},
  clear: () => {},
};

const geom: GridGeometry = {
  cellOriginPx: () => ({ x: 0, y: 0 }),
  cellSizePx: () => ({ width: 0, height: 0 }),
};

const viewport: Viewport = { scrollX: 0, scrollY: 0, width: 200, height: 200, dpr: 1 };

describe("DrawingOverlay shapes", () => {
  it("renders txBody text instead of placeholder label", async () => {
    const { ctx, calls } = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);

    const overlay = new DrawingOverlay(canvas, images, geom);

    const rawXml = `
      <xdr:sp>
        <xdr:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r><a:t>Hello</a:t></a:r>
          </a:p>
        </xdr:txBody>
      </xdr:sp>
    `;

    await overlay.render([createShapeObject(rawXml)], viewport);

    const fillTextValues = calls.filter((call) => call.method === "fillText").map((call) => call.args[0]);
    expect(fillTextValues).toContain("Hello");
    expect(fillTextValues).not.toContain("shape");
  });
});
