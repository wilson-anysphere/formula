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
    ellipse: (...args: unknown[]) => calls.push({ method: "ellipse", args }),
    moveTo: (...args: unknown[]) => calls.push({ method: "moveTo", args }),
    lineTo: (...args: unknown[]) => calls.push({ method: "lineTo", args }),
    arcTo: (...args: unknown[]) => calls.push({ method: "arcTo", args }),
    closePath: (...args: unknown[]) => calls.push({ method: "closePath", args }),
    fill: (...args: unknown[]) => calls.push({ method: "fill", args }),
    stroke: (...args: unknown[]) => calls.push({ method: "stroke", args }),
    // Used by shape txBody text layout.
    measureText: (text: string) => ({ width: text.length * 8 }),
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

function createShapeObject(raw_xml: string, opts?: { widthPx?: number; heightPx?: number }): DrawingObject {
  const widthPx = opts?.widthPx ?? 20;
  const heightPx = opts?.heightPx ?? 10;
  return {
    id: 1,
    kind: { type: "shape", raw_xml },
    anchor: {
      type: "absolute",
      pos: { xEmu: pxToEmu(5), yEmu: pxToEmu(7) },
      size: { cx: pxToEmu(widthPx), cy: pxToEmu(heightPx) },
    },
    zOrder: 0,
  };
}

const images: ImageStore = {
  get: () => undefined,
  set: () => {},
};

const geom: GridGeometry = {
  cellOriginPx: () => ({ x: 0, y: 0 }),
  cellSizePx: () => ({ width: 0, height: 0 }),
};

const viewport: Viewport = { scrollX: 0, scrollY: 0, width: 100, height: 100, dpr: 1 };

describe("DrawingOverlay shapes", () => {
  it("renders supported shapes using canvas primitives (no placeholder)", async () => {
    const { ctx, calls } = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);

    const xml = `
      <xdr:sp>
        <xdr:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p><a:r><a:t>Hello</a:t></a:r></a:p>
        </xdr:txBody>
        <xdr:spPr>
          <a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
          <a:ln w="12700">
            <a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>
          </a:ln>
        </xdr:spPr>
      </xdr:sp>
    `;

    const overlay = new DrawingOverlay(canvas, images, geom);
    await overlay.render([createShapeObject(xml)], viewport);

    expect(calls.some((call) => call.method === "ellipse")).toBe(true);
    expect(calls.some((call) => call.method === "fill")).toBe(true);
    expect(calls.some((call) => call.method === "stroke")).toBe(true);
    expect(calls.some((call) => call.method === "fillText" && call.args[0] === "Hello")).toBe(true);
    expect(calls.some((call) => call.method === "strokeRect")).toBe(false);
  });

  it("renders line shapes using moveTo/lineTo", async () => {
    const { ctx, calls } = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);

    const xml = `
      <xdr:sp>
        <xdr:spPr>
          <a:prstGeom prst="line"><a:avLst/></a:prstGeom>
          <a:ln w="12700">
            <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
          </a:ln>
        </xdr:spPr>
      </xdr:sp>
    `;

    const overlay = new DrawingOverlay(canvas, images, geom);
    await overlay.render([createShapeObject(xml)], viewport);

    expect(calls.some((call) => call.method === "moveTo")).toBe(true);
    expect(calls.some((call) => call.method === "lineTo")).toBe(true);
    expect(calls.some((call) => call.method === "stroke")).toBe(true);
    expect(calls.some((call) => call.method === "strokeRect")).toBe(false);
  });

  it("falls back to placeholder rendering for unsupported shapes", async () => {
    const { ctx, calls } = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);

    const xml = `
      <xdr:spPr>
        <a:prstGeom prst="triangle"><a:avLst/></a:prstGeom>
      </xdr:spPr>
    `;

    const overlay = new DrawingOverlay(canvas, images, geom);
    await overlay.render([createShapeObject(xml)], viewport);

    expect(calls.some((call) => call.method === "strokeRect")).toBe(true);
  });

  it("applies dash patterns for stroked shapes when prstDash is present", async () => {
    const { ctx, calls } = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);

    const xml = `
      <xdr:sp>
        <xdr:spPr>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:ln w="12700">
            <a:prstDash val="dash"/>
          </a:ln>
        </xdr:spPr>
      </xdr:sp>
    `;

    const overlay = new DrawingOverlay(canvas, images, geom);
    await overlay.render([createShapeObject(xml)], viewport);

    expect(
      calls.some(
        (call) => call.method === "setLineDash" && Array.isArray(call.args[0]) && (call.args[0] as any[]).length > 0,
      ),
    ).toBe(true);
    expect(calls.some((call) => call.method === "strokeRect")).toBe(false);
  });

  it("positions label text using txBody alignment (center/middle)", async () => {
    const { ctx, calls } = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);

    const xml = `
      <xdr:sp>
        <xdr:txBody>
          <a:bodyPr anchor="ctr"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr algn="ctr">
              <a:defRPr sz="1200"/>
            </a:pPr>
            <a:r><a:t>Centered</a:t></a:r>
          </a:p>
        </xdr:txBody>
        <xdr:spPr>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </xdr:spPr>
      </xdr:sp>
    `;

    const overlay = new DrawingOverlay(canvas, images, geom);
    // Ensure the text box has enough space (after the 4px padding on each side) for the
    // label to be centered without being clamped to the padding bounds.
    await overlay.render([createShapeObject(xml, { widthPx: 100, heightPx: 40 })], viewport);

    const call = calls.find((c) => c.method === "fillText" && c.args[0] === "Centered");
    expect(call).toBeTruthy();
    // `renderShapeText` uses a 4px padding inset and positions the text at the top-left
    // baseline with center/middle alignment within that padded region.
    expect(Number(call!.args[1])).toBeCloseTo(23, 5);
    expect(Number(call!.args[2])).toBeCloseTo(17.4, 5);
  });
});
