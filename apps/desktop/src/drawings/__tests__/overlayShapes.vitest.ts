import { describe, expect, it } from "vitest";

import { DrawingOverlay, pxToEmu, type GridGeometry, type Viewport } from "../overlay";
import type { DrawingObject, ImageStore } from "../types";

function createStubCanvasContext(): { ctx: CanvasRenderingContext2D; calls: Array<{ method: string; args: unknown[] }> } {
  const calls: Array<{ method: string; args: unknown[] }> = [];
  const ctx: any = {
    clearRect: (...args: unknown[]) => calls.push({ method: "clearRect", args }),
    drawImage: (...args: unknown[]) => calls.push({ method: "drawImage", args }),
    // Used by shape txBody text layout.
    measureText: (text: string) => ({ width: String(text).length * 8 }),
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

function createShapeObject(raw_xml: string, size: { width: number; height: number } = { width: 20, height: 10 }): DrawingObject {
  return {
    id: 1,
    kind: { type: "shape", raw_xml },
    anchor: {
      type: "absolute",
      pos: { xEmu: pxToEmu(5), yEmu: pxToEmu(7) },
      size: { cx: pxToEmu(size.width), cy: pxToEmu(size.height) },
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

const viewport: Viewport = { scrollX: 0, scrollY: 0, width: 100, height: 100, dpr: 1, zoom: 1 };

describe("DrawingOverlay shapes", () => {
  it("renders supported shapes using canvas primitives (no placeholder)", () => {
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
    overlay.render([createShapeObject(xml)], viewport);

    expect(calls.some((call) => call.method === "ellipse")).toBe(true);
    expect(calls.some((call) => call.method === "fill")).toBe(true);
    expect(calls.some((call) => call.method === "stroke")).toBe(true);
    expect(calls.some((call) => call.method === "fillText" && call.args[0] === "Hello")).toBe(true);
    expect(calls.some((call) => call.method === "strokeRect")).toBe(false);
  });

  it("renders <a:tab/> placeholders as spaces on the canvas", () => {
    const { ctx, calls } = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);

    const xml = `
      <xdr:sp>
        <xdr:txBody>
          <a:bodyPr wrap="none"/>
          <a:lstStyle/>
          <a:p>
            <a:r><a:t>Hello</a:t></a:r>
            <a:tab/>
            <a:r><a:t>World</a:t></a:r>
          </a:p>
        </xdr:txBody>
      </xdr:sp>
    `;

    const overlay = new DrawingOverlay(canvas, images, geom);
    overlay.render([createShapeObject(xml, { width: 200, height: 40 })], viewport);

    expect(calls.some((call) => call.method === "fillText" && call.args[0] === "Hello    World")).toBe(true);
  });

  it("preserves leading spaces in txBody text runs (xml:space=\"preserve\")", () => {
    const { ctx, calls } = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);

    const xml = `
      <xdr:sp>
        <xdr:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r><a:t xml:space="preserve">  Indented</a:t></a:r>
          </a:p>
        </xdr:txBody>
      </xdr:sp>
    `;

    const overlay = new DrawingOverlay(canvas, images, geom);
    overlay.render([createShapeObject(xml, { widthPx: 200, heightPx: 40 })], viewport);

    expect(calls.some((call) => call.method === "fillText" && call.args[0] === "  Indented")).toBe(true);
  });

  it("renders line shapes using moveTo/lineTo", () => {
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
    overlay.render([createShapeObject(xml)], viewport);

    expect(calls.some((call) => call.method === "moveTo")).toBe(true);
    expect(calls.some((call) => call.method === "lineTo")).toBe(true);
    expect(calls.some((call) => call.method === "stroke")).toBe(true);
    expect(calls.some((call) => call.method === "strokeRect")).toBe(false);
  });

  it("falls back to placeholder rendering for unsupported shapes", () => {
    const { ctx, calls } = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);

    const xml = `
      <xdr:spPr>
        <a:prstGeom prst="triangle"><a:avLst/></a:prstGeom>
      </xdr:spPr>
    `;

    const overlay = new DrawingOverlay(canvas, images, geom);
    overlay.render([createShapeObject(xml)], viewport);

    expect(calls.some((call) => call.method === "strokeRect")).toBe(true);
  });

  it("applies dash patterns for stroked shapes when prstDash is present", () => {
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
    overlay.render([createShapeObject(xml)], viewport);

    expect(
      calls.some(
        (call) => call.method === "setLineDash" && Array.isArray(call.args[0]) && (call.args[0] as any[]).length > 0,
      ),
    ).toBe(true);
    expect(calls.some((call) => call.method === "strokeRect")).toBe(false);
  });

  it("positions label text using txBody alignment (center/middle)", () => {
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
    const bounds = { x: 5, y: 7, width: 60, height: 40 };
    overlay.render([createShapeObject(xml, { width: bounds.width, height: bounds.height })], viewport);

    const textCall = calls.find((call) => call.method === "fillText" && call.args[0] === "Centered");
    expect(textCall).toBeTruthy();
    const x = textCall!.args[1] as number;
    const y = textCall!.args[2] as number;

    // Center alignment is computed manually by the overlay (textAlign remains "left") so
    // we expect `fillText` at the top-left origin of the centered block.
    const expectedTextWidth = (ctx as any).measureText("Centered").width as number;
    const expectedTextHeight = (12 * 96) / 72 * 1.2; // `sz="1200"` => 12pt, lineHeight = fontPx * 1.2
    const expectedX = bounds.x + bounds.width / 2 - expectedTextWidth / 2;
    const expectedY = bounds.y + bounds.height / 2 - expectedTextHeight / 2;

    expect(x).toBeCloseTo(expectedX, 6);
    expect(y).toBeCloseTo(expectedY, 6);
  });
});
