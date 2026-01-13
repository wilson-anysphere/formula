import { describe, expect, it } from "vitest";

import { patchNvPrId, patchXfrmExt, patchXfrmOff } from "../drawingml/patch";
import { duplicateDrawingObject } from "../duplicate";
import { DrawingInteractionController } from "../interaction";
import { pxToEmu } from "../overlay";
import type { DrawingObject } from "../types";
import type { GridGeometry, Viewport } from "../overlay";

describe("DrawingML patch helpers", () => {
  it("patchNvPrId updates cNvPr/@id without assuming prefixes", () => {
    const input =
      `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Picture 1"/>` +
      `<xdr:cNvPicPr/></xdr:nvPicPr><a:blip r:embed="rId1"/></xdr:pic>`;
    const out = patchNvPrId(input, 42);
    expect(out).toContain(`cNvPr id="42"`);
    expect(out).toContain(`r:embed="rId1"`);
  });

  it("patchNvPrId patches cNvPr id inside graphicFrame payloads", () => {
    const input = `<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr></xdr:graphicFrame>`;
    const out = patchNvPrId(input, 99);
    expect(out).toContain(`cNvPr id="99"`);
    expect(out).toContain(`name="Chart 1"`);
  });

  it("patchXfrmExt updates ext under xfrm (and leaves other ext nodes alone)", () => {
    const input =
      `<xdr:spPr><xdr:ext cx="1" cy="2"/>` +
      `<a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="200"/></a:xfrm></xdr:spPr>`;
    const out = patchXfrmExt(input, 300, 400);
    expect(out).toContain(`<xdr:ext cx="1" cy="2"/>`);
    expect(out).toContain(`<a:ext cx="300" cy="400"/>`);
  });

  it("patchXfrmExt patches ext inside graphicFrame xfrm", () => {
    const input = `<xdr:graphicFrame><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm></xdr:graphicFrame>`;
    const out = patchXfrmExt(input, 123, 456);
    expect(out).toContain(`<a:ext cx="123" cy="456"/>`);
  });

  it("patchXfrmOff updates off under xfrm", () => {
    const input = `<a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm>`;
    const out = patchXfrmOff(input, 10, 20);
    expect(out).toContain(`<a:off x="10" y="20"/>`);
  });

  it("patchXfrmOff does not touch namespaced attributes (guardrail)", () => {
    const input = `<a:xfrm><a:off r:x="99" r:y="88"/><a:ext cx="3" cy="4"/></a:xfrm>`;
    const out = patchXfrmOff(input, 10, 20);
    // No-op because there is no plain `x=`/`y=` attribute to patch.
    expect(out).toBe(input);
  });
});

describe("duplicateDrawingObject", () => {
  it("assigns a new id and patches preserved pic xml (id + best-effort name)", () => {
    const obj: DrawingObject = {
      id: 1,
      kind: { type: "image", imageId: "img1.png" },
      anchor: {
        type: "absolute",
        pos: { xEmu: 0, yEmu: 0 },
        size: { cx: pxToEmu(10), cy: pxToEmu(10) },
      },
      zOrder: 0,
      preserved: {
        "xlsx.pic_xml": `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="Picture 1"/></xdr:nvPicPr></xdr:pic>`,
      },
    };

    const dup = duplicateDrawingObject(obj, 7);
    expect(dup).not.toBe(obj);
    expect(dup.id).toBe(7);
    expect(dup.preserved?.["xlsx.pic_xml"]).toContain(`id="7"`);
    expect(dup.preserved?.["xlsx.pic_xml"]).toContain(`name="Picture 7"`);
    // Ensure original is unchanged (immutability).
    expect(obj.preserved?.["xlsx.pic_xml"]).toContain(`id="1"`);
  });

  it("patches rawXml/raw_xml payloads for shapes and charts", () => {
    const shape: DrawingObject = {
      id: 2,
      kind: {
        type: "shape",
        rawXml: `<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="2" name="Shape 2"/></xdr:nvSpPr></xdr:sp>`,
        raw_xml: `<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="2" name="Shape 2"/></xdr:nvSpPr></xdr:sp>`,
      },
      anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: 0, cy: 0 } },
      zOrder: 0,
    };

    const dupShape = duplicateDrawingObject(shape, 8);
    expect((dupShape.kind as any).rawXml).toContain(`id="8"`);
    expect((dupShape.kind as any).rawXml).toContain(`name="Shape 8"`);
    expect((dupShape.kind as any).raw_xml).toContain(`id="8"`);

    const chart: DrawingObject = {
      id: 3,
      kind: {
        type: "chart",
        chartId: "rId1",
        rawXml: `<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="3" name="Chart 3"/></xdr:nvGraphicFramePr></xdr:graphicFrame>`,
      },
      anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: 0, cy: 0 } },
      zOrder: 0,
    };

    const dupChart = duplicateDrawingObject(chart, 9);
    expect((dupChart.kind as any).rawXml).toContain(`id="9"`);
    expect((dupChart.kind as any).rawXml).toContain(`name="Chart 9"`);
  });
});

describe("DrawingInteractionController commit-time patching", () => {
  function createStubCanvas() {
    const listeners = new Map<string, (e: any) => void>();
    const canvas: any = {
      addEventListener: (type: string, cb: (e: any) => void) => listeners.set(type, cb),
      removeEventListener: (type: string) => listeners.delete(type),
      setPointerCapture: () => {},
      releasePointerCapture: () => {},
      style: { cursor: "default" },
      dispatch: (type: string, event: any) => listeners.get(type)?.(event),
    };
    return canvas as HTMLCanvasElement & { dispatch(type: string, event: any): void };
  }

  const geom: GridGeometry = {
    cellOriginPx: () => ({ x: 0, y: 0 }),
    cellSizePx: () => ({ width: 0, height: 0 }),
  };

  const viewport: Viewport = { scrollX: 0, scrollY: 0, width: 500, height: 500, dpr: 1 };

  it("patches xfrm ext on resize commit (pointerup), not during pointermove", () => {
    const startCx = pxToEmu(100);
    const startCy = pxToEmu(100);
    const obj: DrawingObject = {
      id: 1,
      kind: { type: "image", imageId: "img1.png" },
      anchor: {
        type: "absolute",
        pos: { xEmu: 0, yEmu: 0 },
        size: { cx: startCx, cy: startCy },
      },
      zOrder: 0,
      preserved: {
        "xlsx.pic_xml": `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="Picture 1"/></xdr:nvPicPr><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${startCx}" cy="${startCy}"/></a:xfrm></xdr:spPr></xdr:pic>`,
      },
    };

    let objects: DrawingObject[] = [obj];
    const canvas = createStubCanvas();

    new DrawingInteractionController(canvas, geom, {
      getViewport: () => viewport,
      getObjects: () => objects,
      setObjects: (next) => {
        objects = next;
      },
    });

    // Start resizing at the bottom-right corner (se handle).
    canvas.dispatch("pointerdown", { offsetX: 100, offsetY: 100, pointerId: 1 });
    canvas.dispatch("pointermove", { offsetX: 120, offsetY: 130, pointerId: 1 });

    // Pointermove updates the anchor, but should not patch xml yet.
    expect(objects[0]!.anchor).toMatchObject({
      type: "absolute",
      size: { cx: pxToEmu(120), cy: pxToEmu(130) },
    });
    expect(objects[0]!.preserved?.["xlsx.pic_xml"]).toContain(`cx="${startCx}"`);
    expect(objects[0]!.preserved?.["xlsx.pic_xml"]).toContain(`cy="${startCy}"`);

    canvas.dispatch("pointerup", { offsetX: 120, offsetY: 130, pointerId: 1 });

    // Commit-time: preserved xml is patched to the new size.
    expect(objects[0]!.preserved?.["xlsx.pic_xml"]).toContain(`cx="${pxToEmu(120)}"`);
    expect(objects[0]!.preserved?.["xlsx.pic_xml"]).toContain(`cy="${pxToEmu(130)}"`);
  });

  it("patches xfrm off on move commit when off is non-zero", () => {
    const startOffX = pxToEmu(10);
    const startOffY = pxToEmu(20);
    const obj: DrawingObject = {
      id: 1,
      kind: {
        type: "shape",
        rawXml: `<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="1" name="Shape 1"/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="${startOffX}" y="${startOffY}"/><a:ext cx="${pxToEmu(
          100,
        )}" cy="${pxToEmu(100)}"/></a:xfrm></xdr:spPr></xdr:sp>`,
      },
      anchor: {
        type: "absolute",
        pos: { xEmu: 0, yEmu: 0 },
        size: { cx: pxToEmu(100), cy: pxToEmu(100) },
      },
      zOrder: 0,
    };

    let objects: DrawingObject[] = [obj];
    const canvas = createStubCanvas();

    new DrawingInteractionController(canvas, geom, {
      getViewport: () => viewport,
      getObjects: () => objects,
      setObjects: (next) => {
        objects = next;
      },
    });

    canvas.dispatch("pointerdown", { offsetX: 50, offsetY: 50, pointerId: 2 });
    canvas.dispatch("pointermove", { offsetX: 55, offsetY: 57, pointerId: 2 });
    canvas.dispatch("pointerup", { offsetX: 55, offsetY: 57, pointerId: 2 });

    const xml = (objects[0]!.kind as any).rawXml as string;
    expect(xml).toContain(`x="${pxToEmu(15)}"`);
    expect(xml).toContain(`y="${pxToEmu(27)}"`);
  });
});
