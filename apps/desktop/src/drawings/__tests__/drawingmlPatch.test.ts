import { describe, expect, it, vi } from "vitest";

import {
  patchAnchorExt,
  patchAnchorPoint,
  patchAnchorPos,
  patchNvPrId,
  patchXfrmExt,
  patchXfrmOff,
  patchXfrmRot,
} from "../drawingml/patch";
import { duplicateDrawingObject } from "../duplicate";
import { DrawingInteractionController } from "../interaction";
import { pxToEmu } from "../overlay";
import { getRotationHandleCenter } from "../selectionHandles";
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

  it("patchNvPrId is prefix-tolerant (non-xdr prefixes)", () => {
    const input = `<p:cNvPr id='5' name='Picture 5'/>`;
    const out = patchNvPrId(input, 6);
    expect(out).toContain(`id='6'`);
    expect(out).toContain(`name='Picture 5'`);
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

  it("patchXfrmExt is prefix-tolerant (non-a/xdr prefixes)", () => {
    const input = `<foo:xfrm><bar:off x="0" y="0"/><bar:ext cx="1" cy="2"/></foo:xfrm>`;
    const out = patchXfrmExt(input, 7, 8);
    expect(out).toContain(`<bar:ext cx="7" cy="8"/>`);
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

  it("patchXfrmRot inserts rot on xfrm when missing", () => {
    const input = `<a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="2"/></a:xfrm>`;
    const out = patchXfrmRot(input, 90);
    expect(out).toContain(`<a:xfrm rot="5400000">`);
  });

  it("patchXfrmRot patches rot when present", () => {
    const input = `<a:xfrm rot="0"><a:off x="0" y="0"/><a:ext cx="1" cy="2"/></a:xfrm>`;
    const out = patchXfrmRot(input, 15);
    expect(out).toContain(`<a:xfrm rot="900000">`);
  });

  it("patchAnchorPoint patches <from>/<to> blocks in full anchor XML", () => {
    const input =
      `<xdr:twoCellAnchor>` +
      `<xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
      `<xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>` +
      `</xdr:twoCellAnchor>`;

    const out = patchAnchorPoint(input, "from", { col: 2, row: 3, colOffEmu: 10, rowOffEmu: 20 });
    expect(out).toContain(`<xdr:from><xdr:col>2</xdr:col>`);
    expect(out).toContain(`<xdr:row>3</xdr:row>`);
    expect(out).toContain(`<xdr:colOff>10</xdr:colOff>`);
    expect(out).toContain(`<xdr:rowOff>20</xdr:rowOff>`);
    // Ensure we don't accidentally patch the <to> block.
    expect(out).toContain(`<xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff>`);
  });

  it("patchAnchorPos/Ext patch absoluteAnchor wrapper fields without touching xfrm ext", () => {
    const input =
      `<xdr:absoluteAnchor>` +
      `<xdr:pos x="0" y="0"/><xdr:ext cx="1" cy="2"/>` +
      `<xdr:sp><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="200"/></a:xfrm></xdr:spPr></xdr:sp>` +
      `</xdr:absoluteAnchor>`;

    let out = patchAnchorPos(input, 10, 20);
    out = patchAnchorExt(out, 7, 8);
    expect(out).toContain(`<xdr:pos x="10" y="20"/>`);
    expect(out).toContain(`<xdr:ext cx="7" cy="8"/>`);
    // Inner xfrm ext should remain unchanged (patched elsewhere).
    expect(out).toContain(`<a:ext cx="100" cy="200"/>`);
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
  function createPointerEvent(init: { clientX: number; clientY: number; pointerId: number }) {
    return {
      ...init,
      defaultPrevented: false,
      _propagationStopped: false,
      _immediateStopped: false,
      preventDefault() {
        this.defaultPrevented = true;
      },
      stopPropagation() {
        this._propagationStopped = true;
      },
      stopImmediatePropagation() {
        this._immediateStopped = true;
        this._propagationStopped = true;
      },
    };
  }

  function createStubElement(
    rect: { left: number; top: number; width: number; height: number } = { left: 0, top: 0, width: 500, height: 500 },
  ) {
    const listeners = new Map<string, Array<(e: any) => void>>();
    const element: any = {
      addEventListener: (type: string, cb: (e: any) => void) => {
        const list = listeners.get(type) ?? [];
        list.push(cb);
        listeners.set(type, list);
      },
      removeEventListener: (type: string, cb: (e: any) => void) => {
        const list = listeners.get(type) ?? [];
        listeners.set(
          type,
          list.filter((v) => v !== cb),
        );
      },
      getBoundingClientRect: () =>
        ({
          left: rect.left,
          top: rect.top,
          right: rect.left + rect.width,
          bottom: rect.top + rect.height,
          width: rect.width,
          height: rect.height,
          x: rect.left,
          y: rect.top,
          toJSON: () => ({}),
        }) as unknown as DOMRect,
      setPointerCapture: () => {},
      releasePointerCapture: () => {},
      style: { cursor: "default" },
      dispatch: (type: string, event: any) => {
        for (const cb of listeners.get(type) ?? []) cb(event);
      },
    };
    return element as HTMLElement & { dispatch(type: string, event: any): void };
  }

  const CELL_W = 100;
  const CELL_H = 100;
  const geom: GridGeometry = {
    cellOriginPx: (cell) => ({ x: cell.col * CELL_W, y: cell.row * CELL_H }),
    cellSizePx: () => ({ width: CELL_W, height: CELL_H }),
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
    const el = createStubElement();

    new DrawingInteractionController(el, geom, {
      getViewport: () => viewport,
      getObjects: () => objects,
      setObjects: (next) => {
        objects = next;
      },
    });

    // Start resizing at the bottom-right corner (se handle).
    el.dispatch("pointerdown", createPointerEvent({ clientX: 100, clientY: 100, pointerId: 1 }));
    el.dispatch("pointermove", createPointerEvent({ clientX: 120, clientY: 130, pointerId: 1 }));

    // Pointermove updates the anchor, but should not patch xml yet.
    expect(objects[0]!.anchor).toMatchObject({
      type: "absolute",
      size: { cx: pxToEmu(120), cy: pxToEmu(130) },
    });
    expect(objects[0]!.preserved?.["xlsx.pic_xml"]).toContain(`cx="${startCx}"`);
    expect(objects[0]!.preserved?.["xlsx.pic_xml"]).toContain(`cy="${startCy}"`);

    el.dispatch("pointerup", createPointerEvent({ clientX: 120, clientY: 130, pointerId: 1 }));

    // Commit-time: preserved xml is patched to the new size.
    expect(objects[0]!.preserved?.["xlsx.pic_xml"]).toContain(`cx="${pxToEmu(120)}"`);
    expect(objects[0]!.preserved?.["xlsx.pic_xml"]).toContain(`cy="${pxToEmu(130)}"`);
  });

  it("fires onInteractionCommit once after a resize, with patched preserved xml", () => {
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
    const commits: Array<any> = [];
    const el = createStubElement();

    new DrawingInteractionController(el, geom, {
      getViewport: () => viewport,
      getObjects: () => objects,
      setObjects: (next) => {
        objects = next;
      },
      onInteractionCommit: (payload) => commits.push(payload),
    });

    // Start resizing at the bottom-right corner (se handle).
    el.dispatch("pointerdown", createPointerEvent({ clientX: 100, clientY: 100, pointerId: 10 }));
    el.dispatch("pointermove", createPointerEvent({ clientX: 120, clientY: 130, pointerId: 10 }));
    el.dispatch("pointerup", createPointerEvent({ clientX: 120, clientY: 130, pointerId: 10 }));

    expect(commits).toHaveLength(1);
    const payload = commits[0]!;
    expect(payload.kind).toBe("resize");
    expect(payload.id).toBe(1);

    expect(payload.before.anchor).toMatchObject({
      type: "absolute",
      size: { cx: startCx, cy: startCy },
    });
    expect(payload.after.anchor).toMatchObject({
      type: "absolute",
      size: { cx: pxToEmu(120), cy: pxToEmu(130) },
    });

    // Commit callback should see the patched DrawingML fragment.
    expect(payload.after.preserved?.["xlsx.pic_xml"]).toContain(`cx="${pxToEmu(120)}"`);
    expect(payload.after.preserved?.["xlsx.pic_xml"]).toContain(`cy="${pxToEmu(130)}"`);
    expect(payload.objects.find((o: DrawingObject) => o.id === 1)).toEqual(payload.after);
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
    const el = createStubElement();

    new DrawingInteractionController(el, geom, {
      getViewport: () => viewport,
      getObjects: () => objects,
      setObjects: (next) => {
        objects = next;
      },
    });

    el.dispatch("pointerdown", createPointerEvent({ clientX: 50, clientY: 50, pointerId: 2 }));
    el.dispatch("pointermove", createPointerEvent({ clientX: 55, clientY: 57, pointerId: 2 }));
    el.dispatch("pointerup", createPointerEvent({ clientX: 55, clientY: 57, pointerId: 2 }));

    const xml = (objects[0]!.kind as any).rawXml as string;
    expect(xml).toContain(`x="${pxToEmu(15)}"`);
    expect(xml).toContain(`y="${pxToEmu(27)}"`);
  });

  it("fires onInteractionCommit once after a move, with patched rawXml", () => {
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
    const commits: Array<any> = [];
    const el = createStubElement();

    new DrawingInteractionController(el, geom, {
      getViewport: () => viewport,
      getObjects: () => objects,
      setObjects: (next) => {
        objects = next;
      },
      onInteractionCommit: (payload) => commits.push(payload),
    });

    el.dispatch("pointerdown", createPointerEvent({ clientX: 50, clientY: 50, pointerId: 11 }));
    el.dispatch("pointermove", createPointerEvent({ clientX: 55, clientY: 57, pointerId: 11 }));
    el.dispatch("pointerup", createPointerEvent({ clientX: 55, clientY: 57, pointerId: 11 }));

    expect(commits).toHaveLength(1);
    const payload = commits[0]!;
    expect(payload.kind).toBe("move");
    expect(payload.id).toBe(1);

    expect(payload.after.anchor).toMatchObject({
      type: "absolute",
      pos: { xEmu: pxToEmu(5), yEmu: pxToEmu(7) },
    });

    const xml = (payload.after.kind as any).rawXml as string;
    expect(xml).toContain(`x="${pxToEmu(15)}"`);
    expect(xml).toContain(`y="${pxToEmu(27)}"`);
  });

  it("fires onInteractionCommit once after a rotate, with transform + patched rawXml", () => {
    const obj: DrawingObject = {
      id: 1,
      kind: {
        type: "shape",
        rawXml: `<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="1" name="Shape 1"/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${pxToEmu(
          80,
        )}" cy="${pxToEmu(40)}"/></a:xfrm></xdr:spPr></xdr:sp>`,
      },
      anchor: {
        type: "absolute",
        pos: { xEmu: pxToEmu(100), yEmu: pxToEmu(100) },
        size: { cx: pxToEmu(80), cy: pxToEmu(40) },
      },
      zOrder: 0,
    };

    let objects: DrawingObject[] = [obj];
    const commits: Array<any> = [];
    const el = createStubElement();

    const controller = new DrawingInteractionController(el, geom, {
      getViewport: () => viewport,
      getObjects: () => objects,
      setObjects: (next) => {
        objects = next;
      },
      onInteractionCommit: (payload) => commits.push(payload),
    });

    // Rotation handles only work for the selected object.
    controller.setSelectedId(1);

    const bounds = { x: 100, y: 100, width: 80, height: 40 };
    const rotHandle = getRotationHandleCenter(bounds);
    const centerX = bounds.x + bounds.width / 2;
    const centerY = bounds.y + bounds.height / 2;

    el.dispatch(
      "pointerdown",
      createPointerEvent({ clientX: rotHandle.x, clientY: rotHandle.y, pointerId: 12 }),
    );
    // Drag to the right of center (roughly +90° clockwise from "above").
    el.dispatch("pointermove", createPointerEvent({ clientX: centerX + 100, clientY: centerY, pointerId: 12 }));
    el.dispatch("pointerup", createPointerEvent({ clientX: centerX + 100, clientY: centerY, pointerId: 12 }));

    expect(commits).toHaveLength(1);
    const payload = commits[0]!;
    expect(payload.kind).toBe("rotate");
    expect(payload.id).toBe(1);
    expect(payload.after.transform?.rotationDeg).toBeCloseTo(90);

    // Commit callback should see the patched DrawingML rot attribute.
    const xml = (payload.after.kind as any).rawXml as string;
    expect(xml).toContain(`rot="5400000"`);
  });

  it("does not throw if onInteractionCommit throws (best-effort)", () => {
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
    const el = createStubElement();
    const commitObjects = vi.fn((next: DrawingObject[]) => {
      objects = next;
    });
    const endBatch = vi.fn();

    new DrawingInteractionController(el, geom, {
      getViewport: () => viewport,
      getObjects: () => objects,
      setObjects: (next) => {
        objects = next;
      },
      beginBatch: vi.fn(),
      endBatch,
      cancelBatch: vi.fn(),
      commitObjects,
      onInteractionCommit: () => {
        throw new Error("boom");
      },
    });

    el.dispatch("pointerdown", createPointerEvent({ clientX: 50, clientY: 50, pointerId: 13 }));
    el.dispatch("pointermove", createPointerEvent({ clientX: 55, clientY: 57, pointerId: 13 }));

    expect(() => el.dispatch("pointerup", createPointerEvent({ clientX: 55, clientY: 57, pointerId: 13 }))).not.toThrow();
    expect(commitObjects).toHaveBeenCalledTimes(1);
    expect(endBatch).toHaveBeenCalledTimes(1);
  });

  it("patches full-anchor wrapper <from>/<to> blocks on move commit for unknown objects", () => {
    const obj: DrawingObject = {
      id: 1,
      kind: {
        type: "unknown",
        rawXml:
          `<xdr:twoCellAnchor>` +
          `<xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
          `<xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>` +
          `<xdr:clientData/>` +
          `</xdr:twoCellAnchor>`,
      },
      anchor: {
        type: "twoCell",
        from: { cell: { row: 0, col: 0 }, offset: { xEmu: 0, yEmu: 0 } },
        to: { cell: { row: 1, col: 1 }, offset: { xEmu: 0, yEmu: 0 } },
      },
      zOrder: 0,
    };

    let objects: DrawingObject[] = [obj];
    const el = createStubElement();

    new DrawingInteractionController(el, geom, {
      getViewport: () => viewport,
      getObjects: () => objects,
      setObjects: (next) => {
        objects = next;
      },
    });

    el.dispatch("pointerdown", createPointerEvent({ clientX: 50, clientY: 50, pointerId: 3 }));
    el.dispatch("pointermove", createPointerEvent({ clientX: 60, clientY: 70, pointerId: 3 }));

    // Pointermove updates anchor but should not patch xml yet.
    const before = (objects[0]!.kind as any).rawXml as string;
    expect(before).toContain(`<xdr:colOff>0</xdr:colOff>`);
    expect(before).toContain(`<xdr:rowOff>0</xdr:rowOff>`);

    el.dispatch("pointerup", createPointerEvent({ clientX: 60, clientY: 70, pointerId: 3 }));

    const xml = (objects[0]!.kind as any).rawXml as string;
    // Both from + to should have the same shifted offsets.
    expect(xml.match(new RegExp(`<xdr:colOff>${pxToEmu(10)}</xdr:colOff>`, "g"))?.length).toBe(2);
    expect(xml.match(new RegExp(`<xdr:rowOff>${pxToEmu(20)}</xdr:rowOff>`, "g"))?.length).toBe(2);
  });

  it("patches full-anchor wrapper <pos>/<ext> on resize commit for unknown objects", () => {
    const startCx = pxToEmu(100);
    const startCy = pxToEmu(100);
    const obj: DrawingObject = {
      id: 1,
      kind: {
        type: "unknown",
        rawXml:
          `<xdr:absoluteAnchor>` +
          `<xdr:pos x="0" y="0"/><xdr:ext cx="${startCx}" cy="${startCy}"/>` +
          `<xdr:sp><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${startCx}" cy="${startCy}"/></a:xfrm></xdr:spPr></xdr:sp>` +
          `<xdr:clientData/>` +
          `</xdr:absoluteAnchor>`,
      },
      anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: startCx, cy: startCy } },
      zOrder: 0,
    };

    let objects: DrawingObject[] = [obj];
    const el = createStubElement();

    new DrawingInteractionController(el, geom, {
      getViewport: () => viewport,
      getObjects: () => objects,
      setObjects: (next) => {
        objects = next;
      },
    });

    el.dispatch("pointerdown", createPointerEvent({ clientX: 100, clientY: 100, pointerId: 4 }));
    el.dispatch("pointermove", createPointerEvent({ clientX: 120, clientY: 130, pointerId: 4 }));

    // Pointermove updates anchor but should not patch xml yet.
    expect(objects[0]!.anchor).toMatchObject({
      type: "absolute",
      size: { cx: pxToEmu(120), cy: pxToEmu(130) },
    });
    const before = (objects[0]!.kind as any).rawXml as string;
    expect(before).toContain(`cx="${startCx}"`);
    expect(before).toContain(`cy="${startCy}"`);

    el.dispatch("pointerup", createPointerEvent({ clientX: 120, clientY: 130, pointerId: 4 }));

    const xml = (objects[0]!.kind as any).rawXml as string;
    // Outer anchor wrapper ext updated.
    expect(xml).toContain(`<xdr:ext cx="${pxToEmu(120)}" cy="${pxToEmu(130)}"/>`);
    // Inner xfrm ext updated too.
    expect(xml).toContain(`<a:ext cx="${pxToEmu(120)}" cy="${pxToEmu(130)}"/>`);
  });
});
