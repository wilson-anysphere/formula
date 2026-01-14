import { describe, expect, it } from "vitest";

import { pxToEmu } from "../overlay";
import type { DrawingObject } from "../types";
import {
  bringForward,
  bringToFront,
  deleteSelected,
  duplicateSelected,
  normalizeZOrder,
  sendBackward,
  sendToBack,
} from "../commands";

function makeObject(id: number, zOrder: number): DrawingObject {
  return {
    id,
    kind: { type: "shape", label: `shape-${id}` },
    anchor: { type: "absolute", pos: { xEmu: pxToEmu(id * 10), yEmu: pxToEmu(id * 20) }, size: { cx: pxToEmu(50), cy: pxToEmu(40) } },
    zOrder,
  };
}

describe("drawings z-order + commands", () => {
  it("normalizeZOrder assigns dense 0..n-1 order deterministically", () => {
    const objects = [makeObject(1, 10), makeObject(2, 10), makeObject(3, -5)];
    const normalized = normalizeZOrder(objects);

    // Preserve input array order.
    expect(normalized.map((o) => o.id)).toEqual([1, 2, 3]);
    // Deterministic stacking: lowest zOrder first; tie-break by id.
    expect(normalized.find((o) => o.id === 3)!.zOrder).toBe(0);
    expect(normalized.find((o) => o.id === 1)!.zOrder).toBe(1);
    expect(normalized.find((o) => o.id === 2)!.zOrder).toBe(2);
  });

  it("deleteSelected removes object and renormalizes", () => {
    const objects = [makeObject(1, 0), makeObject(2, 1), makeObject(3, 2)];
    const next = deleteSelected(objects, 2);
    expect(next.map((o) => o.id)).toEqual([1, 3]);
    expect(next.map((o) => o.zOrder)).toEqual([0, 1]);
  });

  it("duplicateSelected clones, offsets anchor, and places on top", () => {
    const objects = [makeObject(1, 0)];
    const res = duplicateSelected(objects, 1);
    expect(res).not.toBeNull();
    if (!res) return;
    expect(res.objects).toHaveLength(2);
    const original = res.objects.find((o) => o.id === 1)!;
    const dup = res.objects.find((o) => o.id === res.duplicatedId)!;
    expect(dup.id).not.toBe(original.id);
    expect(dup.zOrder).toBe(1);
    expect(dup.anchor.type).toBe("absolute");
    if (dup.anchor.type === "absolute" && original.anchor.type === "absolute") {
      expect(dup.anchor.pos.xEmu - original.anchor.pos.xEmu).toBe(pxToEmu(10));
      expect(dup.anchor.pos.yEmu - original.anchor.pos.yEmu).toBe(pxToEmu(10));
    }
  });

  it("duplicateSelected patches preserved DrawingML fragments for the new id", () => {
    const obj: DrawingObject = {
      id: 1,
      kind: { type: "image", imageId: "img1" },
      anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: pxToEmu(10), cy: pxToEmu(10) } },
      zOrder: 0,
      preserved: {
        "xlsx.pic_xml": `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="Picture 1"/></xdr:nvPicPr></xdr:pic>`,
      },
    };
    const res = duplicateSelected([obj], 1);
    expect(res).not.toBeNull();
    if (!res) return;

    const dup = res.objects.find((o) => o.id === res.duplicatedId)!;
    const xml = dup.preserved?.["xlsx.pic_xml"] ?? "";
    expect(xml).toContain(`id=\"${res.duplicatedId}\"`);
    expect(xml).not.toContain(`id=\"1\"`);
  });

  it("bringToFront and sendToBack reorders deterministically", () => {
    const objects = [makeObject(1, 0), makeObject(2, 1), makeObject(3, 2)];
    const front = bringToFront(objects, 1);
    expect(front.map((o) => o.zOrder)).toEqual([2, 0, 1]);
    const back = sendToBack(front, 3);
    expect(back.map((o) => o.zOrder)).toEqual([2, 1, 0]);
  });

  it("bringForward and sendBackward swap adjacent objects in z-stack", () => {
    const objects = [makeObject(1, 0), makeObject(2, 1), makeObject(3, 2)];
    const forward = bringForward(objects, 1);
    expect(forward.map((o) => o.zOrder)).toEqual([1, 0, 2]);
    const backward = sendBackward(objects, 3);
    expect(backward.map((o) => o.zOrder)).toEqual([0, 2, 1]);
  });
});
