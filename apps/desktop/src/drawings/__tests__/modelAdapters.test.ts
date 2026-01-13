import { describe, expect, it } from "vitest";

import { convertModelDrawingObjectToUiDrawingObject, convertModelImageStoreToUiImageStore } from "../modelAdapters";

describe("drawings/modelAdapters", () => {
  it("converts a TwoCell Image drawing object", () => {
    const model = {
      id: { 0: 12 },
      kind: { Image: { image_id: "image1.png" } },
      anchor: {
        TwoCell: {
          from: { cell: { row: 0, col: 0 }, offset: { x_emu: 0, y_emu: 0 } },
          to: { cell: { row: 4, col: 2 }, offset: { x_emu: 914400, y_emu: 457200 } },
        },
      },
      z_order: 3,
      size: { cx: 2000, cy: 1000 },
    };

    const ui = convertModelDrawingObjectToUiDrawingObject(model);
    expect(ui).toEqual({
      id: 12,
      kind: { type: "image", imageId: "image1.png" },
      anchor: {
        type: "twoCell",
        from: { cell: { row: 0, col: 0 }, offset: { xEmu: 0, yEmu: 0 } },
        to: { cell: { row: 4, col: 2 }, offset: { xEmu: 914400, yEmu: 457200 } },
      },
      zOrder: 3,
      size: { cx: 2000, cy: 1000 },
    });
  });

  it("converts a OneCell Shape drawing object", () => {
    const model = {
      id: 1,
      kind: { Shape: { raw_xml: "<xdr:sp><a:prstGeom prst=\"rect\"/></xdr:sp>" } },
      anchor: {
        OneCell: {
          from: { cell: { row: 1, col: 2 }, offset: { x_emu: 9525, y_emu: 19050 } },
          ext: { cx: 111, cy: 222 },
        },
      },
      z_order: 0,
    };

    const ui = convertModelDrawingObjectToUiDrawingObject(model);
    expect(ui.id).toBe(1);
    expect(ui.kind.type).toBe("shape");
    expect(ui.anchor).toEqual({
      type: "oneCell",
      from: { cell: { row: 1, col: 2 }, offset: { xEmu: 9525, yEmu: 19050 } },
      size: { cx: 111, cy: 222 },
    });
  });

  it("converts a ChartPlaceholder to a chart kind with a label", () => {
    const model = {
      id: "7",
      kind: {
        ChartPlaceholder: {
          rel_id: "rId5",
          raw_xml: "<xdr:graphicFrame><a:graphic/></xdr:graphicFrame>",
        },
      },
      anchor: {
        Absolute: {
          pos: { x_emu: 0, y_emu: 0 },
          ext: { cx: 10, cy: 20 },
        },
      },
      z_order: 10,
    };

    const ui = convertModelDrawingObjectToUiDrawingObject(model);
    expect(ui.id).toBe(7);
    expect(ui.kind.type).toBe("chart");
    expect(ui.kind).toMatchObject({ chartId: "rId5" });
    expect(ui.kind.label).toContain("rId5");
  });

  it("converts ImageStore bytes + content_type into a UI ImageStore", () => {
    const modelImages = {
      images: {
        "image1.png": { bytes: [1, 2, 3], content_type: "image/png" },
        "image2.jpg": { bytes: [255, 216, 255], content_type: null },
      },
    };

    const store = convertModelImageStoreToUiImageStore(modelImages);
    expect(store.get("image1.png")).toEqual({
      id: "image1.png",
      bytes: new Uint8Array([1, 2, 3]),
      mimeType: "image/png",
    });
    expect(store.get("image2.jpg")?.mimeType).toBe("image/jpeg");
  });
});

