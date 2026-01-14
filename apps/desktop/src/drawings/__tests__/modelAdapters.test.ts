import { describe, expect, it } from "vitest";

import {
  convertDocumentSheetDrawingsToUiDrawingObjects,
  convertModelDrawingObjectToUiDrawingObject,
  convertModelImageStoreToUiImageStore,
  convertModelWorkbookDrawingsToUiDrawingLayer,
} from "../modelAdapters";
import { pxToEmu } from "../overlay";

function createPngHeaderBytes(width: number, height: number): number[] {
  const bytes = new Uint8Array(24);
  bytes.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 0);
  // 13-byte IHDR chunk length.
  bytes[8] = 0x00;
  bytes[9] = 0x00;
  bytes[10] = 0x00;
  bytes[11] = 0x0d;
  // IHDR chunk type.
  bytes[12] = 0x49;
  bytes[13] = 0x48;
  bytes[14] = 0x44;
  bytes[15] = 0x52;

  const view = new DataView(bytes.buffer);
  view.setUint32(16, width, false);
  view.setUint32(20, height, false);
  return Array.from(bytes);
}

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

  it("accepts UI/other internally-tagged representations for anchor + kind", () => {
    const model = {
      id: 123,
      kind: { type: "image", imageId: "image1.png" },
      anchor: {
        type: "absolute",
        pos: { xEmu: 5, yEmu: 7 },
        size: { cx: 20, cy: 10 },
      },
      zOrder: 1,
    };

    const ui = convertModelDrawingObjectToUiDrawingObject(model);
    expect(ui).toEqual({
      id: 123,
      kind: { type: "image", imageId: "image1.png" },
      anchor: { type: "absolute", pos: { xEmu: 5, yEmu: 7 }, size: { cx: 20, cy: 10 } },
      zOrder: 1,
      size: undefined,
    });
  });

  it("accepts internally-tagged Anchor variants (kind field)", () => {
    const model = {
      id: 1,
      kind: { Image: { image_id: "image1.png" } },
      anchor: {
        kind: "TwoCell",
        from: { cell: { row: 0, col: 0 }, offset: { x_emu: 0, y_emu: 0 } },
        to: { cell: { row: 1, col: 1 }, offset: { x_emu: 0, y_emu: 0 } },
      },
      z_order: 0,
    };

    const ui = convertModelDrawingObjectToUiDrawingObject(model);
    expect(ui.anchor.type).toBe("twoCell");
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

    const ui = convertModelDrawingObjectToUiDrawingObject(model, { sheetId: "sheet1" });
    expect(ui.id).toBe(7);
    expect(ui.kind.type).toBe("chart");
    expect(ui.kind).toMatchObject({ chartId: "sheet1:7" });
    expect(ui.kind.label).toContain("rId5");
  });

  it("hashes unsafe drawing object ids (beyond MAX_SAFE_INTEGER) into stable safe integers", () => {
    const unsafeId = "9007199254740993";
    const model = {
      id: unsafeId,
      kind: { Image: { image_id: "image1.png" } },
      anchor: {
        Absolute: {
          pos: { x_emu: 0, y_emu: 0 },
          ext: { cx: 10, cy: 20 },
        },
      },
      z_order: 0,
    };

    const ui1 = convertModelDrawingObjectToUiDrawingObject(model);
    const ui2 = convertModelDrawingObjectToUiDrawingObject(model);

    expect(Number.isSafeInteger(ui1.id)).toBe(true);
    expect(ui1.id).toBeLessThan(0);
    expect(ui1.id).toBe(ui2.id);

    const parsed = Number(unsafeId);
    expect(Number.isSafeInteger(parsed)).toBe(false);
    expect(ui1.id).not.toBe(parsed);
  });

  it("hashes negative drawing object ids into the reserved large-magnitude namespace (avoids chart collisions)", () => {
    const model = {
      id: -123,
      kind: { Image: { image_id: "image1.png" } },
      anchor: {
        Absolute: {
          pos: { x_emu: 0, y_emu: 0 },
          ext: { cx: 10, cy: 20 },
        },
      },
      z_order: 0,
    };

    const ui1 = convertModelDrawingObjectToUiDrawingObject(model);
    const ui2 = convertModelDrawingObjectToUiDrawingObject(model);

    expect(Number.isSafeInteger(ui1.id)).toBe(true);
    // Hashed ids are offset by 2^33 (see parseDrawingObjectId) to stay disjoint from ChartStore ids.
    expect(ui1.id).toBeLessThanOrEqual(-(2 ** 33));
    expect(ui1.id).toBe(ui2.id);
    expect(ui1.id).not.toBe(-123);
  });

  it("hashes non-canonical numeric string ids so distinct raw ids do not collide", () => {
    const makeModel = (id: string) => ({
      id,
      kind: { Image: { image_id: "image1.png" } },
      anchor: {
        Absolute: {
          pos: { x_emu: 0, y_emu: 0 },
          ext: { cx: 10, cy: 20 },
        },
      },
      z_order: 0,
    });

    const uiLeadingZeros1 = convertModelDrawingObjectToUiDrawingObject(makeModel("001"));
    const uiLeadingZeros2 = convertModelDrawingObjectToUiDrawingObject(makeModel("001"));
    expect(Number.isSafeInteger(uiLeadingZeros1.id)).toBe(true);
    expect(uiLeadingZeros1.id).toBeLessThanOrEqual(-(2 ** 33));
    expect(uiLeadingZeros1.id).toBe(uiLeadingZeros2.id);
    expect(uiLeadingZeros1.id).not.toBe(1);

    const uiExponent = convertModelDrawingObjectToUiDrawingObject(makeModel("1e3"));
    expect(Number.isSafeInteger(uiExponent.id)).toBe(true);
    expect(uiExponent.id).toBeLessThanOrEqual(-(2 ** 33));
    expect(uiExponent.id).not.toBe(1000);
  });

  it("trims non-numeric string ids before hashing (stable across DocumentController normalization)", () => {
    const makeModel = (id: string) => ({
      id,
      kind: { Image: { image_id: "image1.png" } },
      anchor: {
        Absolute: {
          pos: { x_emu: 0, y_emu: 0 },
          ext: { cx: 10, cy: 20 },
        },
      },
      z_order: 0,
    });

    const uiTrimmed = convertModelDrawingObjectToUiDrawingObject(makeModel("foo"));
    const uiWithWhitespace = convertModelDrawingObjectToUiDrawingObject(makeModel("  foo  "));
    expect(uiWithWhitespace.id).toBe(uiTrimmed.id);
    expect(uiWithWhitespace.id).toBeLessThanOrEqual(-(2 ** 33));
  });

  it("hashes very long string ids using a bounded summary (stable + whitespace-tolerant)", () => {
    const long = "a".repeat(10_000);
    const makeModel = (id: string) => ({
      id,
      kind: { Image: { image_id: "image1.png" } },
      anchor: {
        Absolute: {
          pos: { x_emu: 0, y_emu: 0 },
          ext: { cx: 10, cy: 20 },
        },
      },
      z_order: 0,
    });

    const ui1 = convertModelDrawingObjectToUiDrawingObject(makeModel(long));
    const ui2 = convertModelDrawingObjectToUiDrawingObject(makeModel(`  ${long}  `));
    const ui3 = convertModelDrawingObjectToUiDrawingObject(makeModel(long));

    expect(Number.isSafeInteger(ui1.id)).toBe(true);
    expect(ui1.id).toBeLessThanOrEqual(-(2 ** 33));
    expect(ui2.id).toBe(ui1.id);
    expect(ui3.id).toBe(ui1.id);
  });

  it("does not crash when drawing object ids are missing/undefined", () => {
    const model = {
      // Intentionally omit `id`.
      kind: { Image: { image_id: "image1.png" } },
      anchor: {
        Absolute: {
          pos: { x_emu: 0, y_emu: 0 },
          ext: { cx: 10, cy: 20 },
        },
      },
      z_order: 0,
    };

    const ui1 = convertModelDrawingObjectToUiDrawingObject(model);
    const ui2 = convertModelDrawingObjectToUiDrawingObject(model);

    expect(Number.isSafeInteger(ui1.id)).toBe(true);
    expect(ui1.id).toBeLessThan(0);
    // The fallback id should be stable for the same input payload.
    expect(ui1.id).toBe(ui2.id);
  });

  it("extracts DrawingML transform metadata from preserved pic_xml for images", () => {
    const model = {
      id: 1,
      kind: { Image: { image_id: "image1.png" } },
      anchor: {
        Absolute: {
          pos: { x_emu: 0, y_emu: 0 },
          ext: { cx: 1000, cy: 500 },
        },
      },
      z_order: 0,
      preserved: {
        "xlsx.pic_xml": `
          <xdr:pic xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
                   xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <xdr:spPr>
              <a:xfrm rot="5400000" flipV="1">
                <a:off x="0" y="0"/>
                <a:ext cx="1000" cy="500"/>
              </a:xfrm>
            </xdr:spPr>
          </xdr:pic>
        `,
      },
    };

    const ui = convertModelDrawingObjectToUiDrawingObject(model);
    expect(ui.transform).toEqual({ rotationDeg: 90, flipH: false, flipV: true });
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

  it("parses ImageStore bytes from JSON-friendly encodings", () => {
    const modelImages = {
      images: {
        // JSON.stringify(Uint8Array([7,8]))-style object.
        "obj.png": { bytes: { 0: 7, 1: 8 }, content_type: "image/png" },
        // Node Buffer JSON representation.
        "buf.png": { bytes: { type: "Buffer", data: [9, 10, 11] }, content_type: "image/png" },
        // Base64 string.
        "b64.png": { bytes: "AQID", content_type: "image/png" }, // [1,2,3]
      },
    };

    const store = convertModelImageStoreToUiImageStore(modelImages);
    expect(store.get("obj.png")?.bytes).toEqual(new Uint8Array([7, 8]));
    expect(store.get("buf.png")?.bytes).toEqual(new Uint8Array([9, 10, 11]));
    expect(store.get("b64.png")?.bytes).toEqual(new Uint8Array([1, 2, 3]));
  });

  it("ignores invalid image byte payloads (best-effort)", () => {
    const modelImages = {
      images: {
        "good.png": { bytes: [1, 2, 3], content_type: "image/png" },
        // Malformed: not an array / base64 / typed-array-ish object.
        "bad.png": { bytes: { hello: "world" }, content_type: "image/png" },
      },
    };

    const store = convertModelImageStoreToUiImageStore(modelImages);
    expect(store.get("good.png")?.bytes).toEqual(new Uint8Array([1, 2, 3]));
    expect(store.get("bad.png")).toBeUndefined();
  });

  it("ignores images with oversized header dimensions (best-effort)", () => {
    const modelImages = {
      images: {
        "ok.png": { bytes: [1, 2, 3], content_type: "image/png" },
        // Advertises an extremely large bitmap; should be rejected without aborting conversion.
        "bomb.png": { bytes: createPngHeaderBytes(10_001, 1), content_type: "image/png" },
      },
    };

    const store = convertModelImageStoreToUiImageStore(modelImages);
    expect(store.get("ok.png")?.bytes).toEqual(new Uint8Array([1, 2, 3]));
    expect(store.get("bomb.png")).toBeUndefined();
  });

  it("converts a Workbook snapshot to per-sheet drawings + images store", () => {
    const workbook = {
      images: {
        images: {
          "image1.png": { bytes: [1], content_type: "image/png" },
        },
      },
      sheets: [
        {
          name: "Sheet1",
          drawings: [
            null,
            {
              id: { 0: 12 },
              kind: { Image: { image_id: "image1.png" } },
              anchor: {
                TwoCell: {
                  from: { cell: { row: 0, col: 0 }, offset: { x_emu: 0, y_emu: 0 } },
                  to: { cell: { row: 1, col: 1 }, offset: { x_emu: 0, y_emu: 0 } },
                },
              },
              z_order: 0,
            },
          ],
        },
        { name: "Sheet2", drawings: [] },
      ],
    };

    const ui = convertModelWorkbookDrawingsToUiDrawingLayer(workbook);
    expect(ui.images.get("image1.png")?.mimeType).toBe("image/png");
    expect(ui.drawingsBySheetName.Sheet1).toHaveLength(1);
    expect(ui.drawingsBySheetName.Sheet1?.[0]?.kind).toEqual({ type: "image", imageId: "image1.png" });
    expect(ui.drawingsBySheetName.Sheet2).toEqual([]);
  });

  it("converts DocumentController cell-anchored drawings (pixel size) to overlay anchors", () => {
    const drawings = [
      {
        id: "7",
        zOrder: 1,
        anchor: { type: "cell", sheetId: "Sheet1", row: 0, col: 0 },
        kind: { type: "image", imageId: "img1" },
        size: { width: 120, height: 80 },
      },
    ];

    const ui = convertDocumentSheetDrawingsToUiDrawingObjects(drawings);
    expect(ui).toHaveLength(1);
    expect(ui[0]).toMatchObject({
      id: 7,
      kind: { type: "image", imageId: "img1" },
      anchor: {
        type: "oneCell",
        from: { cell: { row: 0, col: 0 }, offset: { xEmu: 0, yEmu: 0 } },
        size: { cx: pxToEmu(120), cy: pxToEmu(80) },
      },
      zOrder: 1,
      size: { cx: pxToEmu(120), cy: pxToEmu(80) },
    });
  });

  it("defaults missing DocumentController drawing sizes to 100x100px", () => {
    const drawings = [
      {
        id: "d1",
        zOrder: 0,
        anchor: { type: "cell", row: 3, col: 2 },
        kind: { type: "image", imageId: "img1" },
      },
    ];

    const ui = convertDocumentSheetDrawingsToUiDrawingObjects(drawings);
    expect(ui).toHaveLength(1);
    expect(ui[0]?.anchor).toMatchObject({
      type: "oneCell",
      from: { cell: { row: 3, col: 2 }, offset: { xEmu: 0, yEmu: 0 } },
      size: { cx: pxToEmu(100), cy: pxToEmu(100) },
    });
  });

  it("accepts cell anchor pixel offsets in DocumentController drawings", () => {
    const drawings = [
      {
        id: "1",
        zOrder: 0,
        anchor: { type: "cell", row: 0, col: 0, x: 12, y: 34 },
        kind: { type: "image", imageId: "img1" },
        size: { width: 10, height: 10 },
      },
    ];

    const ui = convertDocumentSheetDrawingsToUiDrawingObjects(drawings);
    expect(ui).toHaveLength(1);
    expect(ui[0]?.anchor).toMatchObject({
      type: "oneCell",
      from: { cell: { row: 0, col: 0 }, offset: { xEmu: pxToEmu(12), yEmu: pxToEmu(34) } },
    });
  });

  it("preserves DocumentController drawing transform metadata", () => {
    const drawings = [
      {
        id: "1",
        zOrder: 0,
        anchor: { type: "cell", row: 0, col: 0 },
        kind: { type: "image", imageId: "img1" },
        size: { width: 10, height: 10 },
        transform: { rotationDeg: 30, flipH: true, flipV: false },
      },
    ];

    const ui = convertDocumentSheetDrawingsToUiDrawingObjects(drawings);
    expect(ui).toHaveLength(1);
    expect(ui[0]?.transform).toEqual({ rotationDeg: 30, flipH: true, flipV: false });
  });

  it("ignores malformed DocumentController drawing transform payloads (best-effort)", () => {
    const drawings = [
      {
        id: "1",
        zOrder: 0,
        anchor: { type: "cell", row: 0, col: 0 },
        kind: { type: "image", imageId: "img1" },
        size: { width: 10, height: 10 },
        // rotationDeg is invalid and flipH/flipV are not booleans -> should be ignored.
        transform: { rotationDeg: "not-a-number", flipH: 1, flipV: 0 },
      },
    ];

    const ui = convertDocumentSheetDrawingsToUiDrawingObjects(drawings);
    expect(ui).toHaveLength(1);
    expect(ui[0]?.transform).toBeUndefined();
  });

  it("preserves DocumentController drawing preserved metadata maps", () => {
    const drawings = [
      {
        id: "1",
        zOrder: 0,
        anchor: { type: "cell", row: 0, col: 0 },
        kind: { type: "image", imageId: "img1" },
        size: { width: 10, height: 10 },
        preserved: {
          "xlsx.pic_xml": "<xdr:pic>...</xdr:pic>",
          // Malformed entries should be ignored (best-effort).
          other: 123,
        },
        unrelatedKey: { hello: "world" },
      },
    ];

    const ui = convertDocumentSheetDrawingsToUiDrawingObjects(drawings);
    expect(ui).toHaveLength(1);
    expect(ui[0]?.preserved).toEqual({ "xlsx.pic_xml": "<xdr:pic>...</xdr:pic>" });
  });

  it("ignores malformed DocumentController preserved payloads (best-effort)", () => {
    const drawings = [
      {
        id: "1",
        zOrder: 0,
        anchor: { type: "cell", row: 0, col: 0 },
        kind: { type: "image", imageId: "img1" },
        size: { width: 10, height: 10 },
        preserved: "<not-a-map>",
      },
    ];

    const ui = convertDocumentSheetDrawingsToUiDrawingObjects(drawings);
    expect(ui).toHaveLength(1);
    expect(ui[0]?.preserved).toBeUndefined();
  });

  it("preserves transform metadata even when falling back to the formula-model adapter", () => {
    // This object is in the formula-model/Rust JSON shape (externally-tagged enums),
    // but includes UI-authored `transform` metadata at the top-level. The DocumentController
    // adapter should keep that metadata rather than dropping it during conversion.
    const drawings = [
      {
        id: 1,
        kind: { Image: { image_id: "img1" } },
        anchor: {
          Absolute: {
            pos: { x_emu: 0, y_emu: 0 },
            ext: { cx: 10, cy: 10 },
          },
        },
        z_order: 0,
        preserved: { "xlsx.pic_xml": "<xdr:pic>...</xdr:pic>" },
        transform: { rotationDeg: 30, flipH: true, flipV: false },
      },
    ];

    const ui = convertDocumentSheetDrawingsToUiDrawingObjects(drawings);
    expect(ui).toHaveLength(1);
    expect(ui[0]?.kind).toEqual({ type: "image", imageId: "img1" });
    expect(ui[0]?.transform).toEqual({ rotationDeg: 30, flipH: true, flipV: false });
  });

  it("preserves preserved metadata even when falling back to the formula-model adapter", () => {
    const drawings = [
      {
        id: 1,
        kind: { Image: { image_id: "img1" } },
        anchor: {
          Absolute: {
            pos: { x_emu: 0, y_emu: 0 },
            ext: { cx: 10, cy: 10 },
          },
        },
        z_order: 0,
        preserved: {
          "xlsx.pic_xml": "<xdr:pic>...</xdr:pic>",
          other: 123,
        },
      },
    ];

    const ui = convertDocumentSheetDrawingsToUiDrawingObjects(drawings);
    expect(ui).toHaveLength(1);
    expect(ui[0]?.preserved).toEqual({ "xlsx.pic_xml": "<xdr:pic>...</xdr:pic>" });
  });

  it("accepts workbook snapshots with sheets encoded as an object map", () => {
    const workbook = {
      images: {
        images: {
          "image1.png": { bytes: [1], content_type: "image/png" },
        },
      },
      sheets: {
        Sheet1: {
          drawings: [
            {
              id: 1,
              kind: { Image: { image_id: "image1.png" } },
              anchor: {
                OneCell: {
                  from: { cell: { row: 0, col: 0 }, offset: { x_emu: 0, y_emu: 0 } },
                  ext: { cx: 10, cy: 10 },
                },
              },
              z_order: 0,
            },
          ],
        },
      },
    };

    const ui = convertModelWorkbookDrawingsToUiDrawingLayer(workbook);
    expect(ui.drawingsBySheetName.Sheet1).toHaveLength(1);
    expect(ui.drawingsBySheetName.Sheet1?.[0]?.kind).toEqual({ type: "image", imageId: "image1.png" });
  });

  it("converts DocumentController cell-anchored drawings (pixel size) into UI drawing objects", () => {
    const drawings = [
      {
        id: "d1",
        zOrder: 1,
        anchor: { type: "cell", sheetId: "Sheet1", row: 0, col: 0 },
        kind: { type: "image", imageId: "img1" },
        size: { width: 120, height: 80 },
      },
    ];

    const ui = convertDocumentSheetDrawingsToUiDrawingObjects(drawings, { sheetId: "Sheet1" });
    expect(ui).toHaveLength(1);
    expect(ui[0]?.kind).toEqual({ type: "image", imageId: "img1" });
    expect(ui[0]?.anchor).toEqual({
      type: "oneCell",
      from: { cell: { row: 0, col: 0 }, offset: { xEmu: 0, yEmu: 0 } },
      size: { cx: 120 * 9525, cy: 80 * 9525 },
    });
    expect(ui[0]?.zOrder).toBe(1);
    expect(ui[0]?.size).toEqual({ cx: 120 * 9525, cy: 80 * 9525 });
  });

  it("treats DocumentController chart kind with relId/chartId=unknown as an unknown (SmartArt) kind", () => {
    const rawXml = `<xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="2" name="SmartArt 1"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">
          <dgm:relIds r:dm="rId1" r:lo="rId2" r:qs="rId3" r:cs="rId4"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>`;

    const drawings = [
      {
        id: "1",
        zOrder: 0,
        anchor: { type: "cell", sheetId: "Sheet1", row: 0, col: 0 },
        kind: { type: "chart", chartId: "unknown", rawXml },
        size: { width: 10, height: 10 },
      },
    ];

    const ui = convertDocumentSheetDrawingsToUiDrawingObjects(drawings, { sheetId: "Sheet1" });
    expect(ui).toHaveLength(1);
    expect(ui[0]?.kind).toMatchObject({ type: "unknown", rawXml });
    expect(ui[0]?.kind.label).toBe("SmartArt 1");
  });

  it("converts DocumentController kind encodings paired with formula-model anchors (imported charts)", () => {
    const drawings = [
      {
        id: "7",
        zOrder: 0,
        // Formula-model/Rust anchor enum encoding (externally tagged).
        anchor: {
          Absolute: {
            pos: { x_emu: 0, y_emu: 0 },
            ext: { cx: 10, cy: 20 },
          },
        },
        // DocumentController-style kind encoding (internally tagged).
        kind: { type: "chart", chart_id: "Sheet1:7", raw_xml: "<xdr:graphicFrame/>" },
      },
    ];

    const ui = convertDocumentSheetDrawingsToUiDrawingObjects(drawings, { sheetId: "Sheet1" });
    expect(ui).toHaveLength(1);
    expect(ui[0]).toMatchObject({
      id: 7,
      zOrder: 0,
      anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: 10, cy: 20 } },
      kind: { type: "chart", chartId: "Sheet1:7", rawXml: "<xdr:graphicFrame/>" },
    });
  });
});
