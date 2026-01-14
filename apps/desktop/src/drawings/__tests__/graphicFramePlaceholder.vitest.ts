import { describe, expect, it } from "vitest";

import { DrawingOverlay, pxToEmu, type GridGeometry, type Viewport } from "../overlay";
import { graphicFramePlaceholderLabel, isGraphicFrame } from "../shapeRenderer";
import { convertModelDrawingObjectToUiDrawingObject } from "../modelAdapters";
import type { DrawingObject, ImageStore } from "../types";

// Extracted from `fixtures/xlsx/basic/smartart.xlsx` -> `xl/drawings/drawing1.xml`.
const SMARTART_GRAPHIC_FRAME_XML = `<xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="2" name="SmartArt 1"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm>
        <a:off x="0" y="0"/>
        <a:ext cx="0" cy="0"/>
      </xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">
          <dgm:relIds r:dm="rId1" r:lo="rId2" r:qs="rId3" r:cs="rId4"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>`;

// Full anchor payload for the same object (also from `drawing1.xml`). The model
// can preserve either the `graphicFrame` subtree or the containing anchor.
const SMARTART_TWO_CELL_ANCHOR_XML = `<xdr:twoCellAnchor>
    <xdr:from>
      <xdr:col>1</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>1</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>6</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>10</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:to>
    ${SMARTART_GRAPHIC_FRAME_XML}
    <xdr:clientData/>
  </xdr:twoCellAnchor>`;

function createStubCanvasContext(): {
  ctx: CanvasRenderingContext2D;
  calls: Array<{ method: string; args: unknown[] }>;
  state: { strokeStyle?: unknown };
} {
  const calls: Array<{ method: string; args: unknown[] }> = [];
  const state: { strokeStyle?: unknown } = {};

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
    set strokeStyle(value: unknown) {
      state.strokeStyle = value;
    },
    get strokeStyle() {
      return state.strokeStyle;
    },
  };

  return { ctx: ctx as CanvasRenderingContext2D, calls, state };
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

const viewport: Viewport = { scrollX: 0, scrollY: 0, width: 100, height: 100, dpr: 1 };

function createUnknownObject(kind: DrawingObject["kind"]): DrawingObject {
  return {
    id: 1,
    kind,
    anchor: {
      type: "absolute",
      pos: { xEmu: pxToEmu(5), yEmu: pxToEmu(7) },
      size: { cx: pxToEmu(20), cy: pxToEmu(10) },
    },
    zOrder: 0,
  };
}

describe("DrawingML graphicFrame detection", () => {
  it("detects xdr:graphicFrame in SmartArt drawing XML", () => {
    expect(isGraphicFrame(SMARTART_GRAPHIC_FRAME_XML)).toBe(true);
    expect(graphicFramePlaceholderLabel(SMARTART_GRAPHIC_FRAME_XML)).toBe("SmartArt");
  });

  it("detects xdr:graphicFrame when preserved as a full anchor subtree", () => {
    expect(isGraphicFrame(SMARTART_TWO_CELL_ANCHOR_XML)).toBe(true);
    expect(graphicFramePlaceholderLabel(SMARTART_TWO_CELL_ANCHOR_XML)).toBe("SmartArt");
  });
});

describe("DrawingOverlay graphicFrame placeholders", () => {
  it("renders SmartArt placeholder label when no label is provided", async () => {
    const { ctx, calls, state } = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);
    const overlay = new DrawingOverlay(canvas, images, geom);

    await overlay.render(
      [
        createUnknownObject({
          type: "unknown",
          rawXml: SMARTART_GRAPHIC_FRAME_XML,
        }),
      ],
      viewport,
    );

    // `resolveCssVar` falls back to "magenta" in the Node test environment.
    expect(state.strokeStyle).toBe("magenta");
    expect(calls.some((call) => call.method === "strokeRect")).toBe(true);
    expect(
      calls.some((call) => call.method === "fillText" && call.args[0] === "SmartArt"),
    ).toBe(true);
  });

  it("uses the object's label when available", async () => {
    const { ctx, calls } = createStubCanvasContext();
    const canvas = createStubCanvas(ctx);
    const overlay = new DrawingOverlay(canvas, images, geom);

    await overlay.render(
      [
        createUnknownObject({
          type: "unknown",
          rawXml: SMARTART_GRAPHIC_FRAME_XML,
          label: "SmartArt 1",
        }),
      ],
      viewport,
    );

    expect(
      calls.some((call) => call.method === "fillText" && call.args[0] === "SmartArt 1"),
    ).toBe(true);
  });
});

describe("drawings/modelAdapters SmartArt", () => {
  it("converts non-chart graphicFrames (SmartArt) into an unknown kind with a name label", () => {
    const model = {
      id: 2,
      kind: {
        ChartPlaceholder: {
          rel_id: "unknown",
          raw_xml: SMARTART_GRAPHIC_FRAME_XML,
        },
      },
      anchor: {
        Absolute: {
          pos: { x_emu: 0, y_emu: 0 },
          ext: { cx: 10, cy: 20 },
        },
      },
      z_order: 0,
    };

    const ui = convertModelDrawingObjectToUiDrawingObject(model);
    expect(ui.kind.type).toBe("unknown");
    expect(ui.kind).toMatchObject({
      rawXml: SMARTART_GRAPHIC_FRAME_XML,
      label: "SmartArt 1",
    });
  });
});
