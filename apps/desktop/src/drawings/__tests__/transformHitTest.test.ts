import { describe, expect, it } from "vitest";

import { buildHitTestIndex, hitTestDrawings } from "../hitTest";
import { pxToEmu, type GridGeometry, type Viewport } from "../overlay";
import { parseDrawingTransformFromRawXml } from "../transform";
import type { DrawingObject } from "../types";

describe("drawings transform parsing", () => {
  it("extracts rotation + flips from a:xfrm", () => {
    const rawXml = `
      <xdr:sp xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <xdr:spPr>
          <a:xfrm rot="5400000" flipH="1">
            <a:off x="0" y="0"/>
            <a:ext cx="100" cy="200"/>
          </a:xfrm>
        </xdr:spPr>
      </xdr:sp>
    `;

    expect(parseDrawingTransformFromRawXml(rawXml)).toEqual({
      rotationDeg: 90,
      flipH: true,
      flipV: false,
    });
  });

  it("parses single-quoted attributes (common in some OOXML payloads)", () => {
    const rawXml = `
      <xdr:sp xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <xdr:spPr>
          <a:xfrm rot='-5400000' flipV='true'>
            <a:off x='0' y='0'/>
            <a:ext cx='100' cy='200'/>
          </a:xfrm>
        </xdr:spPr>
      </xdr:sp>
    `;

    expect(parseDrawingTransformFromRawXml(rawXml)).toEqual({
      rotationDeg: -90,
      flipH: false,
      flipV: true,
    });
  });

  it('treats "on"/"off" and "yes"/"no" as boolean values', () => {
    const rawXml = `
      <xdr:sp xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <xdr:spPr>
          <a:xfrm rot="0" flipH="off" flipV="on">
            <a:off x="0" y="0"/>
            <a:ext cx="100" cy="200"/>
          </a:xfrm>
        </xdr:spPr>
      </xdr:sp>
    `;

    expect(parseDrawingTransformFromRawXml(rawXml)).toEqual({
      rotationDeg: 0,
      flipH: false,
      flipV: true,
    });
  });
});

describe("drawings hit testing with rotation", () => {
  const geom: GridGeometry = {
    cellOriginPx: () => ({ x: 0, y: 0 }),
    cellSizePx: () => ({ width: 0, height: 0 }),
  };

  const viewport: Viewport = { scrollX: 0, scrollY: 0, width: 500, height: 500, dpr: 1 };

  it("hits a rotated rect and excludes points only inside the original AABB", () => {
    const obj: DrawingObject = {
      id: 1,
      kind: { type: "shape" },
      anchor: {
        type: "absolute",
        pos: { xEmu: pxToEmu(100), yEmu: pxToEmu(100) },
        size: { cx: pxToEmu(100), cy: pxToEmu(50) },
      },
      zOrder: 0,
      transform: { rotationDeg: 90, flipH: false, flipV: false },
    };

    const index = buildHitTestIndex([obj], geom, { bucketSizePx: 64 });

    // Inside the rotated rect but outside the unrotated AABB (y < 100).
    expect(hitTestDrawings(index, viewport, 150, 80, geom)?.object.id).toBe(1);

    // Inside the unrotated AABB but outside the rotated rect.
    expect(hitTestDrawings(index, viewport, 110, 110, geom)).toBe(null);
  });

  it("still hits rotated objects when the hit-test index zoom differs (fallback scan)", () => {
    const obj: DrawingObject = {
      id: 1,
      kind: { type: "shape" },
      anchor: {
        type: "absolute",
        pos: { xEmu: pxToEmu(100), yEmu: pxToEmu(100) },
        size: { cx: pxToEmu(100), cy: pxToEmu(50) },
      },
      zOrder: 0,
      transform: { rotationDeg: 90, flipH: false, flipV: false },
    };

    // Build index at zoom=1, but hit test with zoom=2 to exercise the fallback code path.
    const index = buildHitTestIndex([obj], geom, { bucketSizePx: 64 });
    const zoomedViewport: Viewport = { ...viewport, zoom: 2 };

    // Scale the prior "inside rotated, outside unrotated" point by 2x.
    expect(hitTestDrawings(index, zoomedViewport, 300, 160, geom)?.object.id).toBe(1);
  });
});
