import { expect, test } from "@playwright/test";

import { readFileSync } from "node:fs";

import { gotoDesktop } from "./helpers";
import { parseZipEntries, readZipEntry } from "./zip";

function extractInt(xml: string, tag: string): number {
  const re = new RegExp(`<${tag}>(\\d+)<\\/${tag}>`);
  const match = xml.match(re);
  if (!match) throw new Error(`missing <${tag}> in: ${xml.slice(0, 200)}…`);
  return Number.parseInt(match[1]!, 10);
}

function extractFirstMatch(xml: string, re: RegExp, label: string): string {
  const match = xml.match(re);
  if (!match) throw new Error(`missing ${label} in: ${xml.slice(0, 200)}…`);
  return match[1]!;
}

function loadShapeTextFixture(): {
  anchor: {
    fromCol: number;
    fromColOff: number;
    fromRow: number;
    fromRowOff: number;
    toCol: number;
    toColOff: number;
    toRow: number;
    toRowOff: number;
  };
  shapeXml: string;
} {
  const fixtureUrl = new URL("../../../../fixtures/xlsx/basic/shape-textbox.xlsx", import.meta.url);
  const bytes = readFileSync(fixtureUrl);
  const entries = parseZipEntries(bytes);

  const drawingPath = "xl/drawings/drawing1.xml";
  const drawingEntry = entries.get(drawingPath);
  if (!drawingEntry) throw new Error(`fixture shape-textbox.xlsx missing part: ${drawingPath}`);

  const drawingXml = readZipEntry(bytes, drawingEntry).toString("utf8");
  const fromBlock = extractFirstMatch(drawingXml, /<xdr:from>([\s\S]*?)<\/xdr:from>/, "xdr:from block");
  const toBlock = extractFirstMatch(drawingXml, /<xdr:to>([\s\S]*?)<\/xdr:to>/, "xdr:to block");
  const shapeXml = extractFirstMatch(drawingXml, /(<xdr:sp>[\s\S]*?<\/xdr:sp>)/, "xdr:sp block");

  return {
    shapeXml,
    anchor: {
      fromCol: extractInt(fromBlock, "xdr:col"),
      fromColOff: extractInt(fromBlock, "xdr:colOff"),
      fromRow: extractInt(fromBlock, "xdr:row"),
      fromRowOff: extractInt(fromBlock, "xdr:rowOff"),
      toCol: extractInt(toBlock, "xdr:col"),
      toColOff: extractInt(toBlock, "xdr:colOff"),
      toRow: extractInt(toBlock, "xdr:row"),
      toRowOff: extractInt(toBlock, "xdr:rowOff"),
    },
  };
}

test.describe("drawing shape text rendering regressions", () => {
  test("renders DrawingML txBody text from shape-textbox.xlsx via DrawingOverlay canvas pixels", async ({ page }) => {
    const fixture = loadShapeTextFixture();

    await gotoDesktop(page);
    // Ensure the built-in drawing overlay canvas is present before we inject drawings.
    await page.waitForSelector("#grid");
    await page.waitForSelector('[data-testid="drawing-layer-canvas"]');

    const result = await page.evaluate(async ({ fixture }) => {
      const { anchorToRectPx } = await import("/src/drawings/overlay.ts");

      const app = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp");

      const overlay = (app as any).drawingOverlay;
      if (!overlay) throw new Error("Missing SpreadsheetApp.drawingOverlay");

      const doc = app.getDocument?.();
      if (!doc) throw new Error("Missing SpreadsheetApp.getDocument()");
      if (typeof doc.setSheetDrawings !== "function") {
        throw new Error("Missing DocumentController.setSheetDrawings()");
      }

      const sheetId = app.getCurrentSheetId?.();
      if (!sheetId) throw new Error("Missing SpreadsheetApp.getCurrentSheetId()");

      const canvas = document.querySelector<HTMLCanvasElement>('[data-testid="drawing-layer-canvas"]');
      if (!canvas) throw new Error("Missing drawing-layer-canvas element");

      // Capture the app-driven render pass triggered by `setSheetDrawings`.
      const origRender = overlay.render.bind(overlay);
      (overlay as any).render = (objects: any, viewport: any, ...rest: any[]) => {
        const promise = origRender(objects, viewport, ...rest);
        (window as any).__testLastDrawingOverlayRender = { promise, objects, viewport };
        return promise;
      };
      (window as any).__testLastDrawingOverlayRender = null;

      // Store a formula-model style drawing object and let SpreadsheetApp convert it.
      doc.setSheetDrawings(sheetId, [
        {
          id: "1",
          zOrder: 0,
          kind: { type: "shape", raw_xml: fixture.shapeXml },
          anchor: {
            type: "twoCell",
            from: {
              cell: { row: fixture.anchor.fromRow, col: fixture.anchor.fromCol },
              offset: { xEmu: fixture.anchor.fromColOff, yEmu: fixture.anchor.fromRowOff },
            },
            to: {
              cell: { row: fixture.anchor.toRow, col: fixture.anchor.toCol },
              offset: { xEmu: fixture.anchor.toColOff, yEmu: fixture.anchor.toRowOff },
            },
          },
        },
      ]);

      const last = (window as any).__testLastDrawingOverlayRender as
        | { promise: Promise<void>; objects: any[]; viewport: any }
        | null;
      if (!last) throw new Error("Expected SpreadsheetApp to invoke DrawingOverlay.render after setSheetDrawings");

      const renderedHasShape = Array.isArray(last.objects) ? last.objects.some((o) => o?.kind?.type === "shape") : false;
      if (!renderedHasShape) {
        throw new Error("DrawingOverlay.render was invoked, but the rendered object list did not include a shape");
      }

      await last.promise;

      const geom = (overlay as any).geom;
      if (!geom) throw new Error("Missing DrawingOverlay.geom");

      const obj = (last.objects as any[]).find((o) => o?.kind?.type === "shape");
      if (!obj) throw new Error("Expected rendered shape object to exist");

      const viewport = last.viewport;
      if (!viewport) throw new Error("Missing drawing overlay viewport from render()");

      const rect = anchorToRectPx(obj.anchor, geom);
      // Sample a small area near the expected text position (top + centered), avoiding
      // placeholder label rendering at the top-left corner.
      const sampleRect = (() => {
        const centerX = rect.x + rect.width / 2 - viewport.scrollX;
        const centerY = rect.y + 12 - viewport.scrollY;
        const size = 20;
        const x = Math.max(0, Math.floor(centerX - size / 2));
        const y = Math.max(0, Math.floor(centerY - size / 2));
        const w = Math.min(Math.floor(size), Math.max(1, Math.floor(viewport.width) - x));
        const h = Math.min(Math.floor(size), Math.max(1, Math.floor(viewport.height) - y));
        return { x, y, width: w, height: h };
      })();

      const ctx = canvas.getContext("2d");
      if (!ctx) throw new Error("Missing 2d context");
      const dpr = typeof viewport.dpr === "number" && viewport.dpr > 0 ? viewport.dpr : window.devicePixelRatio || 1;
      const samplePx = {
        x: Math.max(0, Math.floor(sampleRect.x * dpr)),
        y: Math.max(0, Math.floor(sampleRect.y * dpr)),
        width: Math.max(1, Math.floor(sampleRect.width * dpr)),
        height: Math.max(1, Math.floor(sampleRect.height * dpr)),
      };
      const imageData = ctx.getImageData(samplePx.x, samplePx.y, samplePx.width, samplePx.height);
      let nonTransparent = 0;
      let sumR = 0;
      let sumG = 0;
      let sumB = 0;
      for (let i = 0; i < imageData.data.length; i += 4) {
        const a = imageData.data[i + 3] ?? 0;
        if (a === 0) continue;
        nonTransparent += 1;
        sumR += imageData.data[i] ?? 0;
        sumG += imageData.data[i + 1] ?? 0;
        sumB += imageData.data[i + 2] ?? 0;
      }

      const avgR = nonTransparent > 0 ? sumR / nonTransparent : 0;
      const avgG = nonTransparent > 0 ? sumG / nonTransparent : 0;
      const avgB = nonTransparent > 0 ? sumB / nonTransparent : 0;

      return { nonTransparent, avgR, avgG, avgB, sampleRect };
    }, { fixture });

    expect(
      result.nonTransparent,
      `expected overlay canvas to contain non-transparent pixels inside rendered shape bounds (sample=${JSON.stringify(
        result.sampleRect,
      )})`,
    ).toBeGreaterThan(0);

    // The fixture text is explicitly `srgbClr val="00FF00"` (green); ensure we aren't falling back
    // to default black placeholder label rendering.
    expect(result.avgG).toBeGreaterThan(result.avgR + 20);
    expect(result.avgG).toBeGreaterThan(result.avgB + 20);
  });
});
