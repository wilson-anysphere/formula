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

    // Inject the fixture drawing, then assert we can observe non-transparent pixels in the
    // drawing canvas within the computed drawing bounds.
    //
    // NOTE: This relies on SpreadsheetApp's test/e2e debug hooks instead of reaching into
    // private overlay internals like `drawingOverlay.geom`.
    await page.evaluate(
      ({ fixture }) => {
        const app = (window as any).__formulaApp;
        if (!app) throw new Error("Missing window.__formulaApp");

        const doc = app.getDocument?.();
        if (!doc) throw new Error("Missing SpreadsheetApp.getDocument()");
        if (typeof doc.setSheetDrawings !== "function") {
          throw new Error("Missing DocumentController.setSheetDrawings()");
        }

        const sheetId = app.getCurrentSheetId?.();
        if (!sheetId) throw new Error("Missing SpreadsheetApp.getCurrentSheetId()");

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
      },
      { fixture },
    );

    await expect
      .poll(
        async () => {
          return await page.evaluate(() => {
            const app = (window as any).__formulaApp;
            if (!app) return false;
            if (typeof app.getDrawingsDebugState !== "function") return false;

            const state = app.getDrawingsDebugState();
            const drawings = Array.isArray(state?.drawings) ? state.drawings : [];
            const shape = drawings.find((d: any) => d?.kind === "shape");
            const rect = shape?.rectPx;
            if (!rect) return false;

            const canvas = document.querySelector<HTMLCanvasElement>('[data-testid="drawing-layer-canvas"]');
            if (!canvas) return false;
            const ctx = canvas.getContext("2d");
            if (!ctx) return false;

            const bounds = canvas.getBoundingClientRect();
            const viewportW = Math.max(1, Math.floor(bounds.width));
            const viewportH = Math.max(1, Math.floor(bounds.height));

            // Sample a small area near the expected text position (top + centered), avoiding
            // placeholder label rendering at the top-left corner.
            const centerX = rect.x + rect.width / 2;
            const centerY = rect.y + 12;
            const size = 20;
            const x = Math.max(0, Math.floor(centerX - size / 2));
            const y = Math.max(0, Math.floor(centerY - size / 2));
            const w = Math.min(Math.floor(size), Math.max(1, viewportW - x));
            const h = Math.min(Math.floor(size), Math.max(1, viewportH - y));

            const dpr = canvas.width / Math.max(1, bounds.width);
            const samplePx = {
              x: Math.max(0, Math.floor(x * dpr)),
              y: Math.max(0, Math.floor(y * dpr)),
              width: Math.max(1, Math.floor(w * dpr)),
              height: Math.max(1, Math.floor(h * dpr)),
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

            return nonTransparent > 0 && avgG > avgR + 20 && avgG > avgB + 20;
          });
        },
        { timeout: 20_000 },
      )
      .toBe(true);

    const result = await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      if (!app) return null;
      if (typeof app.getDrawingsDebugState !== "function") return null;

      const state = app.getDrawingsDebugState();
      const drawings = Array.isArray(state?.drawings) ? state.drawings : [];
      const shape = drawings.find((d: any) => d?.kind === "shape");
      const rect = shape?.rectPx;
      if (!rect) return null;

      const canvas = document.querySelector<HTMLCanvasElement>('[data-testid="drawing-layer-canvas"]');
      if (!canvas) return null;
      const ctx = canvas.getContext("2d");
      if (!ctx) return null;

      const bounds = canvas.getBoundingClientRect();
      const viewportW = Math.max(1, Math.floor(bounds.width));
      const viewportH = Math.max(1, Math.floor(bounds.height));
      const centerX = rect.x + rect.width / 2;
      const centerY = rect.y + 12;
      const size = 20;
      const x = Math.max(0, Math.floor(centerX - size / 2));
      const y = Math.max(0, Math.floor(centerY - size / 2));
      const w = Math.min(Math.floor(size), Math.max(1, viewportW - x));
      const h = Math.min(Math.floor(size), Math.max(1, viewportH - y));

      const dpr = canvas.width / Math.max(1, bounds.width);
      const samplePx = {
        x: Math.max(0, Math.floor(x * dpr)),
        y: Math.max(0, Math.floor(y * dpr)),
        width: Math.max(1, Math.floor(w * dpr)),
        height: Math.max(1, Math.floor(h * dpr)),
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

      return { nonTransparent, avgR, avgG, avgB };
    });

    expect(result).not.toBeNull();
    expect(result!.nonTransparent).toBeGreaterThan(0);
    // The fixture text is explicitly `srgbClr val="00FF00"` (green); ensure we aren't falling back
    // to default black placeholder label rendering.
    expect(result!.avgG).toBeGreaterThan(result!.avgR + 20);
    expect(result!.avgG).toBeGreaterThan(result!.avgB + 20);
  });
});
