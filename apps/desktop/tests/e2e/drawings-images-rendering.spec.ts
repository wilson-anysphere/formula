import { expect, test } from "@playwright/test";

import { readFileSync } from "node:fs";
import path from "node:path";

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

function resolveZipTarget(basePath: string, target: string): string {
  // Target paths in XLSX relationships use forward slashes (posix style).
  const baseDir = path.posix.dirname(basePath);
  return path.posix.normalize(path.posix.join(baseDir, target));
}

function loadFloatingImageFixture(): {
  imageId: string;
  imagePngBase64: string;
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
} {
  const fixtureUrl = new URL("../../../../fixtures/xlsx/basic/image.xlsx", import.meta.url);
  const bytes = readFileSync(fixtureUrl);
  const entries = parseZipEntries(bytes);

  const drawingPath = "xl/drawings/drawing1.xml";
  const drawingRelsPath = "xl/drawings/_rels/drawing1.xml.rels";

  const drawingXml = readZipEntry(bytes, entries.get(drawingPath)!).toString("utf8");
  const drawingRelsXml = readZipEntry(bytes, entries.get(drawingRelsPath)!).toString("utf8");

  const fromBlock = extractFirstMatch(drawingXml, /<xdr:from>([\s\S]*?)<\/xdr:from>/, "xdr:from block");
  const toBlock = extractFirstMatch(drawingXml, /<xdr:to>([\s\S]*?)<\/xdr:to>/, "xdr:to block");
  const embedRid = extractFirstMatch(drawingXml, /<a:blip\b[^>]*\br:embed="([^"]+)"/, "a:blip r:embed");

  const target = extractFirstMatch(
    drawingRelsXml,
    new RegExp(`<Relationship\\b[^>]*\\bId="${embedRid}"[^>]*\\bTarget="([^"]+)"`, "i"),
    `Relationship Target for ${embedRid}`,
  );
  const imageZipPath = resolveZipTarget(drawingPath, target);
  const imageEntry = entries.get(imageZipPath);
  if (!imageEntry) {
    throw new Error(`fixture image.xlsx missing image part resolved from rels: ${imageZipPath}`);
  }
  const imageBytes = readZipEntry(bytes, imageEntry);
  const imageId = path.posix.basename(imageZipPath);

  return {
    imageId,
    imagePngBase64: imageBytes.toString("base64"),
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

function loadInCellImageFixture(): { imageId: string; imagePngBase64: string } {
  const fixtureUrl = new URL("../../../../fixtures/xlsx/images-in-cells/image-in-cell.xlsx", import.meta.url);
  const bytes = readFileSync(fixtureUrl);
  const entries = parseZipEntries(bytes);

  const sheetPath = "xl/worksheets/sheet1.xml";
  const sheetXml = readZipEntry(bytes, entries.get(sheetPath)!).toString("utf8");
  if (!/<c r="A1"[^>]*\bvm="1"/.test(sheetXml)) {
    throw new Error(`fixture image-in-cell.xlsx missing expected Sheet1!A1 vm="1" cell binding`);
  }

  const imageZipPath = "xl/media/image1.png";
  const imageEntry = entries.get(imageZipPath);
  if (!imageEntry) throw new Error(`fixture image-in-cell.xlsx missing image part: ${imageZipPath}`);
  const imageBytes = readZipEntry(bytes, imageEntry);

  return { imageId: "image1.png", imagePngBase64: imageBytes.toString("base64") };
}

test.describe("drawing + image rendering regressions", () => {
  test("renders floating DrawingML image.xlsx via the desktop drawing layer canvas", async ({ page }) => {
    const fixture = loadFloatingImageFixture();

    await gotoDesktop(page);

    // Ensure the built-in overlay canvases exist (charts + drawings).
    await expect(page.locator("canvas.grid-canvas--chart")).toHaveCount(1);
    await expect(page.getByTestId("drawing-layer-canvas")).toHaveCount(1);

    // Regression guard: overlay stacking should be deterministic across modes.
    // Drawings/images must render above charts and cell content, but below selection UI.
    const z = await page.evaluate(() => {
      const drawing = document.querySelector(".drawing-layer");
      const chart = document.querySelector(".grid-canvas--chart");
      const selection = document.querySelector(".grid-canvas--selection");
      if (!drawing || !chart || !selection) return null;
      return {
        drawing: getComputedStyle(drawing).zIndex,
        chart: getComputedStyle(chart).zIndex,
        selection: getComputedStyle(selection).zIndex,
      };
    });
    expect(z).not.toBeNull();
    expect(Number(z!.chart)).toBeLessThan(Number(z!.drawing));
    expect(Number(z!.drawing)).toBeLessThan(Number(z!.selection));

    const result = await page.evaluate(async ({ fixture }) => {
      const { anchorToRectPx } = await import("/src/drawings/overlay.ts");

      const app = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp");

      const overlay = (app as any).drawingOverlay;
      if (!overlay) throw new Error("Missing SpreadsheetApp.drawingOverlay");
      const canvas = document.querySelector<HTMLCanvasElement>('[data-testid="drawing-layer-canvas"]');
      if (!canvas) throw new Error("Missing drawing-layer-canvas element");

      const bytes = Uint8Array.from(atob(fixture.imagePngBase64), (c) => c.charCodeAt(0));
      const doc = app.getDocument?.();
      if (!doc) throw new Error("Missing SpreadsheetApp.getDocument()");
      if (typeof doc.setImage !== "function") {
        throw new Error("Missing DocumentController.setImage()");
      }
      doc.setImage(fixture.imageId, { bytes, mimeType: "image/png" });

      if (typeof doc.setSheetDrawings !== "function") {
        throw new Error("Missing DocumentController.setSheetDrawings()");
      }

      // Capture the *app-driven* render pass (DocumentController change -> SpreadsheetApp.renderDrawings -> DrawingOverlay.render).
      // This is more realistic than calling `overlay.render(...)` directly and will fail if the integration
      // between document drawings and the overlay regresses.
      const origRender = overlay.render.bind(overlay);
      (overlay as any).render = (objects: any, viewport: any, ...rest: any[]) => {
        const promise = origRender(objects, viewport, ...rest);
        (window as any).__testLastDrawingOverlayRender = { promise, objects, viewport };
        return promise;
      };
      (window as any).__testLastDrawingOverlayRender = null;

      const sheetId = app.getCurrentSheetId?.();
      if (!sheetId) throw new Error("Missing SpreadsheetApp.getCurrentSheetId()");

      // Store a formula-model style drawing object (JSON-serializable) and let SpreadsheetApp convert it
      // through `convertModelWorksheetDrawingsToUiDrawingObjects(...)`.
      doc.setSheetDrawings(sheetId, [
        {
          id: "1",
          zOrder: 0,
          kind: { type: "image", imageId: fixture.imageId },
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

      const renderedHasImage = Array.isArray(last.objects)
        ? last.objects.some((o) => o?.kind?.type === "image" && o?.kind?.imageId === fixture.imageId)
        : false;
      if (!renderedHasImage) {
        throw new Error(`DrawingOverlay.render was invoked, but the rendered object list did not include imageId=${fixture.imageId}`);
      }

      await last.promise;

      const geom = (overlay as any).geom;
      if (!geom) throw new Error("Missing DrawingOverlay.geom");
      const obj = (last.objects as any[]).find((o) => o?.kind?.type === "image" && o?.kind?.imageId === fixture.imageId);
      if (!obj) throw new Error("Expected rendered image object to exist");
      const viewport = last.viewport;
      if (!viewport) throw new Error("Missing drawing overlay viewport from render()");
      const headerOffsetX = typeof viewport.headerOffsetX === "number" && Number.isFinite(viewport.headerOffsetX) ? viewport.headerOffsetX : 0;
      const headerOffsetY = typeof viewport.headerOffsetY === "number" && Number.isFinite(viewport.headerOffsetY) ? viewport.headerOffsetY : 0;
      const rect = anchorToRectPx(obj.anchor, geom);
      const sampleRect = (() => {
        const centerX = rect.x + rect.width / 2 - viewport.scrollX + headerOffsetX;
        const centerY = rect.y + rect.height / 2 - viewport.scrollY + headerOffsetY;
        const size = 20;
        const x = Math.max(0, Math.floor(centerX - size / 2));
        const y = Math.max(0, Math.floor(centerY - size / 2));
        const w = Math.min(Math.floor(size), Math.max(1, Math.floor(viewport.width) - x));
        const h = Math.min(Math.floor(size), Math.max(1, Math.floor(viewport.height) - y));
        return { x, y, width: w, height: h };
      })();

      const ctx = canvas.getContext("2d");
      if (!ctx) throw new Error("Missing 2d context");
      // `getImageData` uses device-pixel coordinates, not the CSS pixel coordinates
      // we render at after applying the DPR transform.
      const dpr = typeof viewport.dpr === "number" && viewport.dpr > 0 ? viewport.dpr : window.devicePixelRatio || 1;
      const samplePx = {
        x: Math.max(0, Math.floor(sampleRect.x * dpr)),
        y: Math.max(0, Math.floor(sampleRect.y * dpr)),
        width: Math.max(1, Math.floor(sampleRect.width * dpr)),
        height: Math.max(1, Math.floor(sampleRect.height * dpr)),
      };
      const imageData = ctx.getImageData(samplePx.x, samplePx.y, samplePx.width, samplePx.height);
      let nonTransparent = 0;
      for (let i = 3; i < imageData.data.length; i += 4) {
        if (imageData.data[i] !== 0) nonTransparent += 1;
      }

      return { nonTransparent, sampleRect };
    }, { fixture });

    expect(
      result.nonTransparent,
      `expected overlay canvas to contain non-transparent pixels near rendered floating image (sample=${JSON.stringify(
        result.sampleRect,
      )})`,
    ).toBeGreaterThan(0);
  });

  test("renders images-in-cells image-in-cell.xlsx as an in-cell image (shared grid)", async ({ page }) => {
    const fixture = loadInCellImageFixture();

    await gotoDesktop(page, "/?grid=shared");
    await expect(page.locator("canvas.grid-canvas--chart")).toHaveCount(1);
    await expect(page.getByTestId("drawing-layer-canvas")).toHaveCount(1);

    await page.evaluate(async ({ fixture }) => {
      const app = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp");

      const sharedGrid = (app as any).sharedGrid;
      if (!sharedGrid) throw new Error("Expected shared grid mode to be enabled");
      const renderer = sharedGrid.renderer;
      if (!renderer) throw new Error("Missing shared grid renderer");

      // Store the fixture's `xl/media/image1.png` bytes in the DocumentController image store so the
      // normal shared-grid imageResolver path is exercised (DocumentController -> Blob -> ImageBitmap).
      const bytes = Uint8Array.from(atob(fixture.imagePngBase64), (c) => c.charCodeAt(0));
      const doc = app.getDocument();
      if (!doc || typeof doc.setImage !== "function") {
        throw new Error("Missing DocumentController.setImage");
      }
      doc.setImage(fixture.imageId, { bytes, mimeType: "image/png" });
      if (typeof renderer.clearImageCache === "function") {
        renderer.clearImageCache();
      }

      const sheetId = app.getCurrentSheetId();
      // Use the formula-model envelope shape (`{type:"image", value:{...}}`) so the DocumentCellProvider
      // image detection logic is exercised.
      // Keep alt text empty so if image rendering regresses and falls back to drawing `cell.value` as text,
      // we still have a low-coverage glyph ("[Image]") that is easy to distinguish from a fully-painted image.
      doc.setCellValue(sheetId, "A1", { type: "image", value: { imageId: fixture.imageId, altText: "" } });

      // Force an immediate render so the grid requests the image bitmap.
      renderer.markAllDirty?.();
      renderer.renderImmediately?.();
    }, { fixture });

    await expect
      .poll(async () => {
        return await page.evaluate((imageId) => {
          const app = (window as any).__formulaApp;
          const renderer = app?.sharedGrid?.renderer;
          const cache: Map<string, any> | undefined = (renderer as any)?.imageBitmapCache;
          const entry = cache?.get?.(imageId);
          return entry?.state ?? null;
        }, fixture.imageId);
      })
      .toBe("ready");

    const result = await page.evaluate((imageId) => {
      const app = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp");
      const rect = app.getCellRectA1("A1");
      if (!rect) throw new Error("Missing A1 rect");

      const canvas = document.querySelector<HTMLCanvasElement>("canvas.grid-canvas--content");
      if (!canvas) throw new Error("Missing grid content canvas");
      const ctx = canvas.getContext("2d");
      if (!ctx) throw new Error("Missing content canvas 2d context");

      // Ensure the image has been painted in the current frame.
      app.sharedGrid?.renderer?.renderImmediately?.();

      const renderer = app.sharedGrid?.renderer;
      if (!renderer) throw new Error("Missing shared grid renderer");

      const cache: Map<string, any> | undefined = (renderer as any)?.imageBitmapCache;
      const cached = cache?.get?.(imageId);
      const state = cached?.state ?? null;
      const bitmap = cached?.bitmap ?? null;

      const zoom = typeof renderer.getZoom === "function" ? renderer.getZoom() : 1;
      const paddingX = 4 * zoom;
      const paddingY = 2 * zoom;

      const availableWidth = Math.max(0, rect.width - paddingX * 2);
      const availableHeight = Math.max(0, rect.height - paddingY * 2);

      const srcW =
        bitmap && typeof bitmap.width === "number" && Number.isFinite(bitmap.width) && bitmap.width > 0 ? bitmap.width : 1;
      const srcH =
        bitmap && typeof bitmap.height === "number" && Number.isFinite(bitmap.height) && bitmap.height > 0 ? bitmap.height : 1;

      const scale =
        availableWidth > 0 && availableHeight > 0 && srcW > 0 && srcH > 0
          ? Math.min(availableWidth / srcW, availableHeight / srcH)
          : 1;

      const destW = srcW * scale;
      const destH = srcH * scale;
      const destX = rect.x + paddingX + Math.max(0, (availableWidth - destW) / 2);
      const destY = rect.y + paddingY + Math.max(0, (availableHeight - destH) / 2);

      // Sample inside the expected image destination rect (not the full cell) so we fail if the renderer
      // silently falls back to drawing plain text and leaves the image region transparent.
      const sampleRect = (() => {
        const x = Math.max(0, Math.floor(destX));
        const y = Math.max(0, Math.floor(destY));
        const w = Math.max(1, Math.floor(destW));
        const h = Math.max(1, Math.floor(destH));
        return { x, y, width: w, height: h };
      })();

      const dpr = canvas.width / Math.max(1, canvas.getBoundingClientRect().width);

      const samplePx = {
        x: Math.max(0, Math.floor(sampleRect.x * dpr)),
        y: Math.max(0, Math.floor(sampleRect.y * dpr)),
        width: Math.max(1, Math.floor(sampleRect.width * dpr)),
        height: Math.max(1, Math.floor(sampleRect.height * dpr)),
      };
      const imageData = ctx.getImageData(samplePx.x, samplePx.y, samplePx.width, samplePx.height);
      let nonTransparent = 0;
      for (let i = 3; i < imageData.data.length; i += 4) {
        if (imageData.data[i] !== 0) nonTransparent += 1;
      }
      const coverage = imageData.data.length > 0 ? nonTransparent / (imageData.data.length / 4) : 0;

      return { nonTransparent, coverage, sampleRect, state, zoom, srcW, srcH, rect };
    }, fixture.imageId);

    expect(result.state).toBe("ready");

    expect(
      result.coverage,
      `expected shared-grid content canvas to paint the in-cell image region (coverage=${result.coverage}, zoom=${result.zoom}, src=${result.srcW}x${result.srcH}, cellRect=${JSON.stringify(result.rect)}, sample=${JSON.stringify(result.sampleRect)})`,
    ).toBeGreaterThan(0.6);
  });
});
