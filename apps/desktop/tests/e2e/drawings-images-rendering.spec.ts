import { expect, test } from "@playwright/test";

import { readFileSync } from "node:fs";
import path from "node:path";
import { inflateRawSync } from "node:zlib";

import { gotoDesktop } from "./helpers";

type ZipEntry = {
  name: string;
  compressionMethod: number;
  compressedSize: number;
  uncompressedSize: number;
  localHeaderOffset: number;
};

function findEocd(buf: Buffer): number {
  // EOCD record is at least 22 bytes; the comment can make it larger.
  for (let i = buf.length - 22; i >= 0; i -= 1) {
    if (buf.readUInt32LE(i) === 0x06054b50) return i;
  }
  throw new Error("zip: end of central directory not found");
}

function parseZipEntries(buf: Buffer): Map<string, ZipEntry> {
  const eocd = findEocd(buf);
  const totalEntries = buf.readUInt16LE(eocd + 10);
  const centralDirOffset = buf.readUInt32LE(eocd + 16);

  const out = new Map<string, ZipEntry>();
  let offset = centralDirOffset;

  for (let i = 0; i < totalEntries; i += 1) {
    if (buf.readUInt32LE(offset) !== 0x02014b50) {
      throw new Error(`zip: invalid central directory signature at ${offset}`);
    }

    const compressionMethod = buf.readUInt16LE(offset + 10);
    const compressedSize = buf.readUInt32LE(offset + 20);
    const uncompressedSize = buf.readUInt32LE(offset + 24);
    const nameLen = buf.readUInt16LE(offset + 28);
    const extraLen = buf.readUInt16LE(offset + 30);
    const commentLen = buf.readUInt16LE(offset + 32);
    const localHeaderOffset = buf.readUInt32LE(offset + 42);
    const name = buf.toString("utf8", offset + 46, offset + 46 + nameLen);

    out.set(name, { name, compressionMethod, compressedSize, uncompressedSize, localHeaderOffset });
    offset += 46 + nameLen + extraLen + commentLen;
  }

  return out;
}

function readZipEntry(buf: Buffer, entry: ZipEntry): Buffer {
  const offset = entry.localHeaderOffset;
  if (buf.readUInt32LE(offset) !== 0x04034b50) {
    throw new Error(`zip: invalid local header signature at ${offset}`);
  }

  const nameLen = buf.readUInt16LE(offset + 26);
  const extraLen = buf.readUInt16LE(offset + 28);
  const dataOffset = offset + 30 + nameLen + extraLen;
  const compressed = buf.subarray(dataOffset, dataOffset + entry.compressedSize);

  if (entry.compressionMethod === 0) return Buffer.from(compressed);
  if (entry.compressionMethod === 8) return inflateRawSync(compressed);
  throw new Error(`zip: unsupported compression method ${entry.compressionMethod} for ${entry.name}`);
}

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
  test("renders floating DrawingML image.xlsx via DrawingOverlay canvas pixels", async ({ page }) => {
    const fixture = loadFloatingImageFixture();

    await gotoDesktop(page);

    // Ensure the built-in overlay container exists (charts + future drawings).
    await expect(page.locator(".chart-layer")).toHaveCount(1);

    const result = await page.evaluate(async ({ fixture }) => {
      const { DrawingOverlay, anchorToRectPx } = await import("/src/drawings/overlay.ts");

      const app = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp");

      const gridRoot = document.querySelector<HTMLElement>("#grid");
      if (!gridRoot) throw new Error("Missing #grid root");

      const dpr = window.devicePixelRatio || 1;
      const { width, height } = gridRoot.getBoundingClientRect();

      const existing = gridRoot.querySelector<HTMLCanvasElement>('[data-testid="e2e-drawing-overlay"]');
      existing?.remove();

      const canvas = document.createElement("canvas");
      canvas.dataset.testid = "e2e-drawing-overlay";
      canvas.style.position = "absolute";
      canvas.style.inset = "0";
      canvas.style.pointerEvents = "none";
      // Above chart overlay (z=2) / selection overlay (z=3) for debug readability.
      canvas.style.zIndex = "6";
      gridRoot.appendChild(canvas);

      const bytes = Uint8Array.from(atob(fixture.imagePngBase64), (c) => c.charCodeAt(0));
      const imageEntry = { id: fixture.imageId, bytes, mimeType: "image/png" };
      const backing = new Map<string, typeof imageEntry>();
      const images = {
        get(id: string) {
          return backing.get(id);
        },
        set(entry: typeof imageEntry) {
          backing.set(entry.id, entry);
        },
      };
      images.set(imageEntry);

      const geom = {
        cellOriginPx: (cell: { row: number; col: number }) => {
          const rect = app.getCellRect(cell);
          if (!rect) throw new Error(`Missing rect for cell r${cell.row}c${cell.col}`);
          return { x: rect.x, y: rect.y };
        },
        cellSizePx: (cell: { row: number; col: number }) => {
          const rect = app.getCellRect(cell);
          if (!rect) throw new Error(`Missing rect for cell r${cell.row}c${cell.col}`);
          return { width: rect.width, height: rect.height };
        },
      };

      const overlay = new DrawingOverlay(canvas, images, geom);
      const viewport = { scrollX: 0, scrollY: 0, width, height, dpr };
      overlay.resize(viewport);

      const obj = {
        id: 1,
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
      };

      await overlay.render([obj], viewport);

      const rect = anchorToRectPx(obj.anchor, geom);
      const sampleRect = (() => {
        const centerX = rect.x + rect.width / 2;
        const centerY = rect.y + rect.height / 2;
        const size = 20;
        const x = Math.max(0, Math.floor(centerX - size / 2));
        const y = Math.max(0, Math.floor(centerY - size / 2));
        const w = Math.min(Math.floor(size), Math.max(1, Math.floor(width) - x));
        const h = Math.min(Math.floor(size), Math.max(1, Math.floor(height) - y));
        return { x, y, width: w, height: h };
      })();

      const ctx = canvas.getContext("2d");
      if (!ctx) throw new Error("Missing 2d context");
      // `getImageData` uses device-pixel coordinates, not the CSS pixel coordinates
      // we render at after applying the DPR transform.
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

  test("renders images-in-cells image-in-cell.xlsx via DrawingOverlay canvas pixels", async ({ page }) => {
    const fixture = loadInCellImageFixture();

    await gotoDesktop(page);
    await expect(page.locator(".chart-layer")).toHaveCount(1);

    const result = await page.evaluate(async ({ fixture }) => {
      const { DrawingOverlay, anchorToRectPx } = await import("/src/drawings/overlay.ts");

      const app = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp");

      const gridRoot = document.querySelector<HTMLElement>("#grid");
      if (!gridRoot) throw new Error("Missing #grid root");

      const dpr = window.devicePixelRatio || 1;
      const { width, height } = gridRoot.getBoundingClientRect();

      const existing = gridRoot.querySelector<HTMLCanvasElement>('[data-testid="e2e-drawing-overlay"]');
      existing?.remove();

      const canvas = document.createElement("canvas");
      canvas.dataset.testid = "e2e-drawing-overlay";
      canvas.style.position = "absolute";
      canvas.style.inset = "0";
      canvas.style.pointerEvents = "none";
      canvas.style.zIndex = "6";
      gridRoot.appendChild(canvas);

      const bytes = Uint8Array.from(atob(fixture.imagePngBase64), (c) => c.charCodeAt(0));
      const imageEntry = { id: fixture.imageId, bytes, mimeType: "image/png" };
      const backing = new Map<string, typeof imageEntry>();
      const images = {
        get(id: string) {
          return backing.get(id);
        },
        set(entry: typeof imageEntry) {
          backing.set(entry.id, entry);
        },
      };
      images.set(imageEntry);

      const geom = {
        cellOriginPx: (cell: { row: number; col: number }) => {
          const rect = app.getCellRect(cell);
          if (!rect) throw new Error(`Missing rect for cell r${cell.row}c${cell.col}`);
          return { x: rect.x, y: rect.y };
        },
        cellSizePx: (cell: { row: number; col: number }) => {
          const rect = app.getCellRect(cell);
          if (!rect) throw new Error(`Missing rect for cell r${cell.row}c${cell.col}`);
          return { width: rect.width, height: rect.height };
        },
      };

      const overlay = new DrawingOverlay(canvas, images, geom);
      const viewport = { scrollX: 0, scrollY: 0, width, height, dpr };
      overlay.resize(viewport);

      // In-cell images should render inside the cell bounds. For the fixture we use Sheet1!A1.
      const obj = {
        id: 1,
        zOrder: 0,
        kind: { type: "image", imageId: fixture.imageId },
        anchor: {
          type: "twoCell",
          from: { cell: { row: 0, col: 0 }, offset: { xEmu: 0, yEmu: 0 } },
          // See DrawingOverlay.anchorToRectPx: `to` specifies the containing cell for the bottom-right corner,
          // so `row=1,col=1,offset=0` maps exactly to the bottom-right boundary of A1.
          to: { cell: { row: 1, col: 1 }, offset: { xEmu: 0, yEmu: 0 } },
        },
      };

      await overlay.render([obj], viewport);

      const rect = anchorToRectPx(obj.anchor, geom);
      const sampleRect = (() => {
        const centerX = rect.x + rect.width / 2;
        const centerY = rect.y + rect.height / 2;
        const size = 16;
        const x = Math.max(0, Math.floor(centerX - size / 2));
        const y = Math.max(0, Math.floor(centerY - size / 2));
        const w = Math.min(Math.floor(size), Math.max(1, Math.floor(width) - x));
        const h = Math.min(Math.floor(size), Math.max(1, Math.floor(height) - y));
        return { x, y, width: w, height: h };
      })();

      const ctx = canvas.getContext("2d");
      if (!ctx) throw new Error("Missing 2d context");
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
      `expected overlay canvas to contain non-transparent pixels near rendered in-cell image (sample=${JSON.stringify(
        result.sampleRect,
      )})`,
    ).toBeGreaterThan(0);
  });
});
