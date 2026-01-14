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
    // Ensure the grid root is present before we inject our test overlay canvas.
    await page.waitForSelector("#grid");

    const result = await page.evaluate(async ({ fixture }) => {
      const { DrawingOverlay, anchorToRectPx } = await import("/src/drawings/overlay.ts");

      const app = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp");

      const gridRoot = document.querySelector<HTMLElement>("#grid");
      if (!gridRoot) throw new Error("Missing #grid root");

      const { width, height } = gridRoot.getBoundingClientRect();
      const viewport = { scrollX: 0, scrollY: 0, width, height, dpr: window.devicePixelRatio || 1 };

      const existing = gridRoot.querySelector<HTMLCanvasElement>('[data-testid="e2e-drawing-overlay-shapes"]');
      existing?.remove();

      const canvas = document.createElement("canvas");
      canvas.dataset.testid = "e2e-drawing-overlay-shapes";
      canvas.style.position = "absolute";
      canvas.style.inset = "0";
      canvas.style.pointerEvents = "none";
      canvas.style.zIndex = "6";
      gridRoot.appendChild(canvas);

      const images = {
        get() {
          return undefined;
        },
        set() {},
      };

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

      const overlay = new DrawingOverlay(canvas, images as any, geom);
      overlay.resize(viewport);

      const obj = {
        id: 1,
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
      };

      await overlay.render([obj], viewport);

      const rect = anchorToRectPx(obj.anchor, geom);
      const sampleRect = (() => {
        const x = Math.max(0, Math.floor(rect.x));
        const y = Math.max(0, Math.floor(rect.y));
        const w = Math.min(Math.floor(rect.width), Math.max(1, Math.floor(width) - x));
        const h = Math.min(Math.floor(rect.height), Math.max(1, Math.floor(height) - y));
        return { x, y, width: w, height: h };
      })();

      const ctx = canvas.getContext("2d");
      if (!ctx) throw new Error("Missing 2d context");
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

      return { nonTransparent, sampleRect };
    }, { fixture });

    expect(
      result.nonTransparent,
      `expected overlay canvas to contain non-transparent pixels inside rendered shape bounds (sample=${JSON.stringify(
        result.sampleRect,
      )})`,
    ).toBeGreaterThan(0);
  });
});
