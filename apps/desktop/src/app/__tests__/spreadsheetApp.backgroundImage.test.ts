/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { readFileSync } from "node:fs";
import { inflateRawSync } from "node:zlib";

import type { ImageEntry } from "../../drawings/types";
import { SpreadsheetApp } from "../spreadsheetApp";

function createInMemoryLocalStorage(): Storage {
  const store = new Map<string, string>();
  return {
    getItem: (key: string) => (store.has(key) ? store.get(key)! : null),
    setItem: (key: string, value: string) => {
      store.set(String(key), String(value));
    },
    removeItem: (key: string) => {
      store.delete(String(key));
    },
    clear: () => {
      store.clear();
    },
    key: (index: number) => Array.from(store.keys())[index] ?? null,
    get length() {
      return store.size;
    },
  } as Storage;
}

function createRoot(): HTMLElement {
  const root = document.createElement("div");
  root.tabIndex = 0;
  root.getBoundingClientRect = () =>
    ({
      width: 800,
      height: 600,
      left: 0,
      top: 0,
      right: 800,
      bottom: 600,
      x: 0,
      y: 0,
      toJSON: () => {},
    }) as any;
  document.body.appendChild(root);
  return root;
}

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

function normalizeZipPath(path: string): string {
  const parts = path.split("/").filter((p) => p.length > 0);
  const out: string[] = [];
  for (const part of parts) {
    if (part === ".") continue;
    if (part === "..") out.pop();
    else out.push(part);
  }
  return out.join("/");
}

function guessMimeType(path: string): string {
  const lower = path.toLowerCase();
  if (lower.endsWith(".png")) return "image/png";
  if (lower.endsWith(".jpg") || lower.endsWith(".jpeg")) return "image/jpeg";
  if (lower.endsWith(".gif")) return "image/gif";
  if (lower.endsWith(".bmp")) return "image/bmp";
  return "application/octet-stream";
}

function parseSheetBackgroundImageFromXlsx(bytes: Buffer): ImageEntry {
  const entries = parseZipEntries(bytes);
  const sheetPath = "xl/worksheets/sheet1.xml";
  const relsPath = "xl/worksheets/_rels/sheet1.xml.rels";
  const sheetEntry = entries.get(sheetPath);
  const relsEntry = entries.get(relsPath);
  if (!sheetEntry) throw new Error(`missing ${sheetPath} in fixture`);
  if (!relsEntry) throw new Error(`missing ${relsPath} in fixture`);

  const sheetXml = readZipEntry(bytes, sheetEntry).toString("utf8");
  const relId = sheetXml.match(/<picture[^>]*\br:id="([^"]+)"/)?.[1];
  if (!relId) throw new Error("fixture sheet1.xml missing <picture r:id=...>");

  const relsXml = readZipEntry(bytes, relsEntry).toString("utf8");
  let target: string | null = null;
  for (const match of relsXml.matchAll(/<Relationship\b([^>]*)\/?>/g)) {
    const attrs = match[1] ?? "";
    const id = attrs.match(/\bId="([^"]+)"/)?.[1];
    if (id !== relId) continue;
    target = attrs.match(/\bTarget="([^"]+)"/)?.[1] ?? null;
    break;
  }
  if (!target) throw new Error(`fixture sheet rels missing target for ${relId}`);

  const baseDir = sheetPath.slice(0, sheetPath.lastIndexOf("/") + 1);
  const resolved = normalizeZipPath(`${baseDir}${target}`);
  const imageEntry = entries.get(resolved);
  if (!imageEntry) throw new Error(`fixture missing resolved image part ${resolved}`);

  const imageBytes = readZipEntry(bytes, imageEntry);
  return { id: resolved, bytes: new Uint8Array(imageBytes), mimeType: guessMimeType(resolved) };
}

describe("SpreadsheetApp worksheet background images", () => {
  const patternToken = { kind: "pattern" };

  let createPatternSpy: ReturnType<typeof vi.fn>;
  let patternFillRects: Array<{ canvasClassName: string; x: number; y: number; width: number; height: number }>;

  function createMockCanvasContext(canvas: HTMLCanvasElement): CanvasRenderingContext2D {
    const noop = () => {};
    const gradient = { addColorStop: noop } as any;

    const state: { fillStyle: any } = { fillStyle: "#000000" };

    const ctxBase: any = {
      canvas,
      fillStyle: state.fillStyle,
      set fillStyle(value: any) {
        state.fillStyle = value;
      },
      get fillStyle() {
        return state.fillStyle;
      },
      measureText: (text: string) => ({ width: text.length * 8 }),
      createLinearGradient: () => gradient,
      createPattern: createPatternSpy,
      fillRect: (x: number, y: number, width: number, height: number) => {
        if (state.fillStyle === patternToken) {
          patternFillRects.push({ canvasClassName: canvas.className, x, y, width, height });
        }
      },
      getImageData: () => ({ data: new Uint8ClampedArray(), width: 0, height: 0 }),
      putImageData: noop,
      drawImage: noop,
    };

    return new Proxy(ctxBase, {
      get(target, prop) {
        if (prop in target) return (target as any)[prop];
        return noop;
      },
      set(target, prop, value) {
        (target as any)[prop] = value;
        return true;
      },
    }) as any;
  }

  beforeEach(() => {
    document.body.innerHTML = "";

    const storage = createInMemoryLocalStorage();
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
    Object.defineProperty(window, "localStorage", { configurable: true, value: storage });
    storage.clear();

    Object.defineProperty(globalThis, "requestAnimationFrame", {
      configurable: true,
      value: (cb: FrameRequestCallback) => {
        cb(0);
        return 0;
      },
    });
    Object.defineProperty(globalThis, "cancelAnimationFrame", { configurable: true, value: () => {} });

    // createImageBitmap is used to decode image bytes; jsdom doesn't implement it.
    Object.defineProperty(globalThis, "createImageBitmap", {
      configurable: true,
      value: vi.fn(async () => {
        const canvas = document.createElement("canvas");
        canvas.width = 8;
        canvas.height = 8;
        return canvas as any;
      }),
    });

    createPatternSpy = vi.fn(() => patternToken as any);
    patternFillRects = [];

    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: function () {
        return createMockCanvasContext(this);
      },
    });

    (globalThis as any).ResizeObserver = class {
      observe() {}
      disconnect() {}
    };
  });

  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it("fills the legacy grid body with a repeated background pattern when the active sheet has a background image", async () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "legacy";
    try {
      const fixtureUrl = new URL("../../../../../fixtures/xlsx/basic/background-image.xlsx", import.meta.url);
      const fixtureBytes = readFileSync(fixtureUrl);
      const imageEntry = parseSheetBackgroundImageFromXlsx(fixtureBytes);

      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);
      createPatternSpy.mockClear();
      patternFillRects = [];

      app.setWorkbookImages([imageEntry]);
      app.setSheetBackgroundImageId(app.getCurrentSheetId(), imageEntry.id);
      await app.whenIdle();

      expect(createPatternSpy).toHaveBeenCalled();
      expect(patternFillRects.filter((rect) => rect.canvasClassName.includes("grid-canvas--base")).length).toBeGreaterThan(0);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("fills the shared grid background layer with a repeated background pattern when the active sheet has a background image", async () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "shared";
    try {
      const fixtureUrl = new URL("../../../../../fixtures/xlsx/basic/background-image.xlsx", import.meta.url);
      const fixtureBytes = readFileSync(fixtureUrl);
      const imageEntry = parseSheetBackgroundImageFromXlsx(fixtureBytes);

      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);
      createPatternSpy.mockClear();
      patternFillRects = [];

      app.setWorkbookImages([imageEntry]);
      app.setSheetBackgroundImageId(app.getCurrentSheetId(), imageEntry.id);
      await app.whenIdle();

      expect(createPatternSpy).toHaveBeenCalled();
      expect(patternFillRects.filter((rect) => rect.canvasClassName.includes("grid-canvas--base")).length).toBeGreaterThan(0);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });
});
