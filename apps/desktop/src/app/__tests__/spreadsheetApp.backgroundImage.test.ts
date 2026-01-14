/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { existsSync, readFileSync } from "node:fs";
import path from "node:path";
import { inflateRawSync } from "node:zlib";

import type { ImageEntry } from "../../drawings/types";
import { SpreadsheetApp } from "../spreadsheetApp";
import { WorkbookSheetStore } from "../../sheets/workbookSheetStore";
import { hydrateSheetBackgroundImagesFromBackend } from "../../workbook/load/hydrateSheetBackgroundImages";

function resolveFixturePath(relativeFromRepoRoot: string): string {
  // Vitest modules can run with non-file `import.meta.url` in some environments. Resolve fixtures
  // from the working directory instead (walking upward until we find the repo root).
  let dir = process.cwd();
  for (let i = 0; i < 8; i += 1) {
    const candidate = path.resolve(dir, relativeFromRepoRoot);
    if (existsSync(candidate)) return candidate;
    const parent = path.dirname(dir);
    if (parent === dir) break;
    dir = parent;
  }
  throw new Error(`fixture not found: ${relativeFromRepoRoot}`);
}

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
  let createPatternSpy: ReturnType<typeof vi.fn>;
  type PatternToken = { kind: "pattern"; source: unknown; setTransform?: unknown };
  let createdPatterns: PatternToken[];
  let patternFillRects: Array<{
    canvasClassName: string;
    pattern: PatternToken;
    x: number;
    y: number;
    width: number;
    height: number;
  }>;

  function createMockCanvasContext(canvas: HTMLCanvasElement): CanvasRenderingContext2D {
    const noop = () => {};
    const gradient = { addColorStop: noop } as any;

    const state: { fillStyle: any } = { fillStyle: "#000000" };

    const ctxBase: any = {
      canvas,
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
        const fill = state.fillStyle as PatternToken | null;
        if (fill && typeof fill === "object" && (fill as any).kind === "pattern") {
          patternFillRects.push({ canvasClassName: canvas.className, pattern: fill, x, y, width, height });
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
    Object.defineProperty(window, "devicePixelRatio", { configurable: true, value: 1 });

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

    // CanvasGridRenderer uses CanvasPattern.setTransform when available to keep background patterns crisp
    // under HiDPI scaling. jsdom doesn't implement CanvasPattern, so stub it for unit tests.
    vi.stubGlobal(
      "CanvasPattern",
      class {
        setTransform() {}
      } as any,
    );

    createdPatterns = [];
    createPatternSpy = vi.fn((source: unknown) => {
      const token: PatternToken = { kind: "pattern", source, setTransform: vi.fn() };
      createdPatterns.push(token);
      return token as any;
    });
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
      const fixtureBytes = readFileSync(resolveFixturePath("fixtures/xlsx/basic/background-image.xlsx"));
      const imageEntry = parseSheetBackgroundImageFromXlsx(fixtureBytes);

      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);
      createPatternSpy.mockClear();
      createdPatterns = [];
      patternFillRects = [];

      const doc = app.getDocument();
      doc.setImage(imageEntry.id, { bytes: imageEntry.bytes, mimeType: imageEntry.mimeType });
      doc.setSheetBackgroundImageId(app.getCurrentSheetId(), imageEntry.id);
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

  it("trims worksheet background image ids when resolving background patterns", async () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "legacy";
    try {
      const fixtureBytes = readFileSync(resolveFixturePath("fixtures/xlsx/basic/background-image.xlsx"));
      const imageEntry = parseSheetBackgroundImageFromXlsx(fixtureBytes);

      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);
      createPatternSpy.mockClear();
      createdPatterns = [];
      patternFillRects = [];

      const doc = app.getDocument();
      doc.setImage(imageEntry.id, { bytes: imageEntry.bytes, mimeType: imageEntry.mimeType });
      // Simulate a sheet view update path that sets the background image id with incidental whitespace.
      doc.setSheetBackgroundImageId(app.getCurrentSheetId(), `  ${imageEntry.id}  `);
      await app.whenIdle();

      expect(app.getSheetBackgroundImageId(app.getCurrentSheetId())).toBe(imageEntry.id);
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
      Object.defineProperty(window, "devicePixelRatio", { configurable: true, value: 2 });
      const fixtureBytes = readFileSync(resolveFixturePath("fixtures/xlsx/basic/background-image.xlsx"));
      const imageEntry = parseSheetBackgroundImageFromXlsx(fixtureBytes);

      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);
      createPatternSpy.mockClear();
      createdPatterns = [];
      patternFillRects = [];

      const doc = app.getDocument();
      doc.setImage(imageEntry.id, { bytes: imageEntry.bytes, mimeType: imageEntry.mimeType });
      doc.setSheetBackgroundImageId(app.getCurrentSheetId(), imageEntry.id);
      await app.whenIdle();

      expect(createPatternSpy).toHaveBeenCalled();
      const basePatternRects = patternFillRects.filter((rect) => rect.canvasClassName.includes("grid-canvas--base"));
      expect(basePatternRects.length).toBeGreaterThan(0);
      // With DPR=2 and CanvasPattern.setTransform available, the shared-grid renderer should bake the
      // DPR into the offscreen tile resolution (but keep the CSS tile size consistent via pattern transforms).
      // Our test double exposes the tile canvas via `pattern.source`.
      expect(basePatternRects.some((rect) => (rect.pattern.source as any)?.width === 16)).toBe(true);
      const hiDpiPattern = createdPatterns.find((p) => (p.source as any)?.width === 16);
      expect((hiDpiPattern as any)?.setTransform).toBeDefined();
      expect(((hiDpiPattern as any).setTransform as any).mock?.calls?.length ?? 0).toBeGreaterThan(0);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("aborts + closes in-flight background decodes when image GC evicts the underlying bytes", async () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "shared";
    try {
      const fixtureBytes = readFileSync(resolveFixturePath("fixtures/xlsx/basic/background-image.xlsx"));
      const imageEntry = parseSheetBackgroundImageFromXlsx(fixtureBytes);

      const close = vi.fn();
      const decodedBitmap = { width: 8, height: 8, close } as any;

      let resolveDecode!: (value: any) => void;
      const inflightDecode = new Promise<any>((resolve) => {
        resolveDecode = resolve;
      });
      const createImageBitmapMock = vi.fn(() => inflightDecode);
      vi.stubGlobal("createImageBitmap", createImageBitmapMock as unknown as typeof createImageBitmap);

      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);
      createPatternSpy.mockClear();
      createdPatterns = [];
      patternFillRects = [];

      const doc = app.getDocument();
      doc.setImage(imageEntry.id, { bytes: imageEntry.bytes, mimeType: imageEntry.mimeType });
      doc.setSheetBackgroundImageId(app.getCurrentSheetId(), imageEntry.id);

      // Decode should have started, but not completed yet.
      expect(createImageBitmapMock).toHaveBeenCalledTimes(1);
      expect(createPatternSpy).not.toHaveBeenCalled();

      // Simulate the image bytes being garbage-collected (the GC path deletes bytes directly from
      // DocumentController caches without emitting a change event). Ensure the app aborts the in-flight
      // decode so the decoded bitmap is closed when it resolves.
      (app as any).workbookImageManager.setExternalImageIds([]);
      await app.runImageGcNow({ force: true });

      resolveDecode(decodedBitmap);
      // Flush internal `.then` handlers (decode completion + abort cleanup).
      await Promise.resolve();
      await Promise.resolve();
      await Promise.resolve();

      expect(close).toHaveBeenCalledTimes(1);
      // The pattern should never have been applied since the bytes were evicted before decode completion.
      expect(createPatternSpy).not.toHaveBeenCalled();
      expect(patternFillRects.length).toBe(0);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("hydrates worksheet background images from the desktop workbook backend during workbook load", async () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "legacy";
    try {
      const fixtureBytes = readFileSync(resolveFixturePath("fixtures/xlsx/basic/background-image.xlsx"));
      const imageEntry = parseSheetBackgroundImageFromXlsx(fixtureBytes);
      const imageId = imageEntry.id.split("/").pop() ?? imageEntry.id;

      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);
      createPatternSpy.mockClear();
      createdPatterns = [];
      patternFillRects = [];
      const doc = app.getDocument();
      // Workbook load flows call `restoreDocumentState()` which resets dirty tracking.
      // Simulate that here so we can assert hydration does not mark the document dirty.
      doc.markSaved();
      expect(doc.isDirty).toBe(false);

      const workbookSheetStore = new WorkbookSheetStore([{ id: "Sheet1", name: "Sheet1", visibility: "visible" }]);
      const bytesBase64 = Buffer.from(imageEntry.bytes).toString("base64");

      await hydrateSheetBackgroundImagesFromBackend({
        app,
        workbookSheetStore,
        backend: {
          async listImportedSheetBackgroundImages() {
            return [
              {
                sheet_name: "Sheet1",
                worksheet_part: "xl/worksheets/sheet1.xml",
                image_id: imageId,
                bytes_base64: bytesBase64,
                mime_type: imageEntry.mimeType,
              },
            ];
          },
        },
      });

      await app.whenIdle();

      expect(app.getSheetBackgroundImageId(app.getCurrentSheetId())).toBe(imageId);
      const workbookImageManager = (app as any).workbookImageManager as any;
      expect(Number(workbookImageManager?.imageRefCount?.get(imageId) ?? 0)).toBeGreaterThan(0);
      expect(doc.isDirty).toBe(false);
      expect(createPatternSpy).toHaveBeenCalled();
      expect(patternFillRects.filter((rect) => rect.canvasClassName.includes("grid-canvas--base")).length).toBeGreaterThan(0);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });

  it("hydrates worksheet background images from the desktop workbook backend in shared grid mode", async () => {
    const prior = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "shared";
    try {
      Object.defineProperty(window, "devicePixelRatio", { configurable: true, value: 2 });
      const fixtureBytes = readFileSync(resolveFixturePath("fixtures/xlsx/basic/background-image.xlsx"));
      const imageEntry = parseSheetBackgroundImageFromXlsx(fixtureBytes);
      const imageId = imageEntry.id.split("/").pop() ?? imageEntry.id;

      const root = createRoot();
      const status = {
        activeCell: document.createElement("div"),
        selectionRange: document.createElement("div"),
        activeValue: document.createElement("div"),
      };

      const app = new SpreadsheetApp(root, status);
      createPatternSpy.mockClear();
      createdPatterns = [];
      patternFillRects = [];
      const doc = app.getDocument();
      doc.markSaved();
      expect(doc.isDirty).toBe(false);

      const workbookSheetStore = new WorkbookSheetStore([{ id: "Sheet1", name: "Sheet1", visibility: "visible" }]);
      const bytesBase64 = Buffer.from(imageEntry.bytes).toString("base64");

      await hydrateSheetBackgroundImagesFromBackend({
        app,
        workbookSheetStore,
        backend: {
          async listImportedSheetBackgroundImages() {
            return [
              {
                sheet_name: "Sheet1",
                worksheet_part: "xl/worksheets/sheet1.xml",
                image_id: imageId,
                bytes_base64: bytesBase64,
                mime_type: imageEntry.mimeType,
              },
            ];
          },
        },
      });

      await app.whenIdle();

      expect(app.getSheetBackgroundImageId(app.getCurrentSheetId())).toBe(imageId);
      const workbookImageManager = (app as any).workbookImageManager as any;
      expect(Number(workbookImageManager?.imageRefCount?.get(imageId) ?? 0)).toBeGreaterThan(0);
      expect(doc.isDirty).toBe(false);
      expect(createPatternSpy).toHaveBeenCalled();
      const basePatternRects = patternFillRects.filter((rect) => rect.canvasClassName.includes("grid-canvas--base"));
      expect(basePatternRects.length).toBeGreaterThan(0);

      // With DPR=2 and CanvasPattern.setTransform available, the shared-grid renderer should bake the
      // DPR into the offscreen tile resolution.
      expect(basePatternRects.some((rect) => (rect.pattern.source as any)?.width === 16)).toBe(true);
      const hiDpiPattern = createdPatterns.find((p) => (p.source as any)?.width === 16);
      expect((hiDpiPattern as any)?.setTransform).toBeDefined();
      expect(((hiDpiPattern as any).setTransform as any).mock?.calls?.length ?? 0).toBeGreaterThan(0);

      app.destroy();
      root.remove();
    } finally {
      if (prior === undefined) delete process.env.DESKTOP_GRID_MODE;
      else process.env.DESKTOP_GRID_MODE = prior;
    }
  });
});
