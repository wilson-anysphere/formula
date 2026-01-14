/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { DocumentController } from "../../document/documentController.js";
import { mergeEmbeddedCellImagesIntoSnapshot } from "../../workbook/load/embeddedCellImages.js";
import { SpreadsheetApp } from "../spreadsheetApp";

let priorGridMode: string | undefined;

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

function createMockCanvasContext(): CanvasRenderingContext2D {
  const noop = () => {};
  const gradient = { addColorStop: noop } as any;
  const context = new Proxy(
    {
      canvas: document.createElement("canvas"),
      measureText: (text: string) => ({ width: text.length * 8 }),
      createLinearGradient: () => gradient,
      createPattern: () => null,
      getImageData: () => ({ data: new Uint8ClampedArray(), width: 0, height: 0 }),
      putImageData: noop,
    },
    {
      get(target, prop) {
        if (prop in target) return (target as any)[prop];
        // Default all unknown properties to no-op functions so rendering code can execute.
        return noop;
      },
      set(target, prop, value) {
        (target as any)[prop] = value;
        return true;
      },
    },
  );
  return context as any;
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

describe("SpreadsheetApp shared grid (in-cell images)", () => {
  afterEach(() => {
    if (priorGridMode === undefined) delete process.env.DESKTOP_GRID_MODE;
    else process.env.DESKTOP_GRID_MODE = priorGridMode;
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  beforeEach(() => {
    priorGridMode = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "shared";
    document.body.innerHTML = "";

    const storage = createInMemoryLocalStorage();
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
    Object.defineProperty(window, "localStorage", { configurable: true, value: storage });
    storage.clear();

    // CanvasGridRenderer schedules renders via requestAnimationFrame; ensure it exists in jsdom.
    Object.defineProperty(globalThis, "requestAnimationFrame", {
      configurable: true,
      value: (cb: FrameRequestCallback) => {
        cb(0);
        return 0;
      },
    });
    Object.defineProperty(globalThis, "cancelAnimationFrame", { configurable: true, value: () => {} });

    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: () => createMockCanvasContext(),
    });

    (globalThis as any).ResizeObserver = class {
      observe() {}
      disconnect() {}
    };
  });

  it("hydrates embedded cell images from an XLSX snapshot and attempts to resolve them during render", async () => {
    // The shared-grid renderer resolves images via SpreadsheetApp's `sharedGridImageResolver`.
    // That resolver prefers `DocumentController.getImageBlob()` when available (so CanvasGridRenderer
    // can use its `<img>`-based decode fallback when `createImageBitmap(blob)` fails for some valid
    // PNGs in headless Chromium), and falls back to `getImage()` (raw bytes) otherwise.
    const blobSpy = vi.spyOn(DocumentController.prototype as any, "getImageBlob");
    const bytesSpy = vi.spyOn(DocumentController.prototype as any, "getImage");

    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);

    blobSpy.mockClear();
    bytesSpy.mockClear();

    // Fixture: `fixtures/xlsx/images-in-cells/image-in-cell.xlsx`
    // - Sheet1!A1: Place-in-Cell image
    // - Sheet1!B1: IMAGE() formula that also resolves to an in-cell image via RichData `vm=`
    //
    // Both ultimately reference `xl/media/image1.png`, which we normalize to `image1.png`.
    const baseSheet = {
      id: "Sheet1",
      name: "Sheet1",
      visibility: "visible",
      cells: [] as any[],
    };

    const imageBytesBase64 =
      // 1×1 PNG (opaque black) to keep the snapshot self-contained.
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGNgYAAAAAMAAWgmWQ0AAAAASUVORK5CYII=";

    const { images } = mergeEmbeddedCellImagesIntoSnapshot({
      sheets: [baseSheet],
      embeddedCellImages: [
        {
          worksheet_part: "xl/worksheets/sheet1.xml",
          sheet_name: "Sheet1",
          row: 0,
          col: 0,
          image_id: "image1.png",
          bytes_base64: imageBytesBase64,
          mime_type: "image/png",
          alt_text: null,
        },
        {
          worksheet_part: "xl/worksheets/sheet1.xml",
          sheet_name: "Sheet1",
          row: 0,
          col: 1,
          image_id: "image1.png",
          bytes_base64: imageBytesBase64,
          mime_type: "image/png",
          alt_text: null,
        },
      ],
      resolveSheetIdByName: (name) => (name === "Sheet1" ? "Sheet1" : null),
      sheetIdsInOrder: ["Sheet1"],
      maxRows: 10_000,
      maxCols: 200,
    });

    const snapshotBytes = new TextEncoder().encode(
      JSON.stringify({
        schemaVersion: 1,
        sheets: [baseSheet],
        images,
      }),
    );

    await app.restoreDocumentState(snapshotBytes);

    const cellA1 = app.getDocument().getCell("Sheet1", "A1");
    expect(cellA1.value).toMatchObject({ type: "image", value: { imageId: "image1.png" } });

    const calls = [...blobSpy.mock.calls.map((args) => args[0]), ...bytesSpy.mock.calls.map((args) => args[0])];
    expect(calls).toContain("image1.png");

    app.destroy();
    root.remove();
  });

  it("treats in-cell image payloads as non-text input while still producing a safe display string", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);

    const baseSheet = {
      id: "Sheet1",
      name: "Sheet1",
      visibility: "visible",
      cells: [] as any[],
    };

    const imageBytesBase64 =
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGNgYAAAAAMAAWgmWQ0AAAAASUVORK5CYII=";

    const { images } = mergeEmbeddedCellImagesIntoSnapshot({
      sheets: [baseSheet],
      embeddedCellImages: [
        {
          worksheet_part: "xl/worksheets/sheet1.xml",
          sheet_name: "Sheet1",
          row: 0,
          col: 0,
          image_id: "image1.png",
          bytes_base64: imageBytesBase64,
          mime_type: "image/png",
          alt_text: null,
        },
      ],
      resolveSheetIdByName: (name) => (name === "Sheet1" ? "Sheet1" : null),
      sheetIdsInOrder: ["Sheet1"],
      maxRows: 10_000,
      maxCols: 200,
    });

    const snapshotBytes = new TextEncoder().encode(
      JSON.stringify({
        schemaVersion: 1,
        sheets: [baseSheet],
        images,
      }),
    );

    await app.restoreDocumentState(snapshotBytes);

    // The image value should render as a stable display placeholder (not "[object Object]").
    expect(app.getCellComputedValueForSheet("Sheet1", { row: 0, col: 0 })).toBe("[Image]");
    await expect(app.getCellDisplayTextForRenderA1("A1")).resolves.toBe("[Image]");

    // Editing should start from an empty input string and committing that empty draft should be a no-op
    // (must not clear or overwrite the image payload).
    (app as any).applyEdit("Sheet1", { row: 0, col: 0 }, "");
    const cell = app.getDocument().getCell("Sheet1", "A1");
    expect(cell.value).toMatchObject({ type: "image", value: { imageId: "image1.png" } });

    app.destroy();
    root.remove();
  });
});
