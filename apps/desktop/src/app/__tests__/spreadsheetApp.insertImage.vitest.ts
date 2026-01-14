/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { pxToEmu } from "../../drawings/overlay";
import { MAX_INSERT_IMAGE_BYTES } from "../../drawings/insertImageLimits.js";
import * as ui from "../../extensions/ui.js";
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

async function waitFor(condition: () => boolean, timeoutMs: number = 1000): Promise<void> {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    if (condition()) return;
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
  throw new Error("Timed out waiting for condition");
}

describe("SpreadsheetApp insert image (floating drawing)", () => {
  afterEach(() => {
    vi.useRealTimers();
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
    delete process.env.DESKTOP_GRID_MODE;
  });

  beforeEach(() => {
    document.body.innerHTML = "";
    process.env.DESKTOP_GRID_MODE = "legacy";

    const storage = createInMemoryLocalStorage();
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
    Object.defineProperty(window, "localStorage", { configurable: true, value: storage });
    storage.clear();

    Object.defineProperty(globalThis, "requestAnimationFrame", {
      configurable: true,
      writable: true,
      value: (cb: FrameRequestCallback) => {
        cb(0);
        return 0;
      },
    });
    Object.defineProperty(globalThis, "cancelAnimationFrame", { configurable: true, writable: true, value: () => {} });

    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: () => createMockCanvasContext(),
    });

    (globalThis as any).ResizeObserver = class {
      observe() {}
      disconnect() {}
    };

    Object.defineProperty(globalThis, "createImageBitmap", {
      configurable: true,
      value: vi.fn(async () => ({})),
    });
  });

  it("stores image bytes and creates a drawing anchored at the active cell", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const sheetId = app.getCurrentSheetId();

    app.activateCell({ row: 3, col: 4 }, { scrollIntoView: false, focus: false });
    app.insertImageFromLocalFile();

    const input = root.querySelector<HTMLInputElement>('input[data-testid="insert-image-input"]');
    expect(input).not.toBeNull();

    const bytes = new Uint8Array([1, 2, 3, 4, 5, 6]);
    const file = new File([bytes], "test.png", { type: "image/png" });
    Object.defineProperty(input!, "files", { configurable: true, value: [file] });
    input!.dispatchEvent(new Event("change", { bubbles: true }));

    await waitFor(() => {
      const drawings = (app.getDocument() as any)?.getSheetDrawings?.(sheetId) ?? [];
      return Array.isArray(drawings) && drawings.length === 1;
    });

    const drawings = (app.getDocument() as any)?.getSheetDrawings?.(sheetId) ?? [];
    expect(drawings).toHaveLength(1);
    const obj = drawings[0]!;

    expect(obj?.kind?.type).toBe("image");
    expect(obj?.anchor).toEqual({
      type: "oneCell",
      from: { cell: { row: 3, col: 4 }, offset: { xEmu: 0, yEmu: 0 } },
      size: { cx: pxToEmu(200), cy: pxToEmu(150) },
    });

    const entry = app.getDrawingImages().get(obj.kind.imageId);
    expect(entry).toBeTruthy();
    expect(entry.bytes).toEqual(bytes);

    // Drawing image bytes are stored out-of-band (in-memory + IndexedDB) rather than in
    // DocumentController snapshots / image map.
    const stored = (app.getDocument() as any)?.getImage?.(obj.kind.imageId);
    expect(stored).toBeFalsy();
    const snapshotBytes = app.getDocument().encodeState();
    const snapshot = JSON.parse(new TextDecoder().decode(snapshotBytes));
    expect(snapshot.images).toBeUndefined();

    // Insert should also select the new drawing (used by overlay handles + split view).
    expect(String(app.getSelectedDrawingId())).toBe(String(obj.id));

    app.destroy();
    root.remove();
  });

  it("blocks insert image in read-only mode with a throttled insert-pictures toast", () => {
    document.body.innerHTML = `<div id="toast-root"></div>`;

    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    // `isReadOnly()` consults the collab session when present.
    (app as any).collabSession = { isReadOnly: () => true };

    app.insertImageFromLocalFile();
    app.insertImageFromLocalFile();

    // Should not open/attach a file input when read-only.
    const input = root.querySelector<HTMLInputElement>('input[data-testid="insert-image-input"]');
    expect(input).toBeNull();

    // Throttled toast should only appear once (avoid spam on repeated commands).
    expect(document.querySelectorAll('[data-testid="toast"]')).toHaveLength(1);
    expect(document.querySelector("#toast-root")?.textContent ?? "").toContain("insert pictures");

    // Cleanup toast to avoid leaving timers running.
    (document.querySelector<HTMLElement>('[data-testid="toast"]') as any)?.click?.();

    app.destroy();
    root.remove();
  });

  it("does not enable drawing interactions in shared-grid mode when inserting via file picker", async () => {
    process.env.DESKTOP_GRID_MODE = "shared";

    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, { enableDrawingInteractions: false });
    const sheetId = app.getCurrentSheetId();

    app.activateCell({ row: 1, col: 2 }, { scrollIntoView: false, focus: false });
    app.insertImageFromLocalFile();

    const input = root.querySelector<HTMLInputElement>('input[data-testid="insert-image-input"]');
    expect(input).not.toBeNull();

    const bytes = new Uint8Array([1, 2, 3, 4, 5, 6]);
    const file = new File([bytes], "test.png", { type: "image/png" });
    Object.defineProperty(input!, "files", { configurable: true, value: [file] });
    input!.dispatchEvent(new Event("change", { bubbles: true }));

    await waitFor(() => {
      const drawings = (app.getDocument() as any)?.getSheetDrawings?.(sheetId) ?? [];
      return Array.isArray(drawings) && drawings.length === 1;
    });

    const drawings = (app.getDocument() as any)?.getSheetDrawings?.(sheetId) ?? [];
    expect(drawings).toHaveLength(1);
    const obj = drawings[0]!;

    expect(String(app.getSelectedDrawingId())).toBe(String(obj.id));

    expect((app as any).drawingInteractionController).toBeNull();

    app.destroy();
    root.remove();
  });

  it("inserts into the original sheet even if the user switches sheets mid-insert", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const focusSpy = vi.spyOn(app, "focus");
    const doc: any = app.getDocument();
    const sheet1 = app.getCurrentSheetId();

    // Ensure Sheet2 exists so we can switch away while bytes are still loading.
    doc.setCellValue("Sheet2", { row: 0, col: 0 }, "X");

    let resolveBytes: ((buf: ArrayBuffer) => void) | null = null;
    const arrayBuffer = new Promise<ArrayBuffer>((resolve) => {
      resolveBytes = resolve;
    });
    const file = {
      name: "slow.png",
      type: "image/png",
      size: 4,
      arrayBuffer: () => arrayBuffer,
    } as any as File;

    app.insertImageFromLocalFile();
    const input = root.querySelector<HTMLInputElement>('input[data-testid="insert-image-input"]');
    expect(input).not.toBeNull();
    Object.defineProperty(input!, "files", { configurable: true, value: [file] });
    input!.dispatchEvent(new Event("change", { bubbles: true }));

    // Switch sheets while `insertImageFromPickedFile` is awaiting the file bytes.
    app.activateSheet("Sheet2");
    expect(app.getCurrentSheetId()).toBe("Sheet2");
    // Sheet switches shouldn't cause async insert completion to steal focus away from the
    // new sheet (e.g. formula bar editing). Clear any incidental focus calls before
    // the file bytes resolve.
    focusSpy.mockClear();

    resolveBytes?.(new Uint8Array([1, 2, 3, 4]).buffer);

    await waitFor(() => {
      const drawings = doc.getSheetDrawings?.(sheet1) ?? [];
      return Array.isArray(drawings) && drawings.length === 1;
    });

    const sheet1Drawings = Array.isArray(doc.getSheetDrawings(sheet1)) ? doc.getSheetDrawings(sheet1) : [];
    expect(sheet1Drawings).toHaveLength(1);
    expect(Array.isArray(doc.getSheetDrawings("Sheet2")) ? doc.getSheetDrawings("Sheet2") : []).toHaveLength(0);

    // Inserting on a non-active sheet should not disrupt the current sheet's drawing selection.
    const state = app.getDrawingsDebugState();
    expect(state.sheetId).toBe("Sheet2");
    expect(state.drawings).toHaveLength(0);
    expect(state.selectedId).toBe(null);
    expect(focusSpy).not.toHaveBeenCalled();

    app.destroy();
    root.remove();
  });

  it("does not merge unrelated edits into the image insertion undo step while bytes are loading", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const sheetId = app.getCurrentSheetId();
    const doc = app.getDocument();
    const editedCell = { row: 0, col: 2 };
    const initialCellValue = doc.getCell(sheetId, editedCell).value;

    const bytes = new Uint8Array([1, 2, 3, 4]);

    let resolveBytes: ((buffer: ArrayBuffer) => void) | null = null;
    const bytesPromise = new Promise<ArrayBuffer>((resolve) => {
      resolveBytes = resolve;
    });

    const file = {
      name: "slow.png",
      type: "image/png",
      size: bytes.byteLength,
      arrayBuffer: vi.fn(async () => bytesPromise),
    } as any as File;

    const insertPromise = (app as any).insertImageFromPickedFile(file);

    // While the image bytes are still loading, perform a normal cell edit.
    doc.setCellInput(sheetId, editedCell, "hello", { label: "Edit Cell" });

    resolveBytes!(bytes.buffer);
    await insertPromise;

    expect((doc as any).getSheetDrawings(sheetId)).toHaveLength(1);
    expect(doc.getCell(sheetId, editedCell).value).toBe("hello");

    // Undo should first remove the inserted drawing, leaving the cell edit intact.
    doc.undo();
    expect((doc as any).getSheetDrawings(sheetId)).toHaveLength(0);
    expect(doc.getCell(sheetId, editedCell).value).toBe("hello");

    // Undo again should revert the cell edit.
    doc.undo();
    expect(doc.getCell(sheetId, editedCell).value).toEqual(initialCellValue);

    app.destroy();
    root.remove();
  });

  it("rejects PNGs with extremely large dimensions without inserting a drawing", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const toastSpy = vi.spyOn(ui, "showToast").mockImplementation(() => {});

    const app = new SpreadsheetApp(root, status);
    const sheetId = app.getCurrentSheetId();

    // Construct a minimal PNG header with an oversized IHDR width.
    const png = new Uint8Array(24);
    png.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 0);
    png.set([0x49, 0x48, 0x44, 0x52], 12); // IHDR
    // width=10001, height=1 (big-endian)
    png.set([0x00, 0x00, 0x27, 0x11], 16);
    png.set([0x00, 0x00, 0x00, 0x01], 20);

    const file = new File([png], "huge.png", { type: "image/png" });

    const images = app.getDrawingImages();
    const setSpy = vi.spyOn(images, "set");

    await (app as any).insertImageFromPickedFile(file);

    expect(toastSpy).toHaveBeenCalledWith(
      "Image dimensions too large (10001x1). Choose a smaller image.",
      "warning",
    );

    expect(((app.getDocument() as any).getSheetDrawings?.(sheetId) ?? [])).toHaveLength(0);
    expect(setSpy).not.toHaveBeenCalled();

    app.destroy();
    root.remove();
  });

  it("shows a toast and skips oversized files when inserting an image from the local file picker", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const toastSpy = vi.spyOn(ui, "showToast").mockImplementation(() => {});

    const app = new SpreadsheetApp(root, status);
    const sheetId = app.getCurrentSheetId();

    app.activateCell({ row: 3, col: 4 }, { scrollIntoView: false, focus: false });
    app.insertImageFromLocalFile();

    const input = root.querySelector<HTMLInputElement>('input[data-testid="insert-image-input"]');
    expect(input).not.toBeNull();

    const arrayBuffer = vi.fn(async () => {
      throw new Error("should not read oversized files");
    });
    const oversizedFile = {
      name: "big.png",
      type: "image/png",
      size: MAX_INSERT_IMAGE_BYTES + 1,
      arrayBuffer,
    } as any as File;

    Object.defineProperty(input!, "files", { configurable: true, value: [oversizedFile] });
    input!.dispatchEvent(new Event("change", { bubbles: true }));

    expect(arrayBuffer).not.toHaveBeenCalled();
    expect(toastSpy).toHaveBeenCalledWith("Image too large (>10MB). Choose a smaller file.", "warning");

    const doc: any = app.getDocument();
    expect(doc.getSheetDrawings(sheetId)).toHaveLength(0);

    app.destroy();
    root.remove();
  });

  it("does not persist image bytes when drawing insertion fails", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const toastSpy = vi.spyOn(ui, "showToast").mockImplementation(() => {});

    const app = new SpreadsheetApp(root, status);
    const sheetId = app.getCurrentSheetId();
    const docAny = app.getDocument() as any;

    // Force insertDrawing to fail so `insertImageFromPickedFile` bails before persisting bytes.
    docAny.insertDrawing = vi.fn(() => {
      throw new Error("insertDrawing failed");
    });

    const images = app.getDrawingImages();
    const setSpy = vi.spyOn(images, "set");

    const bytes = new Uint8Array([1, 2, 3, 4]);
    const file = new File([bytes], "test.png", { type: "image/png" });

    await (app as any).insertImageFromPickedFile(file);

    expect(setSpy).not.toHaveBeenCalled();
    expect((docAny.getSheetDrawings?.(sheetId) ?? []).length).toBe(0);
    expect(toastSpy).toHaveBeenCalled();

    app.destroy();
    root.remove();
  });

  it("restores focus when the file picker is dismissed without selecting a file", async () => {
    vi.useFakeTimers();

    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const focusSpy = vi.spyOn(app, "focus");

    app.insertImageFromLocalFile();

    // Simulate closing the picker without selecting a file. Some browsers do not
    // fire `change`, so SpreadsheetApp uses a focus-based fallback.
    window.dispatchEvent(new Event("focus"));

    vi.runOnlyPendingTimers();

    expect(focusSpy).toHaveBeenCalled();

    app.destroy();
    root.remove();
  });

  it("does not steal focus when the file picker is dismissed after switching sheets", async () => {
    vi.useFakeTimers();

    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const focusSpy = vi.spyOn(app, "focus");
    const doc: any = app.getDocument();

    // Ensure Sheet2 exists so we can switch away before the focus-based cancel handler runs.
    doc.setCellValue("Sheet2", { row: 0, col: 0 }, "X");

    app.insertImageFromLocalFile();
    app.activateSheet("Sheet2");
    expect(app.getCurrentSheetId()).toBe("Sheet2");

    focusSpy.mockClear();

    // Simulate closing the picker without selecting a file.
    window.dispatchEvent(new Event("focus"));
    vi.runOnlyPendingTimers();

    expect(focusSpy).not.toHaveBeenCalled();

    app.destroy();
    root.remove();
  });
});
