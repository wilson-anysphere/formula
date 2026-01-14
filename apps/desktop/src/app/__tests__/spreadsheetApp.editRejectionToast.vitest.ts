/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import * as Y from "yjs";

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

describe("SpreadsheetApp edit rejection toasts", () => {
  let priorGridMode: string | undefined;
  let mockedNow = 0;

  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();

    if (priorGridMode === undefined) delete process.env.DESKTOP_GRID_MODE;
    else process.env.DESKTOP_GRID_MODE = priorGridMode;
  });

  beforeEach(() => {
    // `showCollabEditRejectedToast` throttles identical messages for 1s using module-level state.
    // Tests run fast enough that the throttle can span multiple `it(...)` blocks, so we advance
    // `Date.now()` between tests to ensure each toast can be asserted independently.
    mockedNow += 2_000;
    vi.spyOn(Date, "now").mockReturnValue(mockedNow);

    priorGridMode = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "legacy";

    document.body.innerHTML = "";

    // `showCollabEditRejectedToast` throttles identical messages based on `Date.now()`. These
    // tests recreate `#toast-root` per case, so advance time between toast invocations to avoid
    // cross-test throttling hiding the warning.
    vi.spyOn(Date, "now").mockImplementation(() => {
      mockedNow += 2_000;
      return mockedNow;
    });

    const toastRoot = document.createElement("div");
    toastRoot.id = "toast-root";
    document.body.appendChild(toastRoot);

    const storage = createInMemoryLocalStorage();
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
    Object.defineProperty(window, "localStorage", { configurable: true, value: storage });
    storage.clear();

    // SpreadsheetApp schedules paints via requestAnimationFrame.
    Object.defineProperty(globalThis, "requestAnimationFrame", {
      configurable: true,
      value: (cb: FrameRequestCallback) => {
        cb(0);
        return 0;
      },
    });
    Object.defineProperty(globalThis, "cancelAnimationFrame", { configurable: true, value: () => {} });

    // jsdom lacks a real canvas implementation; SpreadsheetApp expects a 2D context.
    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: () => createMockCanvasContext(),
    });

    // jsdom doesn't ship ResizeObserver by default.
    (globalThis as any).ResizeObserver = class {
      observe() {}
      disconnect() {}
    };
  });

  it("shows a read-only toast when canEditCell blocks an in-cell edit", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);

    // Simulate a permissions guard installed by collab mode.
    (app as any).document.canEditCell = () => false;

    (app as any).applyEdit("Sheet1", { row: 0, col: 0 }, "hello");

    expect(document.querySelector("#toast-root")?.textContent ?? "").toContain("Read-only");

    app.destroy();
    root.remove();
  });

  it("shows a missing encryption key toast when collab encryption blocks an edit", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);

    (app as any).document.canEditCell = () => false;

    const ydoc = new Y.Doc();
    const cells = ydoc.getMap("cells");
    (app as any).collabSession = {
      cells,
      getEncryptionConfig: () => ({
        keyForCell: () => null,
        shouldEncryptCell: () => true,
      }),
    };

    (app as any).applyEdit("Sheet1", { row: 0, col: 0 }, "hello");

    expect(document.querySelector("#toast-root")?.textContent ?? "").toContain("Missing encryption key");

    app.destroy();
    root.remove();
  });

  it("includes the encrypted payload key id when a collab edit is blocked due to keyId mismatch", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);

    (app as any).document.canEditCell = () => false;

    const ydoc = new Y.Doc();
    const cells = ydoc.getMap("cells");
    const cellMap = new Y.Map<any>();
    cellMap.set("enc", {
      v: 1,
      alg: "AES-256-GCM",
      keyId: "k-range-1",
      ivBase64: "AA==",
      tagBase64: "AA==",
      ciphertextBase64: "AA==",
    });
    cells.set("Sheet1:0:0", cellMap);

    (app as any).collabSession = {
      cells,
      defaultSheetId: "Sheet1",
      getEncryptionConfig: () => ({
        keyForCell: () => ({ keyId: "k-other", keyBytes: new Uint8Array(32) }),
        shouldEncryptCell: () => true,
      }),
    };

    (app as any).applyEdit("Sheet1", { row: 0, col: 0 }, "hello");

    const content = document.querySelector("#toast-root")?.textContent ?? "";
    expect(content).toContain("Missing encryption key");
    expect(content).toContain("k-range-1");

    app.destroy();
    root.remove();
  });

  it("shows an unsupported-format toast when a collab edit is blocked due to unknown enc payload schema", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);

    (app as any).document.canEditCell = () => false;

    const ydoc = new Y.Doc();
    const cells = ydoc.getMap("cells");
    const cellMap = new Y.Map<any>();
    cellMap.set("enc", {
      v: 2,
      alg: "AES-256-GCM",
      keyId: "k-range-1",
      ivBase64: "AA==",
      tagBase64: "AA==",
      ciphertextBase64: "AA==",
    });
    cells.set("Sheet1:0:0", cellMap);

    (app as any).collabSession = {
      cells,
      defaultSheetId: "Sheet1",
      getEncryptionConfig: () => ({
        keyForCell: () => ({ keyId: "k-range-1", keyBytes: new Uint8Array(32) }),
        shouldEncryptCell: () => true,
      }),
    };

    (app as any).applyEdit("Sheet1", { row: 0, col: 0 }, "hello");

    const content = document.querySelector("#toast-root")?.textContent ?? "";
    expect(content).toContain("unsupported format");
    expect(content).toContain("Update Formula");

    app.destroy();
    root.remove();
  });

  it("shows an unsupported-format toast when enc payload schema is unknown even if the encryption key is missing", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);

    (app as any).document.canEditCell = () => false;

    const ydoc = new Y.Doc();
    const cells = ydoc.getMap("cells");
    const cellMap = new Y.Map<any>();
    cellMap.set("enc", {
      v: 2,
      alg: "AES-256-GCM",
      keyId: "k-range-1",
      ivBase64: "AA==",
      tagBase64: "AA==",
      ciphertextBase64: "AA==",
    });
    cells.set("Sheet1:0:0", cellMap);

    (app as any).collabSession = {
      cells,
      defaultSheetId: "Sheet1",
      getEncryptionConfig: () => ({
        keyForCell: () => null,
        shouldEncryptCell: () => true,
      }),
    };

    (app as any).applyEdit("Sheet1", { row: 0, col: 0 }, "hello");

    const content = document.querySelector("#toast-root")?.textContent ?? "";
    expect(content).toContain("unsupported format");
    expect(content).toContain("Update Formula");
    expect(content).toContain("k-range-1");

    app.destroy();
    root.remove();
  });

  it("includes the encrypted payload key id when enc is stored as a nested Y.Map (unsupported schema)", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);

    (app as any).document.canEditCell = () => false;

    const ydoc = new Y.Doc();
    const cells = ydoc.getMap("cells");
    const cellMap = new Y.Map<any>();
    const encMap = new Y.Map<any>();
    encMap.set("v", 2);
    encMap.set("alg", "AES-256-GCM");
    encMap.set("keyId", "k-range-1");
    encMap.set("ivBase64", "AA==");
    encMap.set("tagBase64", "AA==");
    encMap.set("ciphertextBase64", "AA==");
    cellMap.set("enc", encMap);
    cells.set("Sheet1:0:0", cellMap);

    (app as any).collabSession = {
      cells,
      defaultSheetId: "Sheet1",
      getEncryptionConfig: () => ({
        keyForCell: () => null,
        shouldEncryptCell: () => true,
      }),
    };

    (app as any).applyEdit("Sheet1", { row: 0, col: 0 }, "hello");

    const content = document.querySelector("#toast-root")?.textContent ?? "";
    expect(content).toContain("unsupported format");
    expect(content).toContain("Update Formula");
    expect(content).toContain("k-range-1");

    app.destroy();
    root.remove();
  });

  it("shows a read-only toast when the collab session role is viewer/commenter (isReadOnly)", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);

    // Simulate a read-only collab role; SpreadsheetApp should surface a toast when the user
    // attempts to start editing (rather than silently doing nothing).
    (app as any).collabSession = { isReadOnly: () => true };

    app.openCellEditorAtActiveCell();

    expect(document.querySelector("#toast-root")?.textContent ?? "").toContain("Read-only");

    app.destroy();
    root.remove();
  });

  it("shows a read-only toast when typing to start an in-cell edit in read-only collab mode", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };
 
    const app = new SpreadsheetApp(root, status);
    (app as any).collabSession = { isReadOnly: () => true };
 
    root.dispatchEvent(new KeyboardEvent("keydown", { key: "a", bubbles: true, cancelable: true }));
 
    expect(document.querySelector("#toast-root")?.textContent ?? "").toContain("Read-only");
    expect(root.querySelector("textarea.cell-editor--open")).toBeNull();
 
    app.destroy();
    root.remove();
  });

  it("shows a read-only toast when invoking AutoSum in read-only collab mode", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    (app as any).collabSession = { isReadOnly: () => true };

    app.autoSum();

    expect(document.querySelector("#toast-root")?.textContent ?? "").toContain("Read-only");

    app.destroy();
    root.remove();
  });

  it("shows a read-only toast when clearing contents in read-only collab mode", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    (app as any).collabSession = { isReadOnly: () => true };

    app.clearSelectionContents();

    expect(document.querySelector("#toast-root")?.textContent ?? "").toContain("Read-only");

    app.destroy();
    root.remove();
  });

  it("shows a read-only toast when clearing contents via clearContents in read-only collab mode", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const sheetId = app.getCurrentSheetId();
    app.getDocument().setCellValue(sheetId, "A1", "Seed", { label: "Seed Cell" });

    (app as any).collabSession = { isReadOnly: () => true };

    app.clearContents();

    expect(document.querySelector("#toast-root")?.textContent ?? "").toContain("Read-only");
    expect(app.getDocument().getCell(sheetId, "A1").value).toBe("Seed");

    app.destroy();
    root.remove();
  });

  it("shows a read-only toast when invoking Insert Cells in read-only collab mode", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    (app as any).collabSession = { isReadOnly: () => true };

    await app.insertCells({ startRow: 0, endRow: 0, startCol: 0, endCol: 0 }, "right");

    expect(document.querySelector("#toast-root")?.textContent ?? "").toContain("Read-only");

    app.destroy();
    root.remove();
  });

  it("shows a drawing toast when deleting a selected drawing in read-only collab mode", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    const sheetId = app.getCurrentSheetId();
    const drawing: any = {
      id: 1,
      kind: { type: "image", imageId: "img-1" },
      anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: 0, cy: 0 } },
      zOrder: 0,
    };
    app.getDocument().setSheetDrawings(sheetId, [drawing]);
    app.selectDrawing(drawing.id);

    (app as any).collabSession = { isReadOnly: () => true };

    app.deleteSelectedDrawing();

    expect(document.querySelector("#toast-root")?.textContent ?? "").toContain("edit drawings");
    expect(app.getDocument().getSheetDrawings(sheetId)).toHaveLength(1);

    app.destroy();
    root.remove();
  });

  it("shows a read-only toast when opening inline AI edit in read-only collab mode", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    (app as any).collabSession = { isReadOnly: () => true };

    app.openInlineAiEdit();

    expect(document.querySelector("#toast-root")?.textContent ?? "").toContain("Read-only");

    app.destroy();
    root.remove();
  });

  it("shows a read-only toast when canEditCell blocks paste", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    (app as any).document.canEditCell = () => false;
    const setRangeValuesSpy = vi.spyOn((app as any).document, "setRangeValues");
    (app as any).getClipboardProvider = async () => ({
      read: async () => ({ text: "hello" }),
    });

    await app.pasteClipboardToSelection();

    expect(document.querySelector("#toast-root")?.textContent ?? "").toContain("Read-only");
    expect(setRangeValuesSpy).not.toHaveBeenCalled();

    app.destroy();
    root.remove();
  });

  it("allows image-only paste even when canEditCell blocks cell edits", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    (app as any).document.canEditCell = () => false;

    const pasteImageSpy = vi.spyOn(app as any, "pasteClipboardImageAsDrawing").mockResolvedValue(true);
    (app as any).getClipboardProvider = async () => ({
      read: async () => ({
        // Some clipboard sources include `text/plain=""` alongside image payloads.
        text: "",
        pngBase64:
          "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+Xn0kAAAAASUVORK5CYII=",
      }),
    });

    await app.pasteClipboardToSelection();

    expect(pasteImageSpy).toHaveBeenCalled();
    expect(document.querySelector("#toast-root")?.textContent ?? "").toBe("");

    app.destroy();
    root.remove();
  });

  it("shows a read-only toast when canEditCell blocks cut", async () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    (app as any).document.canEditCell = () => false;
    const clearRangeSpy = vi.spyOn((app as any).document, "clearRange");

    await (app as any).cutSelectionToClipboard();

    expect(document.querySelector("#toast-root")?.textContent ?? "").toContain("Read-only");
    expect(clearRangeSpy).not.toHaveBeenCalled();

    app.destroy();
    root.remove();
  });

  it("shows a read-only toast when canEditCell blocks clear contents", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    (app as any).document.canEditCell = () => false;
    const clearRangeSpy = vi.spyOn((app as any).document, "clearRange");

    app.clearSelectionContents();

    expect(document.querySelector("#toast-root")?.textContent ?? "").toContain("Read-only");
    expect(clearRangeSpy).not.toHaveBeenCalled();

    app.destroy();
    root.remove();
  });

  it("shows an encryption toast when canEditCell blocks insertDate", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);
    (app as any).document.canEditCell = () => false;
    (app as any).collabSession = {
      getEncryptionConfig: () => ({
        keyForCell: () => null,
        shouldEncryptCell: () => true,
      }),
    };

    const setRangeValuesSpy = vi.spyOn((app as any).document, "setRangeValues");

    app.insertDate();

    expect(document.querySelector("#toast-root")?.textContent ?? "").toContain("Missing encryption key");
    expect(setRangeValuesSpy).not.toHaveBeenCalled();

    app.destroy();
    root.remove();
  });
});
