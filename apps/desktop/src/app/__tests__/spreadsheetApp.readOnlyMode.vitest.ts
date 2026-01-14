/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import * as Y from "yjs";

vi.mock("../../extensions/ui.js", () => ({
  showToast: vi.fn(),
  showQuickPick: vi.fn(),
}));

import { showToast } from "../../extensions/ui.js";

const mocks = vi.hoisted(() => {
  class IndexedDbCollabPersistence {}

  return {
    IndexedDbCollabPersistence,
    createCollabSession: vi.fn(),
    bindCollabSessionToDocumentController: vi.fn(),
    bindSheetViewToCollabSession: vi.fn(),
  };
});

vi.mock("@formula/collab-persistence/indexeddb", () => ({
  IndexedDbCollabPersistence: mocks.IndexedDbCollabPersistence,
}));

vi.mock("@formula/collab-session", () => ({
  createCollabSession: mocks.createCollabSession,
  bindCollabSessionToDocumentController: mocks.bindCollabSessionToDocumentController,
  makeCellKey: (cell: any) => `${cell.sheetId}:${cell.row}:${cell.col}`,
}));

vi.mock("../../collab/sheetViewBinder", () => ({
  bindSheetViewToCollabSession: mocks.bindSheetViewToCollabSession,
}));

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

function createMockCollabSession(): any {
  const doc = new Y.Doc();
  const origin = { type: "collab-session-mock" };
  let permissions: { role: string; rangeRestrictions: unknown[]; userId: string | null } = {
    role: "editor",
    rangeRestrictions: [],
    userId: "u1",
  };
  const listeners = new Set<(p: any) => void>();

  const api = {
    doc,
    cells: doc.getMap("cells"),
    sheets: doc.getArray("sheets"),
    metadata: doc.getMap("metadata"),
    namedRanges: doc.getMap("namedRanges"),
    origin,
    localOrigins: new Set([origin]),
    presence: null,
    setPermissions: (next: any) => {
      permissions = {
        role: String(next?.role ?? ""),
        rangeRestrictions: Array.isArray(next?.rangeRestrictions) ? next.rangeRestrictions : [],
        userId: next?.userId ?? null,
      };
      for (const listener of [...listeners]) listener(api.getPermissions());
    },
    getPermissions: () => permissions,
    isReadOnly: () => permissions.role === "viewer" || permissions.role === "commenter",
    onPermissionsChanged: (listener: (p: any) => void) => {
      listeners.add(listener);
      listener(api.getPermissions());
      return () => listeners.delete(listener);
    },
  };

  return api;
}

describe("SpreadsheetApp read-only collab UX", () => {
  afterEach(() => {
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

    mocks.createCollabSession.mockReset();
    mocks.bindCollabSessionToDocumentController.mockReset();
    mocks.bindSheetViewToCollabSession.mockReset();

    mocks.createCollabSession.mockImplementation(() => createMockCollabSession());
    mocks.bindCollabSessionToDocumentController.mockResolvedValue({ destroy: () => {} });
    mocks.bindSheetViewToCollabSession.mockReturnValue({ destroy: () => {} });
  });

  it("shows a read-only indicator and refuses cell edits when collab permissions are viewer", () => {
    const root = createRoot();
    const readOnlyIndicator = document.createElement("div");
    readOnlyIndicator.hidden = true;
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
      readOnlyIndicator,
    };

    const app = new SpreadsheetApp(root, status, {
      collab: {
        wsUrl: "ws://example.invalid",
        docId: "doc-1",
        persistenceEnabled: false,
        user: { id: "u1", name: "User 1", color: "#ff0000" },
      },
    });

    const session = app.getCollabSession() as any;
    expect(session).not.toBeNull();

    const sheetId = app.getCurrentSheetId();
    // Seed an undoable edit before switching roles so we can ensure undo is blocked
    // after permissions flip to a read-only role.
    app.getDocument().setCellValue(sheetId, "A1", "Seed", { label: "Seed Cell" });
    expect(app.getDocument().getCell(sheetId, "A1").value).toBe("Seed");

    const inlineEditOverlay = root.querySelector<HTMLElement>('[data-testid="inline-edit-overlay"]');
    expect(inlineEditOverlay).toBeInstanceOf(HTMLElement);
    expect(inlineEditOverlay?.hidden).toBe(true);

    // Ensure that edit-only surfaces like inline AI edit are closed if permissions flip
    // to a read-only role mid-session.
    app.openInlineAiEdit();
    expect(inlineEditOverlay?.hidden).toBe(false);

    // Flip to viewer (read-only) and verify UI updates immediately.
    session.setPermissions({ role: "viewer", rangeRestrictions: [], userId: "u1" });
    expect(readOnlyIndicator.hidden).toBe(false);
    expect(readOnlyIndicator.textContent).toBe("Read-only (viewer)");
    expect(inlineEditOverlay?.hidden).toBe(true);

    // Read-only users should not be able to open inline edit via commands/menus.
    app.openInlineAiEdit();
    expect(inlineEditOverlay?.hidden).toBe(true);

    vi.mocked(showToast).mockClear();

    // Read-only users should not be able to undo/redo local edits (which would diverge
    // from the authoritative remote document).
    root.dispatchEvent(new KeyboardEvent("keydown", { key: "z", ctrlKey: true }));
    expect(app.getDocument().getCell(sheetId, "A1").value).toBe("Seed");
    expect(app.undo()).toBe(false);
    expect(showToast).toHaveBeenCalledWith(expect.stringContaining("undo/redo"), "warning");

    // Read-only users can still adjust sheet view state (e.g. Freeze Panes) locally.
    // Collaboration binders are responsible for preventing these view mutations from
    // being persisted into the shared Yjs document.
    expect(app.getFrozen()).toEqual({ frozenRows: 0, frozenCols: 0 });
    app.freezeTopRow();
    expect(app.getFrozen()).toEqual({ frozenRows: 1, frozenCols: 0 });

    // Sheet background images are persisted in sheet view state and should be blocked in read-only
    // roles (otherwise the viewer's local UI can diverge from the authoritative remote state).
    vi.mocked(showToast).mockClear();
    expect(app.getSheetBackgroundImageId(sheetId)).toBe(null);
    app.setSheetBackgroundImageId(sheetId, "bg.png");
    expect(app.getSheetBackgroundImageId(sheetId)).toBe(null);
    expect(showToast).toHaveBeenCalledWith(expect.stringContaining("background"), "warning");

    // Hide/unhide rows/cols are also sheet-view mutations and must no-op in read-only.
    vi.mocked(showToast).mockClear();
    const provider = (app as any).usedRangeProvider();
    expect(provider.isRowHidden(0)).toBe(false);
    app.hideRows([0]);
    expect(provider.isRowHidden(0)).toBe(false);
    expect(showToast).toHaveBeenCalledWith(expect.stringContaining("hide"), "warning");
    expect(provider.isColHidden(0)).toBe(false);
    app.hideCols([0]);
    expect(provider.isColHidden(0)).toBe(false);

    // Cutting a selected drawing should also be blocked and show a drawing-specific toast.
    vi.mocked(showToast).mockClear();
    const drawing: any = {
      id: 1,
      kind: { type: "image", imageId: "img-1" },
      anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: 0, cy: 0 } },
      zOrder: 0,
    };
    app.getDocument().setSheetDrawings(sheetId, [drawing]);
    app.selectDrawing(drawing.id);
    app.cut();
    expect(showToast).toHaveBeenCalledWith(expect.stringContaining("edit drawings"), "warning");
    expect(app.getDocument().getSheetDrawings(sheetId)).toHaveLength(1);

    // Attempt an in-grid edit (F2).
    root.dispatchEvent(new KeyboardEvent("keydown", { key: "F2" }));
    // The cell editor overlay element is always present; it only becomes "open"
    // when edit mode starts.
    expect(root.querySelector("textarea.cell-editor--open")).toBeNull();

    // Ensure the earlier seeded value is still present (no accidental clear/undo).
    expect(app.getDocument().getCell(sheetId, "A1").value).toBe("Seed");

    app.destroy();
    root.remove();
  });

  it("shows an error toast when encryptedRanges metadata schema is unsupported (fail-closed policy)", () => {
    const root = createRoot();
    const readOnlyIndicator = document.createElement("div");
    readOnlyIndicator.hidden = true;
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
      readOnlyIndicator,
    };

    mocks.createCollabSession.mockImplementation(() => {
      const session = createMockCollabSession();
      session.doc.transact(() => {
        session.metadata.set("encryptedRanges", { foo: "bar" });
      });
      return session;
    });

    vi.mocked(showToast).mockClear();

    const app = new SpreadsheetApp(root, status, {
      collab: {
        wsUrl: "ws://example.invalid",
        docId: "doc-1",
        persistenceEnabled: false,
        user: { id: "u1", name: "User 1", color: "#ff0000" },
      },
    });

    expect(showToast).toHaveBeenCalledWith(
      expect.stringMatching(/Encrypted range metadata is in an unsupported format/i),
      "error",
      expect.anything(),
    );

    app.destroy();
    root.remove();
  });

  it("prefers encrypted cell payload keyId over range policy keyId when resolving keys", async () => {
    const root = createRoot();
    const readOnlyIndicator = document.createElement("div");
    readOnlyIndicator.hidden = true;
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
      readOnlyIndicator,
    };

    let capturedOptions: any = null;
    mocks.createCollabSession.mockImplementation((opts: any) => {
      capturedOptions = opts;
      const session = createMockCollabSession();
      session.doc.transact(() => {
        // Policy says A1 should use "policy-key"...
        session.metadata.set("encryptedRanges", [
          {
            id: "er-1",
            sheetId: "Sheet1",
            startRow: 0,
            startCol: 0,
            endRow: 0,
            endCol: 0,
            keyId: "policy-key",
          },
        ]);

        // ...but the ciphertext payload is tagged with "payload-key".
        const cell = new Y.Map();
        cell.set("enc", {
          v: 1,
          alg: "AES-256-GCM",
          keyId: "payload-key",
          ivBase64: "a",
          tagBase64: "b",
          ciphertextBase64: "c",
        });
        session.cells.set("Sheet1:0:0", cell);
      });
      return session;
    });

    const app = new SpreadsheetApp(root, status, {
      collab: {
        wsUrl: "ws://example.invalid",
        docId: "doc-1",
        persistenceEnabled: false,
        user: { id: "u1", name: "User 1", color: "#ff0000" },
      },
    });

    expect(capturedOptions).not.toBeNull();
    expect(typeof capturedOptions?.encryption?.keyForCell).toBe("function");

    // Only cache the key referenced by the encrypted payload. If SpreadsheetApp falls back
    // to the policy key id ("policy-key"), the lookup should fail.
    const store = app.getCollabEncryptionKeyStore();
    expect(store).not.toBeNull();
    await store!.set("doc-1", "payload-key", Buffer.alloc(32, 1).toString("base64"));

    const resolved = capturedOptions.encryption.keyForCell({ sheetId: "Sheet1", row: 0, col: 0 });
    expect(resolved).toMatchObject({ keyId: "payload-key" });
    expect(resolved?.keyBytes).toBeInstanceOf(Uint8Array);
    expect(resolved?.keyBytes?.byteLength).toBe(32);

    app.destroy();
    root.remove();
  });

  it("falls back to the policy keyId when an encrypted cell payload key is unavailable", async () => {
    const root = createRoot();
    const readOnlyIndicator = document.createElement("div");
    readOnlyIndicator.hidden = true;
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
      readOnlyIndicator,
    };

    let capturedOptions: any = null;
    mocks.createCollabSession.mockImplementation((opts: any) => {
      capturedOptions = opts;
      const session = createMockCollabSession();
      session.doc.transact(() => {
        session.metadata.set("encryptedRanges", [
          {
            id: "er-1",
            sheetId: "Sheet1",
            startRow: 0,
            startCol: 0,
            endRow: 0,
            endCol: 0,
            keyId: "policy-key",
          },
        ]);

        const cell = new Y.Map();
        cell.set("enc", {
          v: 1,
          alg: "AES-256-GCM",
          keyId: "payload-key",
          ivBase64: "a",
          tagBase64: "b",
          ciphertextBase64: "c",
        });
        session.cells.set("Sheet1:0:0", cell);
      });
      return session;
    });

    const app = new SpreadsheetApp(root, status, {
      collab: {
        wsUrl: "ws://example.invalid",
        docId: "doc-1",
        persistenceEnabled: false,
        user: { id: "u1", name: "User 1", color: "#ff0000" },
      },
    });

    expect(capturedOptions).not.toBeNull();
    expect(typeof capturedOptions?.encryption?.keyForCell).toBe("function");

    // Only cache the policy key. When the ciphertext key is unavailable, SpreadsheetApp should
    // still return the policy key so clients can overwrite/rotate encrypted cells.
    const store = app.getCollabEncryptionKeyStore();
    expect(store).not.toBeNull();
    await store!.set("doc-1", "policy-key", Buffer.alloc(32, 2).toString("base64"));

    const resolved = capturedOptions.encryption.keyForCell({ sheetId: "Sheet1", row: 0, col: 0 });
    expect(resolved).toMatchObject({ keyId: "policy-key" });
    expect(resolved?.keyBytes).toBeInstanceOf(Uint8Array);
    expect(resolved?.keyBytes?.byteLength).toBe(32);

    app.destroy();
    root.remove();
  });
});
