/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import * as Y from "yjs";

const mocks = vi.hoisted(() => {
  return {
    createCollabSession: vi.fn(),
    bindCollabSessionToDocumentController: vi.fn(),
  };
});

vi.mock("@formula/collab-session", () => ({
  createCollabSession: mocks.createCollabSession,
  bindCollabSessionToDocumentController: mocks.bindCollabSessionToDocumentController,
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

type MockSession = {
  doc: Y.Doc;
  cells: Y.Map<any>;
  sheets: Y.Array<Y.Map<any>>;
  metadata: Y.Map<any>;
  namedRanges: Y.Map<any>;
  origin: any;
  localOrigins: Set<any>;
  presence: any;
  setPermissions: (perms: any) => void;
  getPermissions: () => any;
  isReadOnly: () => boolean;
  canComment: () => boolean;
  disconnect: () => void;
  destroy: () => void;
};

describe("SpreadsheetApp collab picture propagation", () => {
  const priorGridMode = process.env.DESKTOP_GRID_MODE;

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

    Object.defineProperty(globalThis, "createImageBitmap", {
      configurable: true,
      value: vi.fn(async () => ({})),
    });

    mocks.createCollabSession.mockReset();
    mocks.bindCollabSessionToDocumentController.mockReset();
  });

  afterEach(() => {
    if (priorGridMode === undefined) delete process.env.DESKTOP_GRID_MODE;
    else process.env.DESKTOP_GRID_MODE = priorGridMode;
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it("publishes inserted picture bytes to Yjs metadata and hydrates them in another app", async () => {
    const sharedDoc = new Y.Doc();
    const sheets = sharedDoc.getArray<Y.Map<any>>("sheets");
    if (sheets.length === 0) {
      const sheet = new Y.Map<any>();
      sheet.set("id", "Sheet1");
      sheet.set("name", "Sheet1");
      sheet.set("visibility", "visible");
      sheets.push([sheet]);
    }

    let sessionCounter = 0;
    const createSession = (userId: string): MockSession => {
      sessionCounter += 1;
      const origin = { type: "mock-session-origin", client: sessionCounter };
      return {
        doc: sharedDoc,
        cells: sharedDoc.getMap("cells"),
        sheets,
        metadata: sharedDoc.getMap("metadata"),
        namedRanges: sharedDoc.getMap("namedRanges"),
        origin,
        localOrigins: new Set<any>([origin]),
        presence: null,
        setPermissions: () => {},
        getPermissions: () => ({ role: "editor", rangeRestrictions: [], userId }),
        isReadOnly: () => false,
        canComment: () => true,
        disconnect: () => {},
        destroy: () => {},
      };
    };

    const sessions: MockSession[] = [];
    mocks.createCollabSession.mockImplementation((opts: any) => {
      const userId = String(opts?.presence?.user?.id ?? `u${sessions.length + 1}`);
      const session = createSession(userId);
      sessions.push(session);
      return session;
    });
    mocks.bindCollabSessionToDocumentController.mockResolvedValue({ destroy: () => {} });

    const rootA = createRoot();
    const rootB = createRoot();
    const statusA = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };
    const statusB = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const appA = new SpreadsheetApp(rootA, statusA, {
      collab: {
        wsUrl: "ws://example.invalid",
        docId: "doc-1",
        // Keep this unit test purely in-memory (no IndexedDbCollabPersistence).
        persistenceEnabled: false,
        user: { id: "uA", name: "User A", color: "#ff0000" },
      },
    });
    const appB = new SpreadsheetApp(rootB, statusB, {
      collab: {
        wsUrl: "ws://example.invalid",
        docId: "doc-1",
        persistenceEnabled: false,
        user: { id: "uB", name: "User B", color: "#00ff00" },
      },
    });

    const bytes = new Uint8Array([1, 2, 3, 4, 5]);
    const file = {
      name: "remote.png",
      type: "image/png",
      size: bytes.byteLength,
      arrayBuffer: async () => bytes.buffer.slice(0),
    } as unknown as File;

    await appA.insertPicturesFromFiles([file]);

    const sheetId = appA.getCurrentSheetId();
    const drawingsA = appA.getDocument().getSheetDrawings(sheetId) as any[];
    expect(drawingsA.length).toBeGreaterThan(0);
    const last = drawingsA[drawingsA.length - 1]!;
    expect(last?.kind?.type).toBe("image");
    const imageId = String(last?.kind?.imageId ?? "");
    expect(imageId).toBeTruthy();

    // Verify bytes were published into the shared metadata store.
    const metadata = sharedDoc.getMap("metadata");
    const drawingImages = metadata.get("drawingImages") as any;
    expect(drawingImages).toBeTruthy();
    const stored = drawingImages.get(imageId);
    expect(stored?.mimeType).toBe("image/png");
    expect(typeof stored?.bytesBase64).toBe("string");

    // Verify the remote app hydrates the bytes into its ImageStore.
    const hydrated = appB.getDrawingImages().get(imageId);
    expect(hydrated).toBeTruthy();
    expect(hydrated?.mimeType).toBe("image/png");
    expect(Array.from(hydrated?.bytes ?? [])).toEqual(Array.from(bytes));

    // Verify the remote app also sees the drawing object referencing the image.
    const drawingsB = appB.getDocument().getSheetDrawings(sheetId) as any[];
    expect(drawingsB.some((d) => d?.kind?.type === "image" && String(d?.kind?.imageId ?? "") === imageId)).toBe(true);

    appA.destroy();
    appB.destroy();
    rootA.remove();
    rootB.remove();
  });
});

