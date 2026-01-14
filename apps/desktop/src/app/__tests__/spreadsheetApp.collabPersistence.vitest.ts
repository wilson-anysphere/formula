/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import * as Y from "yjs";

const mocks = vi.hoisted(() => {
  class IndexedDbCollabPersistence {}

  return {
    IndexedDbCollabPersistence,
    createCollabSession: vi.fn(),
    bindCollabSessionToDocumentController: vi.fn(),
  };
});

vi.mock("@formula/collab-persistence/indexeddb", () => ({
  IndexedDbCollabPersistence: mocks.IndexedDbCollabPersistence,
}));

vi.mock("@formula/collab-session", () => ({
  createCollabSession: mocks.createCollabSession,
  bindCollabSessionToDocumentController: mocks.bindCollabSessionToDocumentController,
}));

import { LazyIndexedDbCollabPersistence } from "../../collab/lazyIndexedDbCollabPersistence";
import { SpreadsheetApp } from "../spreadsheetApp";

function encodeBase64Url(value: string): string {
  return Buffer.from(value, "utf8")
    .toString("base64")
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/g, "");
}

function makeJwt(payload: unknown): string {
  const header = encodeBase64Url(JSON.stringify({ alg: "none", typ: "JWT" }));
  const body = encodeBase64Url(JSON.stringify(payload));
  return `${header}.${body}.sig`;
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

function createMockCollabSession(): any {
  const doc = new Y.Doc();
  const origin = { type: "collab-session-mock" };
  return {
    doc,
    cells: doc.getMap("cells"),
    sheets: doc.getArray("sheets"),
    metadata: doc.getMap("metadata"),
    namedRanges: doc.getMap("namedRanges"),
    origin,
    localOrigins: new Set([origin]),
    presence: null,
    setPermissions: vi.fn(),
    // Match the real CollabSession API so SpreadsheetApp doesn't need to monkeypatch `setPermissions`
    // during tests (which would break vi.fn assertions).
    onPermissionsChanged: vi.fn(() => () => {}),
  };
}

describe("SpreadsheetApp collab persistence", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  beforeEach(() => {
    document.body.innerHTML = "";

    const storage = createInMemoryLocalStorage();
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
    Object.defineProperty(window, "localStorage", { configurable: true, value: storage });
    storage.clear();

    Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
      configurable: true,
      value: () => createMockCanvasContext(),
    });

    // jsdom doesn't ship ResizeObserver by default.
    (globalThis as any).ResizeObserver = class {
      observe() {}
      disconnect() {}
    };

    mocks.createCollabSession.mockReset();
    mocks.bindCollabSessionToDocumentController.mockReset();

    mocks.createCollabSession.mockImplementation(() => createMockCollabSession());
    mocks.bindCollabSessionToDocumentController.mockResolvedValue({ destroy: () => {} });
  });

  it("constructs CollabSession with IndexedDbCollabPersistence when persistenceEnabled is true", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, {
      collab: {
        wsUrl: "ws://example.invalid",
        docId: "doc-123",
        persistenceEnabled: true,
        user: { id: "u1", name: "User 1", color: "#ff0000" },
      },
    });

    expect(mocks.createCollabSession).toHaveBeenCalledTimes(1);
    const options = mocks.createCollabSession.mock.calls[0]?.[0] as any;
    expect(options?.connection?.docId).toBe("doc-123");
    expect(options?.persistence).toBeInstanceOf(LazyIndexedDbCollabPersistence);
    expect(options?.offline).toBeUndefined();

    app.destroy();
    root.remove();
  });

  it("does not enable persistence when persistenceEnabled is false", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, {
      collab: {
        wsUrl: "ws://example.invalid",
        docId: "doc-456",
        persistenceEnabled: false,
        user: { id: "u2", name: "User 2", color: "#00ff00" },
      },
    });

    expect(mocks.createCollabSession).toHaveBeenCalledTimes(1);
    const options = mocks.createCollabSession.mock.calls[0]?.[0] as any;
    expect(options?.connection?.docId).toBe("doc-456");
    expect(options?.persistence).toBeUndefined();
    expect(options?.offline).toBeUndefined();

    app.destroy();
    root.remove();
  });

  it("supports legacy offlineEnabled toggle (offlineEnabled=false disables persistence)", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, {
      collab: {
        wsUrl: "ws://example.invalid",
        docId: "doc-legacy-offline",
        offlineEnabled: false,
        user: { id: "u3", name: "User 3", color: "#0000ff" },
      },
    });

    expect(mocks.createCollabSession).toHaveBeenCalledTimes(1);
    const options = mocks.createCollabSession.mock.calls[0]?.[0] as any;
    expect(options?.connection?.docId).toBe("doc-legacy-offline");
    expect(options?.persistence).toBeUndefined();
    expect(options?.offline).toBeUndefined();

    app.destroy();
    root.remove();
  });

  it("supports legacy offlineEnabled toggle (offlineEnabled=true enables persistence)", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, {
      collab: {
        wsUrl: "ws://example.invalid",
        docId: "doc-legacy-offline-true",
        offlineEnabled: true,
        user: { id: "u4", name: "User 4", color: "#ff00ff" },
      },
    });

    expect(mocks.createCollabSession).toHaveBeenCalledTimes(1);
    const options = mocks.createCollabSession.mock.calls[0]?.[0] as any;
    expect(options?.connection?.docId).toBe("doc-legacy-offline-true");
    expect(options?.persistence).toBeInstanceOf(LazyIndexedDbCollabPersistence);
    expect(options?.offline).toBeUndefined();

    app.destroy();
    root.remove();
  });

  it("trims collab connection fields read from the URL", () => {
    // Exercise the `resolveCollabOptionsFromUrl()` path (no explicit `opts.collab`).
    history.replaceState(
      null,
      "",
      `?collab=1&collabDocId=${encodeURIComponent(" doc-trim ")}&collabWsUrl=${encodeURIComponent(" ws://example.invalid ")}&collabToken=${encodeURIComponent(" tok ")}`,
    );

    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);

    expect(mocks.createCollabSession).toHaveBeenCalledTimes(1);
    const options = mocks.createCollabSession.mock.calls[0]?.[0] as any;
    expect(options?.connection?.docId).toBe("doc-trim");
    expect(options?.connection?.wsUrl).toBe("ws://example.invalid");
    expect(options?.connection?.token).toBe("tok");

    app.destroy();
    root.remove();
  });

  it("disables persistence when collabOffline=0 is present in the URL", () => {
    // Exercise the `resolveCollabOptionsFromUrl()` path (no explicit `opts.collab`).
    history.replaceState(
      null,
      "",
      `?collab=1&collabDocId=doc-789&collabWsUrl=${encodeURIComponent("ws://example.invalid")}&collabOffline=0`,
    );

    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);

    expect(mocks.createCollabSession).toHaveBeenCalledTimes(1);
    const options = mocks.createCollabSession.mock.calls[0]?.[0] as any;
    expect(options?.connection?.docId).toBe("doc-789");
    expect(options?.persistence).toBeUndefined();
    expect(options?.offline).toBeUndefined();

    app.destroy();
    root.remove();
  });

  it("disables persistence when collabPersistence=0 is present in the URL", () => {
    // Exercise the `resolveCollabOptionsFromUrl()` path (no explicit `opts.collab`).
    history.replaceState(
      null,
      "",
      `?collab=1&collabDocId=doc-790&collabWsUrl=${encodeURIComponent("ws://example.invalid")}&collabPersistence=0`,
    );

    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status);

    expect(mocks.createCollabSession).toHaveBeenCalledTimes(1);
    const options = mocks.createCollabSession.mock.calls[0]?.[0] as any;
    expect(options?.connection?.docId).toBe("doc-790");
    expect(options?.persistence).toBeUndefined();
    expect(options?.offline).toBeUndefined();

    app.destroy();
    root.remove();
  });

  it("does not crash when JWT-derived rangeRestrictions are invalid (falls back)", () => {
    const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});

    const token = makeJwt({
      sub: "user-123",
      role: "editor",
      rangeRestrictions: ["not-an-object"],
    });

    const collabSession = createMockCollabSession();
    const setPermissionsSpy = collabSession.setPermissions as ReturnType<typeof vi.fn>;
    setPermissionsSpy.mockImplementation((perms: any) => {
      if (
        Array.isArray(perms?.rangeRestrictions) &&
        perms.rangeRestrictions.some((r: unknown) => r == null || typeof r !== "object")
      ) {
        throw new Error("rangeRestrictions[0] invalid: restriction must be an object");
      }
    });

    mocks.createCollabSession.mockImplementationOnce(() => collabSession);

    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    let app: SpreadsheetApp | null = null;
    expect(() => {
      app = new SpreadsheetApp(root, status, {
        collab: {
          wsUrl: "ws://example.invalid",
          docId: "doc-bad-restrictions",
          token,
          user: { id: "u1", name: "User 1", color: "#ff0000" },
        },
      });
    }).not.toThrow();

    // SpreadsheetApp may wrap `collabSession.setPermissions` to keep the UI in sync
    // with permission changes. Assert against the original spy so this test stays
    // stable even when the method is wrapped.
    expect(setPermissionsSpy).toHaveBeenCalledTimes(2);
    expect(setPermissionsSpy.mock.calls[0]?.[0]).toMatchObject({
      role: "editor",
      userId: "user-123",
      rangeRestrictions: ["not-an-object"],
    });
    expect(setPermissionsSpy.mock.calls[1]?.[0]).toMatchObject({
      role: "editor",
      userId: "user-123",
      rangeRestrictions: [],
    });

    // Never log raw token contents.
    expect(warnSpy).toHaveBeenCalled();
    const logged = warnSpy.mock.calls.flat().join(" ");
    expect(logged).not.toContain(token);

    app!.destroy();
    root.remove();
  });
});
