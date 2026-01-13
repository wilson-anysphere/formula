/**
 * @vitest-environment jsdom
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import * as Y from "yjs";

import { InMemoryAwarenessHub, PresenceManager } from "../../collab/presence/index.js";

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

function createMockCollabSession(presence: PresenceManager): any {
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
    presence,
    setPermissions: vi.fn(),
  };
}

describe("SpreadsheetApp collab presence stale eviction", () => {
  let priorGridMode: string | undefined;

  afterEach(() => {
    if (priorGridMode === undefined) delete process.env.DESKTOP_GRID_MODE;
    else process.env.DESKTOP_GRID_MODE = priorGridMode;

    vi.useRealTimers();
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  beforeEach(() => {
    document.body.innerHTML = "";

    // Ensure tests are deterministic across grid modes.
    priorGridMode = process.env.DESKTOP_GRID_MODE;
    process.env.DESKTOP_GRID_MODE = "legacy";

    // Node 22 ships an experimental `localStorage` global that errors unless configured via flags.
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

    mocks.createCollabSession.mockReset();
    mocks.bindCollabSessionToDocumentController.mockReset();
  });

  it("evicts stale remote presences when SpreadsheetApp enables staleAfterMs", async () => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date(0));

    const hub = new InMemoryAwarenessHub();
    const localAwareness = hub.createAwareness(1);
    const remoteAwareness = hub.createAwareness(2);

    let localPresence: PresenceManager | null = null;

    mocks.createCollabSession.mockImplementation((options: any) => {
      localPresence = new PresenceManager(localAwareness, options.presence);
      return createMockCollabSession(localPresence);
    });
    mocks.bindCollabSessionToDocumentController.mockResolvedValue({ destroy: () => {} });

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
        persistenceEnabled: false,
        user: { id: "u1", name: "User 1", color: "#ff0000" },
      },
    });

    expect(mocks.createCollabSession).toHaveBeenCalledTimes(1);
    const options = mocks.createCollabSession.mock.calls[0]?.[0] as any;
    expect(options?.presence?.staleAfterMs).toBe(60_000);
    expect(options?.presence?.throttleMs).toBe(50);

    expect(localPresence).not.toBeNull();

    // Publish a remote collaborator presence, then stop updating it.
    const remotePresence = new PresenceManager(remoteAwareness, {
      user: { id: "u2", name: "User 2", color: "#00ff00" },
      activeSheet: options.presence.activeSheet,
      throttleMs: 0,
    });

    // PresenceManager notifies listeners synchronously for in-memory awareness updates.
    expect((app as any).remotePresences).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          id: "u2",
          name: "User 2",
        }),
      ]),
    );

    // Advance time past the configured staleAfterMs window; the remote presence should be dropped.
    await vi.advanceTimersByTimeAsync(60_001);

    expect((app as any).remotePresences).toEqual([]);

    remotePresence.destroy();
    app.destroy();
    root.remove();
  });
});
