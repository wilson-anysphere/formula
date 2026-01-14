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
  return {
    doc,
    cells: doc.getMap("cells"),
    sheets: doc.getArray("sheets"),
    metadata: doc.getMap("metadata"),
    namedRanges: doc.getMap("namedRanges"),
    origin,
    localOrigins: new Set<any>([origin]),
    presence: null,
    setPermissions: vi.fn(),
    getPermissions: () => ({ role: "editor", rangeRestrictions: [], userId: "u1" }),
    isReadOnly: () => false,
    canComment: () => true,
    disconnect: () => {},
    destroy: () => {},
  };
}

describe("SpreadsheetApp collab drawing image bytes", () => {
  beforeEach(() => {
    document.body.innerHTML = "";
    mocks.createCollabSession.mockReset();
    mocks.bindCollabSessionToDocumentController.mockReset();

    // CanvasGridRenderer schedules paints via rAF; provide a synchronous stub for jsdom.
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

    mocks.createCollabSession.mockImplementation(() => createMockCollabSession());
    mocks.bindCollabSessionToDocumentController.mockResolvedValue({ destroy: () => {} });
  });

  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it("hydrates bytes from Yjs metadata without storing them in DocumentController snapshots", () => {
    const root = createRoot();
    const status = {
      activeCell: document.createElement("div"),
      selectionRange: document.createElement("div"),
      activeValue: document.createElement("div"),
    };

    const app = new SpreadsheetApp(root, status, {
      collab: {
        wsUrl: "ws://example.invalid",
        docId: "doc-1",
        // Avoid exercising IndexedDbCollabPersistence in this unit test.
        persistenceEnabled: false,
        user: { id: "u1", name: "User 1", color: "#ff0000" },
      },
    });

    const session = (app as any).collabSession as any;
    expect(session?.metadata).toBeTruthy();

    const bytes = new Uint8Array([1, 2, 3, 4]);
    const base64 = Buffer.from(bytes).toString("base64");

    // Simulate a remote collaborator publishing image bytes into metadata.
    session.doc.transact(
      () => {
        const imagesMap = new Y.Map<any>();
        session.metadata.set("drawingImages", imagesMap);
        imagesMap.set("img-1", { mimeType: "image/png", bytesBase64: base64 });
      },
      { type: "remote" },
    );

    const hydrated = app.getDrawingImages().get("img-1");
    expect(hydrated).toBeTruthy();
    expect(Array.from(hydrated!.bytes)).toEqual([1, 2, 3, 4]);

    // Drawing images are intentionally stored out-of-band (IndexedDB + in-memory cache).
    // Even in collab mode, hydrating remote bytes must not bloat DocumentController snapshots.
    const snapshot = JSON.parse(new TextDecoder().decode(app.getDocument().encodeState()));
    expect(snapshot.images).toBeUndefined();

    app.destroy();
    root.remove();
  });
});

