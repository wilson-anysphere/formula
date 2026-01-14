import { afterEach, beforeEach } from "vitest";

// ---------------------------------------------------------------------------
// Global env defaults for desktop unit tests
// ---------------------------------------------------------------------------
//
// SpreadsheetApp supports both legacy (DOM/SVG) charts and unified canvas charts
// (ChartStore charts render through the drawings overlay). Canvas charts can be
// disabled via URL params (`?canvasCharts=0`) or env vars (`CANVAS_CHARTS=0`).
//
// Canvas charts are the default rendering path. Some unit tests still exercise
// legacy chart rendering; those tests should explicitly opt out via
// `process.env.CANVAS_CHARTS = "0"`.
//
// Keep the default deterministic across the suite, and reset after each test so
// tests that toggle/delete these env vars cannot leak state across files.
const DEFAULT_CANVAS_CHARTS = "1";
const DEFAULT_USE_CANVAS_CHARTS = "1";

beforeEach(() => {
  if (process.env.CANVAS_CHARTS === undefined) process.env.CANVAS_CHARTS = DEFAULT_CANVAS_CHARTS;
  if (process.env.USE_CANVAS_CHARTS === undefined) process.env.USE_CANVAS_CHARTS = DEFAULT_USE_CANVAS_CHARTS;
});

afterEach(() => {
  process.env.CANVAS_CHARTS = DEFAULT_CANVAS_CHARTS;
  process.env.USE_CANVAS_CHARTS = DEFAULT_USE_CANVAS_CHARTS;
});

class MemoryLocalStorage implements Storage {
  private readonly store = new Map<string, string>();

  get length(): number {
    return this.store.size;
  }

  clear(): void {
    this.store.clear();
  }

  getItem(key: string): string | null {
    return this.store.get(String(key)) ?? null;
  }

  key(index: number): string | null {
    if (index < 0) return null;
    return Array.from(this.store.keys())[index] ?? null;
  }

  removeItem(key: string): void {
    this.store.delete(String(key));
  }

  setItem(key: string, value: string): void {
    this.store.set(String(key), String(value));
  }
}

function storageUsable(storage: Storage | null | undefined): boolean {
  try {
    if (!storage) return false;
    // Node ships an experimental `globalThis.localStorage` that can be present but unusable unless the
    // process is started with `--localstorage-file`. Some methods may throw even if reads appear to work,
    // so probe the API surface our tests rely on.
    const probeKey = "vitest-probe";
    storage.setItem(probeKey, "1");
    storage.getItem(probeKey);
    storage.removeItem(probeKey);
    storage.clear();
    return true;
  } catch {
    return false;
  }
}

function installLocalStorage(storage: Storage): void {
  try {
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
  } catch {
    try {
      // eslint-disable-next-line no-global-assign
      (globalThis as any).localStorage = storage;
    } catch {
      // ignore
    }
  }

  if (typeof window !== "undefined") {
    try {
      Object.defineProperty(window, "localStorage", { configurable: true, value: storage });
    } catch {
      try {
        // eslint-disable-next-line no-global-assign
        (window as any).localStorage = storage;
      } catch {
        // ignore
      }
    }
  }
}

// Node 25+ ships an experimental `globalThis.localStorage` accessor that throws unless Node is started
// with `--localstorage-file`. Desktop tests rely on localStorage; provide a stable in-memory shim when
// the built-in accessor is unusable.
const existing = (() => {
  try {
    return (globalThis as any).localStorage as Storage | undefined;
  } catch {
    return undefined;
  }
})();

if (!storageUsable(existing)) {
  installLocalStorage(new MemoryLocalStorage());
}

// ---------------------------------------------------------------------------
// Environment defaults
// ---------------------------------------------------------------------------
//
// Canvas charts render ChartStore charts via the drawings overlay (so charts show up in
// `SpreadsheetApp.getDrawingObjects()`). Many unit tests exercise the drawings layer in
// isolation and assume charts are *not* part of the drawings list unless explicitly enabled.
//
// Keep canvas charts opt-in for unit tests by defaulting CANVAS_CHARTS off when neither
// CANVAS_CHARTS nor USE_CANVAS_CHARTS is set. Individual suites can opt into canvas charts
// by setting `process.env.CANVAS_CHARTS="1"`.
if (typeof process !== "undefined") {
  const env = process.env as Record<string, string | undefined>;
  if (env.CANVAS_CHARTS === undefined && env.USE_CANVAS_CHARTS === undefined) {
    env.CANVAS_CHARTS = "0";
  }
}

// ---------------------------------------------------------------------------
// DOM polyfills for jsdom-based unit tests
// ---------------------------------------------------------------------------
//
// Vitest's default environment for the desktop package is `node`, but many UI
// tests opt into jsdom via `@vitest-environment jsdom`. Keep these shims
// lightweight and feature-detect so they don't interfere with Node-only tests.

// JSDOM does not always provide PointerEvent. Some tests (e.g. shared-grid drawing interactions)
// rely on dispatching pointer events; provide a minimal shim backed by MouseEvent.
if (typeof (globalThis as any).PointerEvent === "undefined" && typeof (globalThis as any).MouseEvent === "function") {
  const Base = (globalThis as any).MouseEvent as typeof MouseEvent;
  class PointerEventShim extends Base {
    pointerId: number;
    pointerType: string;
    isPrimary: boolean;

    constructor(type: string, init: PointerEventInit = {}) {
      super(type, init);
      this.pointerId = typeof init.pointerId === "number" ? init.pointerId : 1;
      this.pointerType = typeof init.pointerType === "string" ? init.pointerType : "mouse";
      this.isPrimary = typeof init.isPrimary === "boolean" ? init.isPrimary : true;
    }
  }

  try {
    Object.defineProperty(globalThis, "PointerEvent", { configurable: true, value: PointerEventShim });
  } catch {
    // eslint-disable-next-line no-global-assign
    (globalThis as any).PointerEvent = PointerEventShim;
  }

  if (typeof window !== "undefined") {
    try {
      Object.defineProperty(window, "PointerEvent", { configurable: true, value: PointerEventShim });
    } catch {
      // eslint-disable-next-line no-global-assign
      (window as any).PointerEvent = PointerEventShim;
    }
  }
}

// JSDOM does not always implement pointer capture APIs (`setPointerCapture`, etc). SpreadsheetApp uses
// pointer capture to keep drag/selection gestures consistent; provide no-op shims so unit tests
// that dispatch pointer events do not crash.
if (typeof (globalThis as any).Element === "function") {
  const proto = (globalThis as any).Element.prototype as any;
  if (typeof proto.setPointerCapture !== "function") {
    try {
      Object.defineProperty(proto, "setPointerCapture", { configurable: true, value: () => {} });
    } catch {
      proto.setPointerCapture = () => {};
    }
  }
  if (typeof proto.releasePointerCapture !== "function") {
    try {
      Object.defineProperty(proto, "releasePointerCapture", { configurable: true, value: () => {} });
    } catch {
      proto.releasePointerCapture = () => {};
    }
  }
  if (typeof proto.hasPointerCapture !== "function") {
    try {
      Object.defineProperty(proto, "hasPointerCapture", { configurable: true, value: () => false });
    } catch {
      proto.hasPointerCapture = () => false;
    }
  }
}

// jsdom's File/Blob implementations may omit `arrayBuffer()` depending on the
// version. Picture insertion paths rely on it; provide a minimal polyfill so
// tests can exercise those codepaths deterministically.
if (typeof Blob !== "undefined" && typeof (Blob.prototype as any).arrayBuffer !== "function") {
  (Blob.prototype as any).arrayBuffer = function arrayBufferPolyfill(): Promise<ArrayBuffer> {
    // Prefer FileReader when available (jsdom).
    if (typeof FileReader !== "undefined") {
      return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result as ArrayBuffer);
        reader.onerror = () => reject(reader.error ?? new Error("Failed to read Blob as ArrayBuffer"));
        reader.readAsArrayBuffer(this as Blob);
      });
    }

    // Fallback: use Response if present (Node 18+ / undici).
    if (typeof Response !== "undefined") {
      return new Response(this as Blob).arrayBuffer();
    }

    return Promise.reject(new Error("Blob.arrayBuffer is not available in this environment"));
  };
}
