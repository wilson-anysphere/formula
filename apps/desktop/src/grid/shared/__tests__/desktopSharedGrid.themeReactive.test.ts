// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { MockCellProvider } from "@formula/grid";
import { DesktopSharedGrid } from "../desktopSharedGrid";

function createMock2dContext(canvas: HTMLCanvasElement): CanvasRenderingContext2D {
  const noop = () => {};

  // DesktopSharedGrid relies on CanvasGridRenderer, which touches a broad surface
  // area of the 2D canvas context. For theme reactivity unit tests, a no-op
  // context is sufficient as long as the used methods/properties exist.
  return {
    canvas,
    fillStyle: "#000",
    strokeStyle: "#000",
    lineWidth: 1,
    font: "",
    textAlign: "left",
    textBaseline: "alphabetic",
    globalAlpha: 1,
    imageSmoothingEnabled: false,
    setTransform: noop,
    clearRect: noop,
    fillRect: noop,
    strokeRect: noop,
    beginPath: noop,
    rect: noop,
    clip: noop,
    fill: noop,
    stroke: noop,
    moveTo: noop,
    lineTo: noop,
    closePath: noop,
    setLineDash: noop,
    save: noop,
    restore: noop,
    drawImage: noop,
    translate: noop,
    rotate: noop,
    fillText: noop,
    measureText: (text: string) =>
      ({
        width: text.length * 6,
        actualBoundingBoxAscent: 8,
        actualBoundingBoxDescent: 2
      }) as TextMetrics
  } as unknown as CanvasRenderingContext2D;
}

describe("DesktopSharedGrid theme reactivity", () => {
  const originalGetContext = HTMLCanvasElement.prototype.getContext;

  beforeEach(() => {
    vi.stubGlobal("requestAnimationFrame", (cb: FrameRequestCallback) => {
      cb(0);
      return 0;
    });
    vi.stubGlobal("cancelAnimationFrame", () => {});

    HTMLCanvasElement.prototype.getContext = vi.fn(function (this: HTMLCanvasElement) {
      return createMock2dContext(this);
    }) as unknown as typeof HTMLCanvasElement.prototype.getContext;
  });

  afterEach(() => {
    HTMLCanvasElement.prototype.getContext = originalGetContext;
    vi.unstubAllGlobals();
    document.documentElement.removeAttribute("data-theme");
    document.head.querySelectorAll("style[data-testid='theme-style']").forEach((el) => el.remove());
    document.body.innerHTML = "";
  });

  it("updates its renderer theme when the document data-theme attribute changes", async () => {
    const style = document.createElement("style");
    style.dataset.testid = "theme-style";
    style.textContent = `
      :root { --formula-grid-bg: rgb(10, 20, 30); }
      :root[data-theme="dark"] { --formula-grid-bg: rgb(40, 50, 60); }

      /*
       * Note: JSDOM doesn't currently inherit custom properties from :root into
       * descendants via getComputedStyle(). Since DesktopSharedGrid reads CSS
       * vars from its container element, also apply the token directly to the
       * container so the test matches browser behavior.
       */
      :root .grid-host { --formula-grid-bg: rgb(10, 20, 30); }
      :root[data-theme="dark"] .grid-host { --formula-grid-bg: rgb(40, 50, 60); }
    `;
    document.head.appendChild(style);

    document.documentElement.setAttribute("data-theme", "light");

    const container = document.createElement("div");
    container.className = "grid-host";
    document.body.appendChild(container);

    const provider = new MockCellProvider({ rowCount: 10, colCount: 10 });

    const canvases = {
      grid: document.createElement("canvas"),
      content: document.createElement("canvas"),
      selection: document.createElement("canvas")
    };

    const scrollbars = {
      vTrack: document.createElement("div"),
      vThumb: document.createElement("div"),
      hTrack: document.createElement("div"),
      hThumb: document.createElement("div")
    };

    const grid = new DesktopSharedGrid({
      container,
      provider,
      rowCount: 10,
      colCount: 10,
      canvases,
      scrollbars
    });

    try {
      expect(grid.renderer.getTheme().gridBg).toBe("rgb(10, 20, 30)");

      document.documentElement.setAttribute("data-theme", "dark");
      // MutationObserver flush.
      await Promise.resolve();

      expect(grid.renderer.getTheme().gridBg).toBe("rgb(40, 50, 60)");
    } finally {
      grid.destroy();
    }
  });

  it("updates its renderer theme when the container style attribute changes", async () => {
    const container = document.createElement("div");
    container.setAttribute("style", "--formula-grid-bg: rgb(10, 20, 30);");
    document.body.appendChild(container);

    const provider = new MockCellProvider({ rowCount: 10, colCount: 10 });

    const canvases = {
      grid: document.createElement("canvas"),
      content: document.createElement("canvas"),
      selection: document.createElement("canvas")
    };

    const scrollbars = {
      vTrack: document.createElement("div"),
      vThumb: document.createElement("div"),
      hTrack: document.createElement("div"),
      hThumb: document.createElement("div")
    };

    const grid = new DesktopSharedGrid({
      container,
      provider,
      rowCount: 10,
      colCount: 10,
      canvases,
      scrollbars
    });

    try {
      expect(grid.renderer.getTheme().gridBg).toBe("rgb(10, 20, 30)");

      container.setAttribute("style", "--formula-grid-bg: rgb(40, 50, 60);");
      await Promise.resolve();

      expect(grid.renderer.getTheme().gridBg).toBe("rgb(40, 50, 60)");
    } finally {
      grid.destroy();
    }
  });

  it("updates its renderer theme when a relevant matchMedia query fires a change event", () => {
    // Disable MutationObserver so this test specifically exercises the matchMedia path.
    vi.stubGlobal("MutationObserver", undefined);

    const listenersByQuery = new Map<string, Set<() => void>>();
    const matchMediaMock = vi.fn((query: string) => {
      const listeners = new Set<() => void>();
      listenersByQuery.set(query, listeners);

      return {
        media: query,
        matches: false,
        addEventListener: (type: string, listener: () => void) => {
          if (type !== "change") return;
          listeners.add(listener);
        },
        removeEventListener: (type: string, listener: () => void) => {
          if (type !== "change") return;
          listeners.delete(listener);
        }
      } as unknown as MediaQueryList;
    });
    vi.stubGlobal("matchMedia", matchMediaMock);

    const container = document.createElement("div");
    container.style.setProperty("--formula-grid-bg", "rgb(10, 20, 30)");
    document.body.appendChild(container);

    const provider = new MockCellProvider({ rowCount: 10, colCount: 10 });

    const canvases = {
      grid: document.createElement("canvas"),
      content: document.createElement("canvas"),
      selection: document.createElement("canvas")
    };

    const scrollbars = {
      vTrack: document.createElement("div"),
      vThumb: document.createElement("div"),
      hTrack: document.createElement("div"),
      hThumb: document.createElement("div")
    };

    const grid = new DesktopSharedGrid({
      container,
      provider,
      rowCount: 10,
      colCount: 10,
      canvases,
      scrollbars
    });

    try {
      expect(matchMediaMock).toHaveBeenCalledWith("(prefers-color-scheme: dark)");
      expect(matchMediaMock).toHaveBeenCalledWith("(prefers-contrast: more)");
      expect(matchMediaMock).toHaveBeenCalledWith("(forced-colors: active)");
      expect(matchMediaMock).toHaveBeenCalledWith("(prefers-reduced-motion: reduce)");

      const requiredQueries = [
        "(prefers-color-scheme: dark)",
        "(prefers-contrast: more)",
        "(forced-colors: active)",
        "(prefers-reduced-motion: reduce)"
      ];
      for (const query of requiredQueries) {
        expect(listenersByQuery.get(query)?.size).toBeGreaterThan(0);
      }

      expect(grid.renderer.getTheme().gridBg).toBe("rgb(10, 20, 30)");

      container.style.setProperty("--formula-grid-bg", "rgb(40, 50, 60)");

      // Simulate a system preference change (e.g. dark mode / forced-colors) by firing the
      // media query listener. The implementation should re-resolve the theme from CSS vars.
      const listeners = listenersByQuery.get("(prefers-color-scheme: dark)");
      expect(listeners?.size).toBeGreaterThan(0);
      for (const listener of listeners ?? []) listener();

      expect(grid.renderer.getTheme().gridBg).toBe("rgb(40, 50, 60)");
    } finally {
      grid.destroy();
    }

    // Ensure listeners were detached on destroy.
    for (const query of [
      "(prefers-color-scheme: dark)",
      "(prefers-contrast: more)",
      "(forced-colors: active)",
      "(prefers-reduced-motion: reduce)"
    ]) {
      expect(listenersByQuery.get(query)?.size ?? 0).toBe(0);
    }
  });

  it("disconnects MutationObservers on destroy and uses them to refresh theme vars", () => {
    type ObserverRecord = {
      target: Element | null;
      options: MutationObserverInit | undefined;
      callback: MutationCallback;
      observe: ReturnType<typeof vi.fn>;
      disconnect: ReturnType<typeof vi.fn>;
    };

    const records: ObserverRecord[] = [];

    class MockMutationObserver {
      target: Element | null = null;
      options: MutationObserverInit | undefined;
      observe = vi.fn((target: Element, options?: MutationObserverInit) => {
        this.target = target;
        this.options = options;
      });
      disconnect = vi.fn();

      constructor(public callback: MutationCallback) {
        records.push({
          target: null,
          options: undefined,
          callback,
          observe: this.observe,
          disconnect: this.disconnect
        });
      }
    }

    vi.stubGlobal("MutationObserver", MockMutationObserver as unknown as typeof MutationObserver);

    const container = document.createElement("div");
    container.style.setProperty("--formula-grid-bg", "rgb(10, 20, 30)");
    document.body.appendChild(container);

    const provider = new MockCellProvider({ rowCount: 10, colCount: 10 });

    const canvases = {
      grid: document.createElement("canvas"),
      content: document.createElement("canvas"),
      selection: document.createElement("canvas")
    };

    const scrollbars = {
      vTrack: document.createElement("div"),
      vThumb: document.createElement("div"),
      hTrack: document.createElement("div"),
      hThumb: document.createElement("div")
    };

    const grid = new DesktopSharedGrid({
      container,
      provider,
      rowCount: 10,
      colCount: 10,
      canvases,
      scrollbars
    });

    expect(grid.renderer.getTheme().gridBg).toBe("rgb(10, 20, 30)");
    const renderSpy = vi.spyOn(grid.renderer, "renderImmediately").mockImplementation(() => {});

    // Should observe the container + root + body.
    expect(records.length).toBeGreaterThanOrEqual(1);
    const observedTargets = records.map((r) => (r.observe.mock.calls[0]?.[0] as Element | undefined) ?? null);
    expect(observedTargets).toEqual(expect.arrayContaining([container, document.documentElement, document.body]));

    for (const record of records) {
      const options = record.observe.mock.calls[0]?.[1] as MutationObserverInit | undefined;
      expect(options?.attributes).toBe(true);
      expect(options?.attributeFilter).toEqual(
        expect.arrayContaining(["style", "class", "data-theme", "data-reduced-motion"])
      );
    }

    // Simulate a CSS var change on the container.
    container.style.setProperty("--formula-grid-bg", "rgb(40, 50, 60)");

    // Trigger whichever observer was installed for the container.
    const containerObserverIndex = observedTargets.findIndex((t) => t === container);
    expect(containerObserverIndex).toBeGreaterThanOrEqual(0);
    const containerRecord = records[containerObserverIndex]!;
    containerRecord.callback([], {} as any);

    expect(grid.renderer.getTheme().gridBg).toBe("rgb(40, 50, 60)");
    expect(renderSpy).toHaveBeenCalledTimes(1);

    grid.destroy();

    for (const record of records) {
      expect(record.disconnect).toHaveBeenCalledTimes(1);
    }
  });

  it("supports legacy matchMedia addListener/removeListener APIs", () => {
    // Disable MutationObserver so this test isolates the matchMedia subscription logic.
    vi.stubGlobal("MutationObserver", undefined);

    const listenersByQuery = new Map<string, Set<() => void>>();
    const matchMediaMock = vi.fn((query: string) => {
      const listeners = new Set<() => void>();
      listenersByQuery.set(query, listeners);

      return {
        media: query,
        matches: false,
        addListener: (listener: () => void) => {
          listeners.add(listener);
        },
        removeListener: (listener: () => void) => {
          listeners.delete(listener);
        }
      } as unknown as MediaQueryList;
    });
    vi.stubGlobal("matchMedia", matchMediaMock);

    const container = document.createElement("div");
    container.style.setProperty("--formula-grid-bg", "rgb(10, 20, 30)");
    document.body.appendChild(container);

    const provider = new MockCellProvider({ rowCount: 10, colCount: 10 });

    const canvases = {
      grid: document.createElement("canvas"),
      content: document.createElement("canvas"),
      selection: document.createElement("canvas")
    };

    const scrollbars = {
      vTrack: document.createElement("div"),
      vThumb: document.createElement("div"),
      hTrack: document.createElement("div"),
      hThumb: document.createElement("div")
    };

    const grid = new DesktopSharedGrid({
      container,
      provider,
      rowCount: 10,
      colCount: 10,
      canvases,
      scrollbars
    });

    try {
      expect(matchMediaMock).toHaveBeenCalledWith("(prefers-color-scheme: dark)");
      expect(matchMediaMock).toHaveBeenCalledWith("(prefers-contrast: more)");
      expect(matchMediaMock).toHaveBeenCalledWith("(forced-colors: active)");
      expect(matchMediaMock).toHaveBeenCalledWith("(prefers-reduced-motion: reduce)");

      // Ensure listeners were registered via the legacy API.
      for (const query of [
        "(prefers-color-scheme: dark)",
        "(prefers-contrast: more)",
        "(forced-colors: active)",
        "(prefers-reduced-motion: reduce)"
      ]) {
        expect(listenersByQuery.get(query)?.size).toBeGreaterThan(0);
      }

      expect(grid.renderer.getTheme().gridBg).toBe("rgb(10, 20, 30)");

      container.style.setProperty("--formula-grid-bg", "rgb(40, 50, 60)");
      for (const listener of listenersByQuery.get("(forced-colors: active)") ?? []) listener();

      expect(grid.renderer.getTheme().gridBg).toBe("rgb(40, 50, 60)");
    } finally {
      grid.destroy();
    }

    // Ensure destroy removes legacy listeners.
    for (const query of [
      "(prefers-color-scheme: dark)",
      "(prefers-contrast: more)",
      "(forced-colors: active)",
      "(prefers-reduced-motion: reduce)"
    ]) {
      expect(listenersByQuery.get(query)?.size ?? 0).toBe(0);
    }
  });
});
