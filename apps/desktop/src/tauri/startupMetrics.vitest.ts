import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import {
  getStartupTimings,
  installStartupTimingsListeners,
  markStartupFirstRender,
  markStartupTimeToInteractive,
  reportStartupWebviewLoaded,
} from "./startupMetrics";

describe("startupMetrics", () => {
  const originalTauriDescriptor = Object.getOwnPropertyDescriptor(globalThis, "__TAURI__");
  const originalTimings = (globalThis as any).__FORMULA_STARTUP_TIMINGS__;
  const originalListenersInstalled = (globalThis as any).__FORMULA_STARTUP_TIMINGS_LISTENERS_INSTALLED__;
  const originalFirstRenderReported = (globalThis as any).__FORMULA_STARTUP_FIRST_RENDER_REPORTED__;

  beforeEach(() => {
    const invoke = vi.fn().mockResolvedValue(null);
    const listen = vi.fn().mockResolvedValue(() => {});
    Object.defineProperty(globalThis, "__TAURI__", {
      configurable: true,
      writable: true,
      value: { core: { invoke }, event: { listen } },
    });
    (globalThis as any).__FORMULA_STARTUP_TIMINGS__ = undefined;
    (globalThis as any).__FORMULA_STARTUP_TIMINGS_LISTENERS_INSTALLED__ = undefined;
    (globalThis as any).__FORMULA_STARTUP_FIRST_RENDER_REPORTED__ = undefined;
  });

  afterEach(() => {
    if (originalTauriDescriptor) {
      Object.defineProperty(globalThis, "__TAURI__", originalTauriDescriptor);
    } else {
      try {
        // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
        delete (globalThis as any).__TAURI__;
      } catch {
        // ignore
      }
    }
    (globalThis as any).__FORMULA_STARTUP_TIMINGS__ = originalTimings;
    (globalThis as any).__FORMULA_STARTUP_TIMINGS_LISTENERS_INSTALLED__ = originalListenersInstalled;
    (globalThis as any).__FORMULA_STARTUP_FIRST_RENDER_REPORTED__ = originalFirstRenderReported;
    vi.restoreAllMocks();
  });

  it("records a TTI mark and invokes the Rust host when running under Tauri", async () => {
    const invoke = (globalThis as any).__TAURI__?.core?.invoke as ReturnType<typeof vi.fn>;

    await markStartupTimeToInteractive({ whenIdle: Promise.resolve() });

    const timings = getStartupTimings();
    expect(typeof timings.ttiFrontendMs).toBe("number");
    expect(Number.isFinite(timings.ttiFrontendMs!)).toBe(true);

    expect(invoke).toHaveBeenCalledWith("report_startup_tti");

    // Idempotent: should not report twice.
    await markStartupTimeToInteractive({ whenIdle: Promise.resolve() });
    expect(invoke).toHaveBeenCalledTimes(1);
  });

  it("notifies the host that the webview is ready (when running under Tauri)", async () => {
    const invoke = (globalThis as any).__TAURI__?.core?.invoke as ReturnType<typeof vi.fn>;
    reportStartupWebviewLoaded();
    // fire-and-forget; flush microtasks once to allow the promise chain to schedule
    await new Promise<void>((resolve) => queueMicrotask(resolve));
    expect(invoke).toHaveBeenCalledWith("report_startup_webview_loaded");
  });

  it("can report webview-loaded before listeners install, then re-emit after listeners are ready", async () => {
    const listeners = new Map<string, (event: any) => void>();

    const invoke = vi.fn((cmd: string) => {
      if (cmd === "report_startup_webview_loaded") {
        listeners.get("startup:window-visible")?.({ payload: 123 });
        listeners.get("startup:webview-loaded")?.({ payload: 456 });
        listeners.get("startup:metrics")?.({ payload: { window_visible_ms: 123, webview_loaded_ms: 456 } });
      }
      return Promise.resolve(null);
    });

    const listen = vi.fn(async (event: string, handler: (event: any) => void) => {
      listeners.set(event, handler);
      return () => listeners.delete(event);
    });

    (globalThis as any).__TAURI__ = { core: { invoke }, event: { listen } };
    (globalThis as any).__FORMULA_STARTUP_TIMINGS__ = undefined;
    (globalThis as any).__FORMULA_STARTUP_TIMINGS_LISTENERS_INSTALLED__ = undefined;

    // First report happens before listeners exist, so no timings should be captured.
    reportStartupWebviewLoaded();
    await new Promise<void>((resolve) => queueMicrotask(resolve));
    expect(getStartupTimings().webviewLoadedMs).toBeUndefined();

    // Install listeners, then report again to re-emit timings.
    await installStartupTimingsListeners();
    reportStartupWebviewLoaded();
    await new Promise<void>((resolve) => queueMicrotask(resolve));

    const timings = getStartupTimings();
    expect(timings.windowVisibleMs).toBe(123);
    expect(timings.webviewLoadedMs).toBe(456);
  });

  it("notifies the host when the grid becomes visible (first render)", async () => {
    const invoke = (globalThis as any).__TAURI__?.core?.invoke as ReturnType<typeof vi.fn>;

    await markStartupFirstRender();
    expect(invoke).toHaveBeenCalledWith("report_startup_first_render");

    // Idempotent: should not report twice.
    await markStartupFirstRender();
    expect(invoke).toHaveBeenCalledTimes(1);
  });

  it("boots startup metrics as early side effects (install listeners -> request host re-emit)", async () => {
    const invoke = vi.fn().mockResolvedValue(null);
    const listen = vi.fn().mockResolvedValue(() => {});
    (globalThis as any).__TAURI__ = { core: { invoke }, event: { listen } };
    (globalThis as any).__FORMULA_STARTUP_TIMINGS__ = undefined;
    (globalThis as any).__FORMULA_STARTUP_TIMINGS_LISTENERS_INSTALLED__ = undefined;

    await import("./startupMetricsBootstrap");

    // Allow the listener-install promise and its `.finally(...)` to run.
    await new Promise<void>((resolve) => queueMicrotask(resolve));
    await new Promise<void>((resolve) => queueMicrotask(resolve));

    expect(invoke).toHaveBeenCalledWith("report_startup_webview_loaded");
    expect(invoke).toHaveBeenCalledTimes(1);
    expect(listen).toHaveBeenCalled();

    // The host re-emit request should happen after listeners are registered.
    expect(listen.mock.invocationCallOrder.at(-1)!).toBeLessThan(invoke.mock.invocationCallOrder[0]);
  });

  it("treats a throwing __TAURI__ getter as \"missing\" (best-effort hardening)", async () => {
    Object.defineProperty(globalThis, "__TAURI__", {
      configurable: true,
      get() {
        throw new Error("Blocked __TAURI__ access");
      },
    });

    // Should never throw, even if the host environment blocks access.
    expect(() => reportStartupWebviewLoaded()).not.toThrow();
    await expect(installStartupTimingsListeners()).resolves.toBeUndefined();
    await expect(markStartupFirstRender()).resolves.toEqual(expect.any(Object));
    await expect(markStartupTimeToInteractive({ whenIdle: Promise.resolve() })).resolves.toEqual(expect.any(Object));
  });
});
