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
  const originalFirstRenderReporting = (globalThis as any).__FORMULA_STARTUP_FIRST_RENDER_REPORTING__;
  const originalTtiReported = (globalThis as any).__FORMULA_STARTUP_TTI_REPORTED__;
  const originalTtiReporting = (globalThis as any).__FORMULA_STARTUP_TTI_REPORTING__;
  const originalBootstrapped = (globalThis as any).__FORMULA_STARTUP_METRICS_BOOTSTRAPPED__;
  const originalWebviewReported = (globalThis as any).__FORMULA_STARTUP_WEBVIEW_LOADED_REPORTED__;

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
    (globalThis as any).__FORMULA_STARTUP_FIRST_RENDER_REPORTING__ = undefined;
    (globalThis as any).__FORMULA_STARTUP_TTI_REPORTED__ = undefined;
    (globalThis as any).__FORMULA_STARTUP_TTI_REPORTING__ = undefined;
    (globalThis as any).__FORMULA_STARTUP_METRICS_BOOTSTRAPPED__ = undefined;
    (globalThis as any).__FORMULA_STARTUP_WEBVIEW_LOADED_REPORTED__ = undefined;
  });

  afterEach(() => {
    vi.useRealTimers();
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
    (globalThis as any).__FORMULA_STARTUP_FIRST_RENDER_REPORTING__ = originalFirstRenderReporting;
    (globalThis as any).__FORMULA_STARTUP_TTI_REPORTED__ = originalTtiReported;
    (globalThis as any).__FORMULA_STARTUP_TTI_REPORTING__ = originalTtiReporting;
    (globalThis as any).__FORMULA_STARTUP_METRICS_BOOTSTRAPPED__ = originalBootstrapped;
    (globalThis as any).__FORMULA_STARTUP_WEBVIEW_LOADED_REPORTED__ = originalWebviewReported;
    vi.restoreAllMocks();
  });

  it("records a TTI mark and invokes the Rust host when running under Tauri", async () => {
    const invoke = (globalThis as any).__TAURI__?.core?.invoke as ReturnType<typeof vi.fn>;
    // In the real app we report first-render earlier in startup; keep this test focused on TTI.
    (globalThis as any).__FORMULA_STARTUP_FIRST_RENDER_REPORTED__ = true;

    await markStartupTimeToInteractive({ whenIdle: Promise.resolve() });

    const timings = getStartupTimings();
    expect(typeof timings.ttiFrontendMs).toBe("number");
    expect(Number.isFinite(timings.ttiFrontendMs!)).toBe(true);

    expect(invoke).toHaveBeenCalledWith("report_startup_tti");

    // Idempotent: should not report twice.
    await markStartupTimeToInteractive({ whenIdle: Promise.resolve() });
    expect(invoke).toHaveBeenCalledTimes(1);
  });

  it("can retry host TTI reporting after a transient invoke failure", async () => {
    const invoke = vi.fn()
      .mockRejectedValueOnce(new Error("transient"))
      .mockResolvedValueOnce(null);
    const listen = vi.fn().mockResolvedValue(() => {});
    (globalThis as any).__TAURI__ = { core: { invoke }, event: { listen } };
    (globalThis as any).__FORMULA_STARTUP_FIRST_RENDER_REPORTED__ = true;

    await markStartupTimeToInteractive({ whenIdle: Promise.resolve() });
    await markStartupTimeToInteractive({ whenIdle: Promise.resolve() });

    // First attempt fails, second succeeds.
    expect(invoke).toHaveBeenCalledWith("report_startup_tti");
    expect(invoke).toHaveBeenCalledTimes(2);
  });

  it("reports first render before TTI when first-render was not previously reported", async () => {
    const invoke = (globalThis as any).__TAURI__?.core?.invoke as ReturnType<typeof vi.fn>;

    // Ensure the first-render report has not happened yet.
    (globalThis as any).__FORMULA_STARTUP_FIRST_RENDER_REPORTED__ = undefined;

    await markStartupTimeToInteractive({ whenIdle: Promise.resolve() });

    expect(invoke).toHaveBeenCalledWith("report_startup_first_render");
    expect(invoke).toHaveBeenCalledWith("report_startup_tti");

    const firstRenderIdx = invoke.mock.calls.findIndex((args) => args[0] === "report_startup_first_render");
    const ttiIdx = invoke.mock.calls.findIndex((args) => args[0] === "report_startup_tti");
    expect(firstRenderIdx).toBeGreaterThanOrEqual(0);
    expect(ttiIdx).toBeGreaterThanOrEqual(0);
    expect(firstRenderIdx).toBeLessThan(ttiIdx);
  });

  it("retries TTI reporting when bootstrapped and __TAURI__ is injected late", async () => {
    vi.useFakeTimers();
    const originalRaf = (globalThis as any).requestAnimationFrame;
    try {
      const invoke = vi.fn().mockResolvedValue(null);
      const listen = vi.fn().mockResolvedValue(() => {});
      (globalThis as any).__FORMULA_STARTUP_METRICS_BOOTSTRAPPED__ = true;
      (globalThis as any).__FORMULA_STARTUP_FIRST_RENDER_REPORTED__ = true;

      // Make `nextFrame()` fast so we hit the retry path quickly.
      (globalThis as any).requestAnimationFrame = undefined;

      try {
        // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
        delete (globalThis as any).__TAURI__;
      } catch {
        (globalThis as any).__TAURI__ = undefined;
      }

      const promise = markStartupTimeToInteractive({ whenIdle: Promise.resolve() });

      setTimeout(() => {
        (globalThis as any).__TAURI__ = { core: { invoke }, event: { listen } };
      }, 100);

      await vi.runAllTimersAsync();
      await promise;

      expect(invoke).toHaveBeenCalledWith("report_startup_tti");
    } finally {
      (globalThis as any).requestAnimationFrame = originalRaf;
      vi.useRealTimers();
    }
  });

  it("can report TTI on a later call if invoke appears after the first retry window", async () => {
    vi.useFakeTimers();
    const originalRaf = (globalThis as any).requestAnimationFrame;
    try {
      const invoke = vi.fn().mockResolvedValue(null);
      const listen = vi.fn().mockResolvedValue(() => {});
      (globalThis as any).__FORMULA_STARTUP_METRICS_BOOTSTRAPPED__ = true;
      (globalThis as any).__FORMULA_STARTUP_FIRST_RENDER_REPORTED__ = true;

      // Make `nextFrame()` fast so the first call only waits on the invoke retry window.
      (globalThis as any).requestAnimationFrame = undefined;

      try {
        // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
        delete (globalThis as any).__TAURI__;
      } catch {
        (globalThis as any).__TAURI__ = undefined;
      }

      const first = markStartupTimeToInteractive({ whenIdle: Promise.resolve() });
      // Let the bounded invoke retry window elapse.
      await vi.advanceTimersByTimeAsync(10_500);
      await first;
      expect(invoke).not.toHaveBeenCalled();

      // Inject invoke after the retry window; a subsequent call should still report to the host.
      (globalThis as any).__TAURI__ = { core: { invoke }, event: { listen } };
      await markStartupTimeToInteractive({ whenIdle: Promise.resolve() });
      expect(invoke).toHaveBeenCalledWith("report_startup_tti");
    } finally {
      (globalThis as any).requestAnimationFrame = originalRaf;
      vi.useRealTimers();
    }
  });

  it("does not hang TTI instrumentation if requestAnimationFrame never fires", async () => {
    vi.useFakeTimers();
    const originalRaf = (globalThis as any).requestAnimationFrame;
    try {
      const invoke = vi.fn().mockResolvedValue(null);
      const listen = vi.fn().mockResolvedValue(() => {});
      (globalThis as any).__TAURI__ = { core: { invoke }, event: { listen } };
      (globalThis as any).__FORMULA_STARTUP_FIRST_RENDER_REPORTED__ = true;

      // Broken rAF implementation (never invokes callback).
      (globalThis as any).requestAnimationFrame = vi.fn();

      const promise = markStartupTimeToInteractive({ whenIdle: Promise.resolve() });
      await vi.runAllTimersAsync();
      await promise;

      expect(invoke).toHaveBeenCalledWith("report_startup_tti");
    } finally {
      (globalThis as any).requestAnimationFrame = originalRaf;
      vi.useRealTimers();
    }
  });

  it("does not hang TTI instrumentation if whenIdle never resolves", async () => {
    vi.useFakeTimers();
    const originalRaf = (globalThis as any).requestAnimationFrame;
    try {
      const invoke = vi.fn().mockResolvedValue(null);
      const listen = vi.fn().mockResolvedValue(() => {});
      (globalThis as any).__TAURI__ = { core: { invoke }, event: { listen } };
      (globalThis as any).__FORMULA_STARTUP_FIRST_RENDER_REPORTED__ = true;

      // Make `nextFrame()` fast so the test only waits on the idle timeout.
      (globalThis as any).requestAnimationFrame = undefined;

      const neverIdle = new Promise<void>(() => {});
      const promise = markStartupTimeToInteractive({ whenIdle: neverIdle, whenIdleTimeoutMs: 1000 });
      await vi.runAllTimersAsync();
      await promise;

      expect(invoke).toHaveBeenCalledWith("report_startup_tti");
    } finally {
      (globalThis as any).requestAnimationFrame = originalRaf;
      vi.useRealTimers();
    }
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

  it("can report first render even if __TAURI__ is injected after the call starts", async () => {
    const invoke = vi.fn().mockResolvedValue(null);
    const listen = vi.fn().mockResolvedValue(() => {});
    try {
      // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
      delete (globalThis as any).__TAURI__;
    } catch {
      (globalThis as any).__TAURI__ = undefined;
    }

    const promise = markStartupFirstRender();

    // Inject `__TAURI__` after the first microtask. This simulates hosts where the JS API is
    // attached slightly after module evaluation begins.
    queueMicrotask(() => {
      (globalThis as any).__TAURI__ = { core: { invoke }, event: { listen } };
    });

    await promise;
    expect(invoke).toHaveBeenCalledWith("report_startup_first_render");
  });

  it("does not hang first-render instrumentation if requestAnimationFrame never fires", async () => {
    vi.useFakeTimers();
    const originalRaf = (globalThis as any).requestAnimationFrame;
    try {
      const invoke = vi.fn().mockResolvedValue(null);
      const listen = vi.fn().mockResolvedValue(() => {});
      (globalThis as any).__TAURI__ = { core: { invoke }, event: { listen } };

      // Broken rAF implementation (never invokes callback).
      (globalThis as any).requestAnimationFrame = vi.fn();

      const promise = markStartupFirstRender();
      await vi.runAllTimersAsync();
      await promise;

      expect(invoke).toHaveBeenCalledWith("report_startup_first_render");
    } finally {
      (globalThis as any).requestAnimationFrame = originalRaf;
      vi.useRealTimers();
    }
  });

  it("retries for a short period when bootstrapped and __TAURI__ is injected late", async () => {
    vi.useFakeTimers();
    const originalRaf = (globalThis as any).requestAnimationFrame;
    try {
      const invoke = vi.fn().mockResolvedValue(null);
      const listen = vi.fn().mockResolvedValue(() => {});

      // Force the bootstrap hint so markStartupFirstRender knows it's safe to retry.
      (globalThis as any).__FORMULA_STARTUP_METRICS_BOOTSTRAPPED__ = true;

      // Make `nextFrame()` fast so we hit the retry path before we inject `__TAURI__`.
      (globalThis as any).requestAnimationFrame = undefined;

      try {
        // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
        delete (globalThis as any).__TAURI__;
      } catch {
        (globalThis as any).__TAURI__ = undefined;
      }

      const promise = markStartupFirstRender();

      setTimeout(() => {
        (globalThis as any).__TAURI__ = { core: { invoke }, event: { listen } };
      }, 100);

      await vi.runAllTimersAsync();
      await promise;

      expect(invoke).toHaveBeenCalledWith("report_startup_first_render");
    } finally {
      (globalThis as any).requestAnimationFrame = originalRaf;
      vi.useRealTimers();
    }
  });

  it("boots startup metrics as early side effects (report -> install listeners -> report again)", async () => {
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
    expect(invoke).toHaveBeenCalledTimes(2);
    expect(listen).toHaveBeenCalled();

    // The first report should happen before listener installation begins; the second should happen
    // after listeners are registered to request a re-emit of cached timings.
    expect(invoke.mock.invocationCallOrder[0]).toBeLessThan(listen.mock.invocationCallOrder[0]);
    expect(listen.mock.invocationCallOrder.at(-1)!).toBeLessThan(invoke.mock.invocationCallOrder[1]);
  });

  it("retries reporting once core.invoke becomes available (delayed __TAURI__ injection)", async () => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date("2025-01-01T00:00:00Z"));

    // Simulate an environment where the page runs inside a Tauri webview (UA contains "Tauri")
    // but the injected `__TAURI__` global is not available on the very first JS tick.
    const originalNavigator = (globalThis as any).navigator;
    Object.defineProperty(globalThis, "navigator", {
      configurable: true,
      value: { userAgent: "Tauri/2.9.5 (test)" },
    });

    try {
      try {
        // Start with no `__TAURI__` to force the bootstrap's "delayed inject" path.
        // (If the property is non-configurable, deletion may fail; ignore.)
        // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
        delete (globalThis as any).__TAURI__;
      } catch {
        // ignore
      }

      (globalThis as any).__FORMULA_STARTUP_TIMINGS__ = undefined;
      (globalThis as any).__FORMULA_STARTUP_TIMINGS_LISTENERS_INSTALLED__ = undefined;
      (globalThis as any).__FORMULA_STARTUP_METRICS_BOOTSTRAPPED__ = undefined;
      (globalThis as any).__FORMULA_STARTUP_WEBVIEW_LOADED_REPORTED__ = undefined;

      // Ensure we re-evaluate the bootstrap module even if a prior test imported it.
      vi.resetModules();
      await import("./startupMetricsBootstrap");

      const invoke = vi.fn().mockResolvedValue(null);
      const listen = vi.fn().mockResolvedValue(() => {});

      // Inject Tauri globals a moment later.
      setTimeout(() => {
        (globalThis as any).__TAURI__ = { core: { invoke }, event: { listen } };
      }, 5);

      // Advance time enough for the retry loop to observe the injected globals.
      // Use the async timer helper so promises scheduled by the retry loop get a chance to
      // resume and schedule follow-up timers within the same advancement window.
      await vi.advanceTimersByTimeAsync(10);
      await new Promise<void>((resolve) => queueMicrotask(resolve));
      await new Promise<void>((resolve) => queueMicrotask(resolve));

      expect(invoke).toHaveBeenCalledWith("report_startup_webview_loaded");
      expect(listen).toHaveBeenCalled();
    } finally {
      // Restore navigator so other tests don't depend on our shim.
      if (typeof originalNavigator === "undefined") {
        // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
        delete (globalThis as any).navigator;
      } else {
        Object.defineProperty(globalThis, "navigator", { configurable: true, value: originalNavigator });
      }
    }
  });

  it("boots when running under tauri:// even if __TAURI__ is injected later (protocol heuristic)", async () => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date("2025-01-01T00:00:00Z"));

    const originalLocation = (globalThis as any).location;
    Object.defineProperty(globalThis, "location", { configurable: true, value: { protocol: "tauri:" } });

    try {
      // Start with no `__TAURI__` to force the bootstrap's heuristic path.
      try {
        // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
        delete (globalThis as any).__TAURI__;
      } catch {
        // ignore
      }

      (globalThis as any).__FORMULA_STARTUP_TIMINGS__ = undefined;
      (globalThis as any).__FORMULA_STARTUP_TIMINGS_LISTENERS_INSTALLED__ = undefined;
      (globalThis as any).__FORMULA_STARTUP_METRICS_BOOTSTRAPPED__ = undefined;
      (globalThis as any).__FORMULA_STARTUP_WEBVIEW_LOADED_REPORTED__ = undefined;

      vi.resetModules();
      await import("./startupMetricsBootstrap");

      const invoke = vi.fn().mockResolvedValue(null);
      const listen = vi.fn().mockResolvedValue(() => {});
      setTimeout(() => {
        (globalThis as any).__TAURI__ = { core: { invoke }, event: { listen } };
      }, 5);

      // Use the async fake-timer helpers so pending promise continuations are flushed as timers run.
      await vi.advanceTimersByTimeAsync(20);
      await Promise.resolve();

      expect(invoke).toHaveBeenCalledWith("report_startup_webview_loaded");
      expect(listen).toHaveBeenCalled();
    } finally {
      if (typeof originalLocation === "undefined") {
        // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
        delete (globalThis as any).location;
      } else {
        Object.defineProperty(globalThis, "location", { configurable: true, value: originalLocation });
      }
    }
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

  it("is idempotent across repeated bootstrap evaluation (global guardrail)", async () => {
    const invoke = vi.fn().mockResolvedValue(null);
    const listen = vi.fn().mockResolvedValue(() => {});
    (globalThis as any).__TAURI__ = { core: { invoke }, event: { listen } };
    (globalThis as any).__FORMULA_STARTUP_TIMINGS__ = undefined;
    (globalThis as any).__FORMULA_STARTUP_TIMINGS_LISTENERS_INSTALLED__ = undefined;
    (globalThis as any).__FORMULA_STARTUP_METRICS_BOOTSTRAPPED__ = undefined;
    (globalThis as any).__FORMULA_STARTUP_WEBVIEW_LOADED_REPORTED__ = undefined;

    vi.resetModules();
    await import("./startupMetricsBootstrap");

    // Allow the listener-install retry loop to schedule its `.then(...)` work.
    await new Promise<void>((resolve) => queueMicrotask(resolve));
    await new Promise<void>((resolve) => queueMicrotask(resolve));

    const firstCallCount = invoke.mock.calls.length;
    expect(firstCallCount).toBeGreaterThan(0);

    // Force a re-evaluation of the module (simulates it being loaded via multiple entrypoints),
    // but keep the global bootstrapped marker intact.
    vi.resetModules();
    await import("./startupMetricsBootstrap");

    await new Promise<void>((resolve) => queueMicrotask(resolve));
    await new Promise<void>((resolve) => queueMicrotask(resolve));

    expect(invoke).toHaveBeenCalledTimes(firstCallCount);
  });
});
