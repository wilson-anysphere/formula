import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { getStartupTimings, markStartupTimeToInteractive, reportStartupWebviewLoaded } from "./startupMetrics";

describe("startupMetrics", () => {
  const originalTauri = (globalThis as any).__TAURI__;
  const originalTimings = (globalThis as any).__FORMULA_STARTUP_TIMINGS__;

  beforeEach(() => {
    const invoke = vi.fn().mockResolvedValue(null);
    const listen = vi.fn().mockResolvedValue(() => {});
    (globalThis as any).__TAURI__ = { core: { invoke }, event: { listen } };
    (globalThis as any).__FORMULA_STARTUP_TIMINGS__ = undefined;
  });

  afterEach(() => {
    (globalThis as any).__TAURI__ = originalTauri;
    (globalThis as any).__FORMULA_STARTUP_TIMINGS__ = originalTimings;
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
});

