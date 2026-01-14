import { getTauriEventApiOrNull, getTauriInvokeOrNull, type TauriInvoke, type TauriListen } from "./api";

export type StartupTimings = {
  /**
   * Monotonic ms since native process start, reported by the Rust host.
   * (`startup:window-visible`)
   */
  windowVisibleMs?: number;
  /**
   * Monotonic ms since native process start, reported by the Rust host.
   *
   * This is recorded from Rust as soon as the native WebView reports that the
   * initial page load/navigation finished (`startup:webview-loaded`). It does
   * not include renderer JS bootstrap time; use `ttiMs` / `ttiFrontendMs` for
   * "app is interactive".
   */
  webviewLoadedMs?: number;
  /**
   * Monotonic ms since native process start, reported by the Rust host.
   * (`startup:first-render`)
   */
  firstRenderMs?: number;
  /**
   * Monotonic ms since native process start, reported by the Rust host.
   * (`startup:tti`)
   */
  ttiMs?: number;
  /**
   * Frontend-local mark (usually `performance.now()`), useful for web builds.
   */
  ttiFrontendMs?: number;
};

const GLOBAL_KEY = "__FORMULA_STARTUP_TIMINGS__";
const LISTENERS_KEY = "__FORMULA_STARTUP_TIMINGS_LISTENERS_INSTALLED__";
const BOOTSTRAPPED_KEY = "__FORMULA_STARTUP_METRICS_BOOTSTRAPPED__";
const FIRST_RENDER_REPORTED_KEY = "__FORMULA_STARTUP_FIRST_RENDER_REPORTED__";
const TTI_REPORTED_KEY = "__FORMULA_STARTUP_TTI_REPORTED__";
const HOST_INVOKE_RETRY_DEADLINE_MS = 10_000;
const FIRST_RENDER_REPORTING_KEY = "__FORMULA_STARTUP_FIRST_RENDER_REPORTING__";
const TTI_REPORTING_KEY = "__FORMULA_STARTUP_TTI_REPORTING__";

function getStore(): StartupTimings {
  const g = globalThis as any;
  const existing = g[GLOBAL_KEY];
  if (existing && typeof existing === "object") return existing as StartupTimings;
  const next: StartupTimings = {};
  g[GLOBAL_KEY] = next;
  return next;
}

function parseMs(payload: unknown): number | null {
  if (typeof payload === "number" && Number.isFinite(payload)) return payload;
  if (typeof payload === "string" && payload.trim() !== "") {
    const n = Number(payload);
    if (Number.isFinite(n)) return n;
  }
  return null;
}

function nowMs(): number {
  const perf = (globalThis as any)?.performance;
  if (perf && typeof perf.now === "function") {
    try {
      return perf.now();
    } catch {
      // Fall through to Date.now below.
    }
  }
  return Date.now();
}

async function nextFrame(): Promise<void> {
  const raf = (globalThis as any)?.requestAnimationFrame;
  if (typeof raf === "function") {
    await new Promise<void>((resolve) => {
      // Some environments can throttle or pause rAF (hidden webviews, headless runs, etc). Keep a
      // short timeout fallback so best-effort startup instrumentation doesn't hang forever.
      let done = false;
      const finish = () => {
        if (done) return;
        done = true;
        resolve();
      };

      const timeout = setTimeout(finish, 100);
      try {
        raf(() => {
          clearTimeout(timeout);
          finish();
        });
      } catch {
        clearTimeout(timeout);
        finish();
      }
    });
    return;
  }
  await new Promise<void>((resolve) => queueMicrotask(resolve));
}

function getTauriListen(): TauriListen | null {
  return getTauriEventApiOrNull()?.listen ?? null;
}

export function getStartupTimings(): StartupTimings {
  return { ...getStore() };
}

export async function installStartupTimingsListeners(): Promise<void> {
  const listen = getTauriListen();
  if (!listen) return;

  const g = globalThis as any;
  if (g[LISTENERS_KEY]) return;
  g[LISTENERS_KEY] = true;

  const store = getStore();

  const record = (key: keyof StartupTimings) => (event: any) => {
    const ms = parseMs(event?.payload);
    if (ms == null) return;
    (store as any)[key] = ms;
  };

  // Best-effort: keep these listeners extremely small and never throw.
  await Promise.all([
    listen("startup:window-visible", record("windowVisibleMs")).catch(() => {}),
    listen("startup:webview-loaded", record("webviewLoadedMs")).catch(() => {}),
    listen("startup:first-render", record("firstRenderMs")).catch(() => {}),
    listen("startup:tti", record("ttiMs")).catch(() => {}),
    listen("startup:metrics", (event: any) => {
      const payload = event?.payload;
      if (!payload || typeof payload !== "object") return;
      const windowVisible = parseMs((payload as any).window_visible_ms ?? (payload as any).windowVisibleMs);
      const webviewLoaded = parseMs((payload as any).webview_loaded_ms ?? (payload as any).webviewLoadedMs);
      const firstRender = parseMs((payload as any).first_render_ms ?? (payload as any).firstRenderMs);
      const tti = parseMs((payload as any).tti_ms ?? (payload as any).ttiMs);
      if (windowVisible != null) store.windowVisibleMs = windowVisible;
      if (webviewLoaded != null) store.webviewLoadedMs = webviewLoaded;
      if (firstRender != null) store.firstRenderMs = firstRender;
      if (tti != null) store.ttiMs = tti;
    }).catch(() => {}),
  ]);
}

/**
 * Notify the Rust host that the frontend has installed its startup timing
 * listeners and is ready to receive startup timing events.
 *
 * Tauri does not guarantee that early events are queued before JS listeners are
 * registered, so calling this after `installStartupTimingsListeners()` prompts
 * the host to (re-)emit cached `startup:*` events.
 *
 * Safe to call multiple times; the host will not overwrite the authoritative
 * `webviewLoadedMs` value recorded via a native page-load callback.
 *
 * Note: `webviewLoadedMs` is recorded in Rust via a native page-load callback.
 * Calling this is safe and will not overwrite earlier host-recorded timings; it
 * may re-emit cached metrics so late listeners can still observe them.
 *
 * No-op outside of Tauri.
 */
export function reportStartupWebviewLoaded(): void {
  try {
    const invoke = getTauriInvokeOrNull();
    if (!invoke) return;
    void invoke("report_startup_webview_loaded").catch(() => {});
  } catch {
    // Best-effort: reporting must never crash startup.
  }
}

/**
 * Notify the Rust host that the grid has rendered and is visible.
 *
 * Intended to be called at (or just after) the first meaningful UI paint of the
 * spreadsheet view.
 *
 * No-op outside of Tauri.
 */
export async function markStartupFirstRender(): Promise<StartupTimings> {
  const store = getStore();

  const g = globalThis as any;
  if (g[FIRST_RENDER_REPORTED_KEY]) return { ...store };
  if (g[FIRST_RENDER_REPORTING_KEY]) return { ...store };
  g[FIRST_RENDER_REPORTING_KEY] = true;

  try {
    // Give the renderer a frame (or two) to paint the initial grid before reporting.
    await nextFrame();
    await nextFrame();

    let invoke = getTauriInvokeOrNull();
    if (!invoke) {
      // In most environments `__TAURI__` is available immediately, but some host builds can delay
      // injection until shortly after module evaluation begins. When the early bootstrap runs we
      // set `__FORMULA_STARTUP_METRICS_BOOTSTRAPPED__` as a low-risk signal that we're in a Tauri
      // environment and should retry for a short bounded period.
      const shouldRetry = Boolean(g[BOOTSTRAPPED_KEY]);
      if (!shouldRetry) return { ...store };

      const deadlineMs = Date.now() + HOST_INVOKE_RETRY_DEADLINE_MS;
      let delayMs = 1;
      while (!invoke && Date.now() < deadlineMs) {
        await new Promise<void>((resolve) => setTimeout(resolve, delayMs));
        delayMs = Math.min(50, delayMs * 2);
        invoke = getTauriInvokeOrNull();
      }

      if (!invoke) return { ...store };
    }

    if (g[FIRST_RENDER_REPORTED_KEY]) return { ...store };

    try {
      await invoke("report_startup_first_render");
      g[FIRST_RENDER_REPORTED_KEY] = true;
    } catch {
      // Ignore host IPC failures; desktop startup should never be blocked on perf reporting.
    }

    return { ...store };
  } finally {
    try {
      // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
      delete (g as any)[FIRST_RENDER_REPORTING_KEY];
    } catch {
      g[FIRST_RENDER_REPORTING_KEY] = false;
    }
  }
}

/**
 * Record a "time-to-interactive" mark and (when running under Tauri) notify the
 * Rust host so it can compute a process-start monotonic TTI duration.
 */
export async function markStartupTimeToInteractive(options?: {
  /**
   * Optional promise (or callback returning a promise) representing "app is
   * stable and interactive" (e.g. `app.whenIdle()`).
   */
  whenIdle?: PromiseLike<unknown> | (() => PromiseLike<unknown>);
  /**
   * Maximum time (ms) to wait for `whenIdle` before proceeding anyway.
   *
   * This is best-effort startup instrumentation; we cap the wait so a hung "idle" promise
   * cannot prevent recording TTI forever (which would break the desktop startup perf harness).
   */
  whenIdleTimeoutMs?: number;
}): Promise<StartupTimings> {
  const store = getStore();
  const g = globalThis as any;

  const shouldMarkFrontend = typeof store.ttiFrontendMs !== "number";
  if (shouldMarkFrontend) {
    if (options?.whenIdle) {
      try {
        const idle = typeof options.whenIdle === "function" ? options.whenIdle() : options.whenIdle;
        const timeoutMsRaw = options.whenIdleTimeoutMs;
        const timeoutMs =
          typeof timeoutMsRaw === "number" && Number.isFinite(timeoutMsRaw) && timeoutMsRaw >= 0 ? timeoutMsRaw : 10_000;
        const idlePromise = Promise.resolve(idle).catch(() => {});
        if (timeoutMs === 0) {
          // Don't block on idle at all.
        } else {
          let timeoutId: ReturnType<typeof setTimeout> | null = null;
          await Promise.race([
            idlePromise,
            new Promise<void>((resolve) => {
              timeoutId = setTimeout(resolve, timeoutMs);
            }),
          ]);
          if (timeoutId != null) clearTimeout(timeoutId);
        }
      } catch {
        // Still record a mark even if the idle promise fails; this is best-effort
        // instrumentation and should never crash the app.
      }
    }

    // Give the renderer at least one frame to paint after becoming idle.
    await nextFrame();
    await nextFrame();

    const ttiFrontendMs = Math.round(nowMs());
    store.ttiFrontendMs = ttiFrontendMs;

    try {
      const perf = (globalThis as any)?.performance;
      if (perf && typeof perf.mark === "function") {
        perf.mark("formula:startup:tti");
      }
    } catch {
      // Ignore performance API errors.
    }
  }

  if (g[TTI_REPORTED_KEY]) return { ...store };
  if (g[TTI_REPORTING_KEY]) return { ...store };
  g[TTI_REPORTING_KEY] = true;

  const ensureFirstRenderReported = async (invoke: TauriInvoke): Promise<void> => {
    if (g[FIRST_RENDER_REPORTED_KEY]) return;
    try {
      await invoke("report_startup_first_render");
      g[FIRST_RENDER_REPORTED_KEY] = true;
    } catch {
      // Best-effort: ignore failures; if we can't report first render, still attempt TTI so the
      // desktop process can emit a `[startup] ...` line.
    }
  };

  try {
    const invoke = getTauriInvokeOrNull();
    if (invoke) {
      try {
        await ensureFirstRenderReported(invoke);
        await invoke("report_startup_tti");
        g[TTI_REPORTED_KEY] = true;
      } catch {
        // Ignore host IPC failures; desktop startup should never be blocked on perf reporting.
      }
    } else {
      // If the bootstrap ran but `__TAURI__` is injected late, retry briefly so we don't miss the
      // Rust-side TTI mark (required for the `[startup] ...` line the perf harness parses).
      const shouldRetry = Boolean(g[BOOTSTRAPPED_KEY]);
      if (shouldRetry) {
        const deadlineMs = Date.now() + HOST_INVOKE_RETRY_DEADLINE_MS;
        let delayMs = 1;
        let retriedInvoke: TauriInvoke | null = null;
        while (!retriedInvoke && Date.now() < deadlineMs) {
          await new Promise<void>((resolve) => setTimeout(resolve, delayMs));
          delayMs = Math.min(50, delayMs * 2);
          retriedInvoke = getTauriInvokeOrNull();
        }
        if (retriedInvoke) {
          try {
            await ensureFirstRenderReported(retriedInvoke);
            await retriedInvoke("report_startup_tti");
            g[TTI_REPORTED_KEY] = true;
          } catch {
            // Ignore host IPC failures; desktop startup should never be blocked on perf reporting.
          }
        }
      }
    }
  } finally {
    try {
      // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
      delete (g as any)[TTI_REPORTING_KEY];
    } catch {
      g[TTI_REPORTING_KEY] = false;
    }
  }

  return { ...store };
}
