import { getTauriEventApiOrNull, type TauriListen } from "./api";

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
const FIRST_RENDER_REPORTED_KEY = "__FORMULA_STARTUP_FIRST_RENDER_REPORTED__";

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
    await new Promise<void>((resolve) => raf(() => resolve()));
    return;
  }
  await new Promise<void>((resolve) => queueMicrotask(resolve));
}

type TauriInvoke = (cmd: string, args?: any) => Promise<any>;

function getTauriInvoke(): TauriInvoke | null {
  const invoke = (globalThis as any).__TAURI__?.core?.invoke as TauriInvoke | undefined;
  return typeof invoke === "function" ? invoke : null;
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
 * Note: `webviewLoadedMs` is recorded in Rust via a native page-load callback.
 * Calling this is safe and will not overwrite earlier host-recorded timings; it
 * may re-emit cached metrics so late listeners can still observe them.
 *
 * No-op outside of Tauri.
 */
export function reportStartupWebviewLoaded(): void {
  const invoke = getTauriInvoke();
  if (!invoke) return;
  void invoke("report_startup_webview_loaded").catch(() => {});
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
  const invoke = getTauriInvoke();
  if (!invoke) return { ...store };

  const g = globalThis as any;
  if (g[FIRST_RENDER_REPORTED_KEY]) return { ...store };
  g[FIRST_RENDER_REPORTED_KEY] = true;

  // Give the renderer a frame (or two) to paint the initial grid before reporting.
  await nextFrame();
  await nextFrame();

  try {
    await invoke("report_startup_first_render");
  } catch {
    // Ignore host IPC failures; desktop startup should never be blocked on perf reporting.
  }

  return { ...store };
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
}): Promise<StartupTimings> {
  const store = getStore();
  if (typeof store.ttiFrontendMs === "number") return { ...store };

  if (options?.whenIdle) {
    try {
      const idle = typeof options.whenIdle === "function" ? options.whenIdle() : options.whenIdle;
      await idle;
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

  const invoke = getTauriInvoke();
  if (invoke) {
    try {
      await invoke("report_startup_tti");
    } catch {
      // Ignore host IPC failures; desktop startup should never be blocked on perf reporting.
    }
  }

  return { ...store };
}
