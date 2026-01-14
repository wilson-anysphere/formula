import { installStartupTimingsListeners, reportStartupWebviewLoaded } from "./startupMetrics.js";
import { hasTauri as hasTauriRuntime } from "./api";

// Startup performance instrumentation (no-op for web builds).
//
// `webviewLoadedMs` is recorded natively by the Rust host when the main WebView finishes its
// initial navigation. Tauri does not guarantee events are queued before listeners are installed,
// so we install listeners early and then ask the host to (re-)emit the cached timings once ready.

const BOOTSTRAPPED_KEY = "__FORMULA_STARTUP_METRICS_BOOTSTRAPPED__";
const LISTENERS_KEY = "__FORMULA_STARTUP_TIMINGS_LISTENERS_INSTALLED__";

const hasTauri = (() => {
  if (hasTauriRuntime()) return true;

  // Fallback: some host environments can delay injecting `__TAURI__` until after the first JS tick.
  // Chromium-based Tauri WebViews typically include "Tauri" in the user agent; use that as a
  // low-risk heuristic so we can still retry listener installation without doing work in normal
  // web builds.
  try {
    const ua = (globalThis as any).navigator?.userAgent;
    return typeof ua === "string" && ua.toLowerCase().includes("tauri");
  } catch {
    return false;
  }
})();

const g = globalThis as any;
if (!g[BOOTSTRAPPED_KEY] && hasTauri) {
  g[BOOTSTRAPPED_KEY] = true;

  // Call immediately (synchronously) to minimize skew for any host-side metrics recorded by this
  // IPC. This may emit `startup:*` events before listeners are registered; we call again after
  // listener installation to re-emit cached timings for late listeners.
  try {
    reportStartupWebviewLoaded();
  } catch {
    // Best-effort; instrumentation should never block startup.
  }

  const ensureListenersInstalled = async (): Promise<boolean> => {
    // Best-effort: Tauri's injected JS APIs may not be immediately available at the earliest
    // point JS can execute (especially in dev / during very early startup). Retry for a short
    // period so we still eventually observe `startup:*` events in the frontend.
    const deadlineMs = Date.now() + 10_000;
    let delayMs = 1;
    while (!g[LISTENERS_KEY] && Date.now() < deadlineMs) {
      try {
        await installStartupTimingsListeners();
      } catch {
        // ignore
      }
      if (g[LISTENERS_KEY]) break;
      await new Promise<void>((resolve) => setTimeout(resolve, delayMs));
      delayMs = Math.min(50, delayMs * 2);
    }
    return Boolean(g[LISTENERS_KEY]);
  };

  void ensureListenersInstalled()
    .then((installed) => {
      if (!installed) return;
      try {
        reportStartupWebviewLoaded();
      } catch {
        // Best-effort; instrumentation should never block startup.
      }
    })
    .catch(() => {
      // Best-effort; instrumentation should never block startup.
    });
}
