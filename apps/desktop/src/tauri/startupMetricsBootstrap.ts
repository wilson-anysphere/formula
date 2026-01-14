import { installStartupTimingsListeners, reportStartupWebviewLoaded } from "./startupMetrics.js";
import { getTauriInvokeOrNull, hasTauri as hasTauriRuntime } from "./api";

// Startup performance instrumentation (no-op for web builds).
//
// `webviewLoadedMs` is recorded natively by the Rust host when the main WebView finishes its
// initial navigation. Tauri does not guarantee events are queued before listeners are installed,
// so we install listeners early and then ask the host to (re-)emit the cached timings once ready.

const BOOTSTRAPPED_KEY = "__FORMULA_STARTUP_METRICS_BOOTSTRAPPED__";
const LISTENERS_KEY = "__FORMULA_STARTUP_TIMINGS_LISTENERS_INSTALLED__";
const WEBVIEW_REPORTED_KEY = "__FORMULA_STARTUP_WEBVIEW_LOADED_REPORTED__";

const hasTauri = (() => {
  if (hasTauriRuntime()) return true;

  // If accessing `__TAURI__` throws (e.g. hardened environment or tests), treat that as "not Tauri"
  // and skip all bootstrap work. We intentionally avoid falling back to the user-agent heuristic in
  // this case to keep behavior a no-op outside of real desktop builds.
  try {
    // eslint-disable-next-line @typescript-eslint/no-unused-expressions
    (globalThis as any).__TAURI__;
  } catch {
    return false;
  }

  // Packaged desktop builds typically run on the `tauri://` protocol, even before the JS bridge
  // has finished injecting `__TAURI__`. Use that as a stable signal so we still bootstrap in
  // production builds even if the user agent does not include "Tauri".
  try {
    const protocol = (globalThis as any).location?.protocol;
    if (protocol === "tauri:" || protocol === "asset:") return true;
  } catch {
    // ignore
  }

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
  try {
    if (getTauriInvokeOrNull()) g[WEBVIEW_REPORTED_KEY] = true;
  } catch {
    // Ignore: if we can't probe the global, we'll rely on the retry loop below.
  }

  // Best-effort: Tauri's injected JS APIs may not be immediately available at the earliest
  // point JS can execute (especially in dev / during very early startup). Retry for a short
  // period so we still eventually observe `startup:*` events in the frontend.
  //
  // Note: this retry loop is intentionally timer-driven (rather than `await`ing at the top of
  // an async loop). This ensures environments using fake timers (tests) still advance the poller
  // deterministically even when promise microtasks are only flushed after `advanceTimersByTime`.
  const deadlineMs = Date.now() + 10_000;
  let delayMs = 1;
  const tick = (): void => {
    if (g[LISTENERS_KEY]) return;
    if (Date.now() >= deadlineMs) return;

    // If the `core.invoke` binding becomes available after the first JS tick, send a best-effort
    // report as soon as possible (still re-emitting again once listeners are installed).
    if (!g[WEBVIEW_REPORTED_KEY]) {
      if (getTauriInvokeOrNull()) {
        try {
          reportStartupWebviewLoaded();
        } catch {
          // ignore
        }
        g[WEBVIEW_REPORTED_KEY] = true;
      }
    }

    try {
      // Fire-and-forget: `installStartupTimingsListeners` sets the global flag synchronously
      // once the event API is available, and it catches individual listener failures.
      void installStartupTimingsListeners().catch(() => {});
    } catch {
      // ignore
    }

    if (g[LISTENERS_KEY]) {
      // Re-emit cached timings now that listeners are installed.
      try {
        reportStartupWebviewLoaded();
      } catch {
        // Best-effort; instrumentation should never block startup.
      }
      return;
    }

    const nextDelay = delayMs;
    delayMs = Math.min(50, delayMs * 2);
    setTimeout(tick, nextDelay);
  };

  tick();
}
