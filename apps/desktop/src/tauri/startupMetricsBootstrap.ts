import { installStartupTimingsListeners, reportStartupWebviewLoaded } from "./startupMetrics.js";

// Startup performance instrumentation (no-op for web builds).
//
// Call `reportStartupWebviewLoaded()` at the earliest point in the module graph so the
// host-side `webview_loaded_ms` measurement does not include startup JS work (including
// listener-install IPC overhead). Once the listeners are installed, call it again
// (idempotent) to re-emit the timing events for the now-ready listeners.
try {
  reportStartupWebviewLoaded();
} catch {
  // Best-effort; instrumentation should never block startup.
}

void installStartupTimingsListeners()
  .catch(() => {
    // Best-effort; instrumentation should never block startup.
  })
  .finally(() => {
    try {
      reportStartupWebviewLoaded();
    } catch {
      // Best-effort; instrumentation should never block startup.
    }
  });

