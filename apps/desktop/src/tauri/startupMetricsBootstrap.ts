import { installStartupTimingsListeners, reportStartupWebviewLoaded } from "./startupMetrics.js";

// Startup performance instrumentation (no-op for web builds).
//
// `webviewLoadedMs` is recorded natively by the Rust host when the main WebView finishes its
// initial navigation. Tauri does not guarantee events are queued before listeners are installed,
// so we install listeners early and then ask the host to (re-)emit the cached timings once ready.

// Call immediately (synchronously) to minimize skew for any host-side metrics recorded by this
// IPC. This may emit `startup:*` events before listeners are registered; we call again after
// listener installation to re-emit cached timings for late listeners.
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
