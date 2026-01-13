import { installStartupTimingsListeners, reportStartupWebviewLoaded } from "./startupMetrics.js";

// Startup performance instrumentation (no-op for web builds).
//
// `webviewLoadedMs` is recorded natively by the Rust host when the main WebView finishes its
// initial navigation. Tauri does not guarantee events are queued before listeners are installed,
// so we install listeners early and then ask the host to (re-)emit the cached timings once ready.

// Send the "webview loaded" IPC as early as possible. This ensures the Rust host can capture a
// reasonable "JS started executing" marker even if the rest of the module graph takes time to load.
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
