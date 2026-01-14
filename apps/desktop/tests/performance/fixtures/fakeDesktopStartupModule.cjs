const fs = require('node:fs');
const path = require('node:path');

// This module is loaded via `NODE_OPTIONS=--require ...` in unit tests so we can exercise the
// desktop startup benchmark harness without building the real Tauri binary.
//
// The benchmark runner redirects HOME/USERPROFILE (and XDG dirs on Linux) to an isolated profile
// dir. We use a marker file under HOME to simulate a persistent cache:
// - First launch (no marker): print "cold" timings and create the marker file.
// - Subsequent launches (marker exists): print "warm" timings.
//
// The process stays alive until the benchmark harness terminates it (so RSS sampling can run).

const home = process.env.HOME || process.env.USERPROFILE || process.cwd();
const markerPath = path.join(home, 'startup-bench-marker.txt');
const isWarm = fs.existsSync(markerPath);

const windowVisibleMs = isWarm ? 10 : 100;
const webviewLoadedMs = isWarm ? 20 : 200;
const firstRenderMs = isWarm ? 30 : 300;
const ttiMs = isWarm ? 40 : 400;

try {
  fs.mkdirSync(home, { recursive: true });
  fs.writeFileSync(markerPath, '1', 'utf8');
} catch {
  // Best-effort: the benchmark harness still needs to run even if we fail to write the marker.
}

// Match the Rust-side log format parsed by `desktopStartupRunnerShared.parseStartupLine`.
// eslint-disable-next-line no-console
console.log(
  `[startup] window_visible_ms=${windowVisibleMs} webview_loaded_ms=${webviewLoadedMs} first_render_ms=${firstRenderMs} tti_ms=${ttiMs}`,
);

// Keep the process alive until the benchmark harness terminates it.
setInterval(() => {}, 1000);

