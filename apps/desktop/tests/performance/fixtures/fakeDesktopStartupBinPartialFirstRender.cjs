#!/usr/bin/env node

// Fake "desktop" binary used by unit tests for the standalone startup runner.
//
// Unlike `fakeDesktopStartupModule.cjs` (which is loaded via NODE_OPTIONS), this file is executed
// directly as the "desktop binary" so it can accept arbitrary CLI args like `--startup-bench`.
//
// Behavior:
// - Prints a `[startup] ...` metrics line immediately.
// - Includes `first_render_ms` only for the first cold-run profile directory (`run-01`).
//   For all other runs, it omits `first_render_ms` entirely to simulate older binaries / IPC flakiness.
// - Stays alive until the benchmark harness terminates it.

const path = require("node:path");

const home = process.env.HOME || process.env.USERPROFILE || process.cwd();
const profileBase = path.basename(home);

const windowVisibleMs = 100;
const webviewLoadedMs = 200;
const ttiMs = 400;

if (profileBase === "run-01") {
  // eslint-disable-next-line no-console
  console.log(
    `[startup] window_visible_ms=${windowVisibleMs} webview_loaded_ms=${webviewLoadedMs} first_render_ms=300 tti_ms=${ttiMs}`,
  );
} else {
  // eslint-disable-next-line no-console
  console.log(
    `[startup] window_visible_ms=${windowVisibleMs} webview_loaded_ms=${webviewLoadedMs} tti_ms=${ttiMs}`,
  );
}

setInterval(() => {}, 1000);

