import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { describe, expect, it } from 'vitest';

import { repoRoot, runOnce } from './desktopStartupUtil.ts';

function isPidAlive(pid: number): boolean {
  try {
    process.kill(pid, 0);
    return true;
  } catch {
    return false;
  }
}

async function waitForPidToExit(pid: number, timeoutMs: number): Promise<void> {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    if (!isPidAlive(pid)) return;
    await new Promise((resolvePromise) => setTimeout(resolvePromise, 20));
  }
  throw new Error(`Timed out waiting for pid ${pid} to exit`);
}

describe('desktopStartupRunnerShared cleanup on early exit', () => {
  it('kills remaining process-group members when the desktop process exits during afterCapture', async () => {
    if (process.platform === 'win32') {
      // This test relies on POSIX process group semantics (`process.kill(-pgid, ...)`).
      return;
    }
    const profileDir = `target/perf-home/vitest-cleanupOnEarlyExit-${Date.now()}-${process.pid}`;
    // `runOnce` resolves `profileDir` relative to the repo root (not the vitest CWD),
    // so compute the pid file path the same way.
    const pidFile = resolve(repoRoot, profileDir, 'grandchild.pid');

    // The spawned "desktop" process (a Node script) prints `[startup] ...` and then exits quickly,
    // leaving a grandchild process alive in the same process group.
    const code = [
      'const { spawn } = require("node:child_process");',
      'const fs = require("node:fs");',
      'const pidFile = process.env.GRANDCHILD_PID_FILE;',
      'const child = spawn(process.execPath, ["-e", "setInterval(() => {}, 1000)"], { stdio: "ignore" });',
      'fs.writeFileSync(pidFile, String(child.pid));',
      'console.log("[startup] window_visible_ms=1 webview_loaded_ms=n/a first_render_ms=n/a tti_ms=2");',
      'process.exit(0);',
    ].join(' ');

    let grandchildPid: number | null = null;
    try {
      const metrics = await runOnce({
        binPath: process.execPath,
        // This harness runs inside the full monorepo Vitest suite where CPU contention can be
        // high (Rust/WASM builds, multiple Vitest workers). Keep this generous to avoid flaky
        // timeouts in CI/sandboxes—this "desktop" process is just a tiny Node script and
        // normally reports metrics almost immediately.
        timeoutMs: 15_000,
        xvfb: false,
        profileDir,
        argv: ['-e', code],
        envOverrides: {
          GRANDCHILD_PID_FILE: pidFile,
        },
        // Keep the afterCapture hook pending long enough that the child process exits before
        // we initiate shutdown, forcing cleanup to happen from the `close` handler.
        afterCaptureTimeoutMs: 200,
        afterCapture: async (_child, _metrics, signal) => {
          await new Promise<void>((resolvePromise) => {
            if (signal.aborted) {
              resolvePromise();
              return;
            }
            const onAbort = () => {
              signal.removeEventListener('abort', onAbort);
              resolvePromise();
            };
            signal.addEventListener('abort', onAbort);
          });
        },
      });

      expect(metrics.windowVisibleMs).toBe(1);
      expect(metrics.ttiMs).toBe(2);

      grandchildPid = Number(readFileSync(pidFile, 'utf8').trim());
      expect(Number.isFinite(grandchildPid)).toBe(true);
      expect(grandchildPid).toBeGreaterThan(0);

      await waitForPidToExit(grandchildPid, 2000);
    } finally {
      if (grandchildPid && isPidAlive(grandchildPid)) {
        try {
          process.kill(grandchildPid, 'SIGKILL');
        } catch {
          // ignore
        }
      }
    }
  });
});
