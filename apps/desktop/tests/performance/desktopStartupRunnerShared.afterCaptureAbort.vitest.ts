import { describe, expect, it } from 'vitest';

import { runOnce } from './desktopStartupRunnerShared.ts';

describe('desktopStartupRunnerShared afterCapture abort', () => {
  it('aborts the afterCapture signal when afterCaptureTimeoutMs elapses', async () => {
    const prevCi = process.env.CI;
    const prevDisplay = process.env.DISPLAY;
    const prevResetHome = process.env.FORMULA_DESKTOP_BENCH_RESET_HOME;

    // Ensure `shouldUseXvfb()` does not force the xvfb wrapper (which would require Xvfb to be
    // installed for unit tests).
    delete process.env.CI;
    process.env.DISPLAY = process.env.DISPLAY || ':99';
    delete process.env.FORMULA_DESKTOP_BENCH_RESET_HOME;

    try {
      let sawAbort = false;
      const metrics = await runOnce({
        binPath: process.execPath,
        timeoutMs: 5000,
        profileDir: `target/perf-home/vitest-afterCaptureAbort-${Date.now()}`,
        argv: [
          '-e',
          [
            'console.log("[startup] window_visible_ms=1 webview_loaded_ms=n/a first_render_ms=n/a tti_ms=2");',
            'setInterval(() => {}, 1000);',
          ].join(' '),
        ],
        afterCaptureTimeoutMs: 50,
        afterCapture: async (_child, _metrics, signal) => {
          await new Promise<void>((resolvePromise) => {
            if (signal.aborted) {
              sawAbort = true;
              resolvePromise();
              return;
            }
            const onAbort = () => {
              sawAbort = true;
              signal.removeEventListener('abort', onAbort);
              resolvePromise();
            };
            signal.addEventListener('abort', onAbort);
          });
        },
      });

      expect(metrics.windowVisibleMs).toBe(1);
      expect(metrics.ttiMs).toBe(2);
      expect(sawAbort).toBe(true);
    } finally {
      if (prevCi === undefined) delete process.env.CI;
      else process.env.CI = prevCi;

      if (prevDisplay === undefined) delete process.env.DISPLAY;
      else process.env.DISPLAY = prevDisplay;

      if (prevResetHome === undefined) delete process.env.FORMULA_DESKTOP_BENCH_RESET_HOME;
      else process.env.FORMULA_DESKTOP_BENCH_RESET_HOME = prevResetHome;
    }
  });
});

