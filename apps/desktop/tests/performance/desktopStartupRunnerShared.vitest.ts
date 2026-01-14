import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { afterEach, describe, expect, test, vi } from 'vitest';

import { parseStartupLine, repoRoot, runOnce } from './desktopStartupUtil.ts';

describe('desktopStartupUtil.parseStartupLine', () => {
  test('parses a full startup metrics line', () => {
    expect(
      parseStartupLine(
        '[startup] window_visible_ms=123 webview_loaded_ms=234 first_render_ms=345 tti_ms=456',
      ),
    ).toEqual({
      windowVisibleMs: 123,
      webviewLoadedMs: 234,
      firstRenderMs: 345,
      ttiMs: 456,
    });
  });

  test('parses webview_loaded_ms=n/a as null', () => {
    expect(
      parseStartupLine(
        '[startup] window_visible_ms=1 webview_loaded_ms=n/a first_render_ms=3 tti_ms=4',
      ),
    ).toEqual({
      windowVisibleMs: 1,
      webviewLoadedMs: null,
      firstRenderMs: 3,
      ttiMs: 4,
    });
  });

  test('parses first_render_ms=n/a as null', () => {
    expect(
      parseStartupLine(
        '[startup] window_visible_ms=10 webview_loaded_ms=20 first_render_ms=n/a tti_ms=40',
      ),
    ).toEqual({
      windowVisibleMs: 10,
      webviewLoadedMs: 20,
      firstRenderMs: null,
      ttiMs: 40,
    });
  });

  test('parses startup metrics embedded in a longer log line', () => {
    expect(
      parseStartupLine(
        '[tauri] info: ready [startup] window_visible_ms=7 webview_loaded_ms=8 first_render_ms=9 tti_ms=10',
      ),
    ).toEqual({
      windowVisibleMs: 7,
      webviewLoadedMs: 8,
      firstRenderMs: 9,
      ttiMs: 10,
    });
  });

  test('returns null when required fields are missing', () => {
    expect(
      parseStartupLine('[startup] window_visible_ms=1 webview_loaded_ms=2 first_render_ms=3'),
    ).toBeNull();
    expect(parseStartupLine('[startup] tti_ms=1 webview_loaded_ms=2 first_render_ms=3')).toBeNull();
  });

  test('returns null when required numeric fields are invalid', () => {
    expect(
      parseStartupLine(
        '[startup] window_visible_ms=abc webview_loaded_ms=2 first_render_ms=3 tti_ms=4',
      ),
    ).toBeNull();
    expect(
      parseStartupLine(
        '[startup] window_visible_ms=1 webview_loaded_ms=2 first_render_ms=3 tti_ms=wat',
      ),
    ).toBeNull();
  });
});

describe('desktopStartupUtil.runOnce env isolation', () => {
  test('redirects HOME and user-data directories under the profile dir', async () => {
    const profileDir = `target/perf-home/vitest-env-isolation-${Date.now()}-${process.pid}`;
    const profileAbs = resolve(repoRoot, profileDir);
    const outFile = resolve(profileAbs, 'env.json');

    const code = [
      "const fs = require('node:fs');",
      "const out = {",
      "  HOME: process.env.HOME,",
      "  USERPROFILE: process.env.USERPROFILE,",
      "  XDG_CONFIG_HOME: process.env.XDG_CONFIG_HOME,",
      "  XDG_CACHE_HOME: process.env.XDG_CACHE_HOME,",
      "  XDG_STATE_HOME: process.env.XDG_STATE_HOME,",
      "  XDG_DATA_HOME: process.env.XDG_DATA_HOME,",
      "  APPDATA: process.env.APPDATA,",
      "  LOCALAPPDATA: process.env.LOCALAPPDATA,",
      "  TMPDIR: process.env.TMPDIR,",
      "  TEMP: process.env.TEMP,",
      "  TMP: process.env.TMP,",
      "  FORMULA_PERF_HOME: process.env.FORMULA_PERF_HOME,",
      "};",
      "fs.writeFileSync(process.env.ENV_OUT_FILE, JSON.stringify(out), 'utf8');",
      "console.log('[startup] window_visible_ms=1 webview_loaded_ms=n/a first_render_ms=n/a tti_ms=2');",
      'setInterval(() => {}, 1000);',
    ].join(' ');

    await runOnce({
      binPath: process.execPath,
      timeoutMs: 5000,
      xvfb: false,
      profileDir,
      argv: ['-e', code],
      envOverrides: {
        ENV_OUT_FILE: outFile,
      },
    });

    const payload = JSON.parse(readFileSync(outFile, 'utf8')) as Record<string, string | undefined>;

    expect(payload.HOME).toBe(profileAbs);
    expect(payload.USERPROFILE).toBe(profileAbs);

    expect(payload.XDG_CONFIG_HOME).toBe(resolve(profileAbs, 'xdg-config'));
    expect(payload.XDG_CACHE_HOME).toBe(resolve(profileAbs, 'xdg-cache'));
    expect(payload.XDG_STATE_HOME).toBe(resolve(profileAbs, 'xdg-state'));
    expect(payload.XDG_DATA_HOME).toBe(resolve(profileAbs, 'xdg-data'));

    expect(payload.APPDATA).toBe(resolve(profileAbs, 'AppData', 'Roaming'));
    expect(payload.LOCALAPPDATA).toBe(resolve(profileAbs, 'AppData', 'Local'));

    expect(payload.TMPDIR).toBe(resolve(profileAbs, 'tmp'));
    expect(payload.TEMP).toBe(resolve(profileAbs, 'tmp'));
    expect(payload.TMP).toBe(resolve(profileAbs, 'tmp'));

    // The perf runner exposes the resolved perf root so downstream tooling can locate artifacts.
    expect(payload.FORMULA_PERF_HOME).toBe(resolve(repoRoot, 'target', 'perf-home'));
  });
});

describe('desktopStartupUtil.runOnce reset guardrails', () => {
  const prevResetHome = process.env.FORMULA_DESKTOP_BENCH_RESET_HOME;
  const prevPerfHome = process.env.FORMULA_PERF_HOME;

  afterEach(() => {
    if (prevResetHome === undefined) {
      delete process.env.FORMULA_DESKTOP_BENCH_RESET_HOME;
    } else {
      process.env.FORMULA_DESKTOP_BENCH_RESET_HOME = prevResetHome;
    }

    if (prevPerfHome === undefined) {
      delete process.env.FORMULA_PERF_HOME;
    } else {
      process.env.FORMULA_PERF_HOME = prevPerfHome;
    }
  });

  test('refuses to reset when profileDir resolves to filesystem root', async () => {
    process.env.FORMULA_DESKTOP_BENCH_RESET_HOME = '1';
    await expect(runOnce({ binPath: 'ignored', timeoutMs: 1, profileDir: '/' })).rejects.toThrow(
      /Refusing to reset unsafe desktop benchmark profile dir/,
    );
  });

  test('refuses to reset when profileDir resolves to the repo root', async () => {
    process.env.FORMULA_DESKTOP_BENCH_RESET_HOME = '1';
    await expect(
      runOnce({ binPath: 'ignored', timeoutMs: 1, profileDir: repoRoot }),
    ).rejects.toThrow(/Refusing to reset unsafe desktop benchmark profile dir/);
  });

  test('refuses to reset when FORMULA_PERF_HOME resolves to filesystem root', async () => {
    process.env.FORMULA_DESKTOP_BENCH_RESET_HOME = '1';
    process.env.FORMULA_PERF_HOME = '/';

    vi.resetModules();
    const mod = await import('./desktopStartupUtil.ts');
    await expect(mod.runOnce({ binPath: 'ignored', timeoutMs: 1, profileDir: '/tmp' })).rejects.toThrow(
      /Refusing to reset unsafe desktop benchmark perf home dir/,
    );
  });

  test('refuses to reset when FORMULA_PERF_HOME resolves to target/', async () => {
    process.env.FORMULA_DESKTOP_BENCH_RESET_HOME = '1';
    process.env.FORMULA_PERF_HOME = 'target';

    await expect(runOnce({ binPath: 'ignored', timeoutMs: 1 })).rejects.toThrow(
      /Refusing to reset FORMULA_PERF_HOME=.*target/,
    );
  });
});

describe('desktopStartupUtil.runOnce failure diagnostics', () => {
  test('includes recent stdout/stderr when timing out waiting for metrics', async () => {
    const code = [
      "console.log('hello stdout');",
      "console.error('hello stderr');",
      // Emit output periodically to make this test resilient under heavy load (when the child
      // process may start slowly and miss a narrow timeout window).
      "setInterval(() => { console.log('hello stdout'); console.error('hello stderr'); }, 250);",
    ].join(' ');

    try {
      await runOnce({
        binPath: process.execPath,
        timeoutMs: 5000,
        xvfb: false,
        argv: ['-e', code],
      });
      throw new Error('expected runOnce to time out');
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      expect(msg).toContain('Timed out after 5000ms waiting for startup metrics');
      expect(msg).toContain('desktop process output');
      expect(msg).toContain('hello stdout');
      expect(msg).toContain('hello stderr');
    }
  });

  test('includes recent stdout/stderr when the desktop process exits early', async () => {
    const code = [
      "console.log('goodbye stdout');",
      "console.error('goodbye stderr');",
      // Give stdout/stderr a moment to flush before exiting.
      'setTimeout(() => process.exit(1), 50);',
      'setInterval(() => {}, 1000);',
    ].join(' ');

    try {
      await runOnce({
        binPath: process.execPath,
        timeoutMs: 20000,
        xvfb: false,
        argv: ['-e', code],
      });
      throw new Error('expected runOnce to reject');
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      expect(msg).toContain('Desktop process exited before reporting metrics');
      expect(msg).toContain('desktop process output');
      expect(msg).toContain('goodbye stdout');
      expect(msg).toContain('goodbye stderr');
    }
  });
});
