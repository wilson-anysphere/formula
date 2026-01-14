import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';

import {
  buildDesktopStartupProfileRoot,
  parseStartupLine,
  resolveDesktopStartupArgv,
  resolveDesktopStartupBenchKind,
  resolveDesktopStartupMode,
  resolveDesktopStartupTargets,
} from './desktopStartupUtil.ts';

describe('desktopStartupUtil parseStartupLine', () => {
  it('parses the legacy startup metrics line (no first_render_ms)', () => {
    expect(
      parseStartupLine('[startup] window_visible_ms=123 webview_loaded_ms=234 tti_ms=456'),
    ).toEqual({
      windowVisibleMs: 123,
      webviewLoadedMs: 234,
      firstRenderMs: null,
      ttiMs: 456,
    });
  });

  it('parses the full startup metrics line (including first_render_ms)', () => {
    expect(
      parseStartupLine(
        '[startup] window_visible_ms=10 webview_loaded_ms=20 first_render_ms=30 tti_ms=40',
      ),
    ).toEqual({
      windowVisibleMs: 10,
      webviewLoadedMs: 20,
      firstRenderMs: 30,
      ttiMs: 40,
    });
  });

  it('treats n/a fields as null', () => {
    expect(
      parseStartupLine(
        '[startup] window_visible_ms=10 webview_loaded_ms=n/a first_render_ms=n/a tti_ms=40',
      ),
    ).toEqual({
      windowVisibleMs: 10,
      webviewLoadedMs: null,
      firstRenderMs: null,
      ttiMs: 40,
    });
  });

  it('parses when [startup] is embedded in a longer log line', () => {
    expect(
      parseStartupLine(
        '2026-01-01T00:00:00.000Z INFO something [startup] window_visible_ms=1 webview_loaded_ms=n/a tti_ms=2',
      ),
    ).toEqual({
      windowVisibleMs: 1,
      webviewLoadedMs: null,
      firstRenderMs: null,
      ttiMs: 2,
    });
  });

  it('returns null when required keys are missing', () => {
    expect(parseStartupLine('[startup] window_visible_ms=123')).toBeNull();
    expect(parseStartupLine('[startup] tti_ms=456')).toBeNull();
  });

  it('returns null when numeric values are invalid', () => {
    expect(parseStartupLine('[startup] window_visible_ms=abc tti_ms=456')).toBeNull();
    expect(parseStartupLine('[startup] window_visible_ms=123 tti_ms=abc')).toBeNull();
    expect(parseStartupLine('[startup] window_visible_ms=123 webview_loaded_ms=abc tti_ms=456')).toBeNull();
    expect(parseStartupLine('[startup] window_visible_ms=123 first_render_ms=abc tti_ms=456')).toBeNull();
  });
});

describe('desktopStartupUtil resolveDesktopStartupTargets', () => {
  it('uses default targets when env is empty (full/cold)', () => {
    expect(
      resolveDesktopStartupTargets({ benchKind: 'full', mode: 'cold', env: {} }),
    ).toEqual({
      windowVisibleTargetMs: 500,
      webviewLoadedTargetMs: 800,
      firstRenderTargetMs: 500,
      ttiTargetMs: 1000,
    });
  });

  it('respects full warm overrides (falling back to cold when missing)', () => {
    expect(
      resolveDesktopStartupTargets({
        benchKind: 'full',
        mode: 'warm',
        env: {
          FORMULA_DESKTOP_COLD_WINDOW_VISIBLE_TARGET_MS: '111',
          FORMULA_DESKTOP_WARM_WINDOW_VISIBLE_TARGET_MS: '222',
          FORMULA_DESKTOP_COLD_TTI_TARGET_MS: '333',
          // warm TTI missing -> fall back to cold
          FORMULA_DESKTOP_COLD_FIRST_RENDER_TARGET_MS: '444',
          FORMULA_DESKTOP_WARM_FIRST_RENDER_TARGET_MS: '555',
          FORMULA_DESKTOP_WEBVIEW_LOADED_TARGET_MS: '666',
        },
      }),
    ).toEqual({
      windowVisibleTargetMs: 222,
      webviewLoadedTargetMs: 666,
      firstRenderTargetMs: 555,
      ttiTargetMs: 333,
    });
  });

  it('uses shell targets (falling back to full targets)', () => {
    expect(
      resolveDesktopStartupTargets({
        benchKind: 'shell',
        mode: 'cold',
        env: {
          FORMULA_DESKTOP_COLD_WINDOW_VISIBLE_TARGET_MS: '111',
          FORMULA_DESKTOP_COLD_TTI_TARGET_MS: '222',
          FORMULA_DESKTOP_SHELL_COLD_WINDOW_VISIBLE_TARGET_MS: '333',
          FORMULA_DESKTOP_SHELL_COLD_TTI_TARGET_MS: '444',
          FORMULA_DESKTOP_WEBVIEW_LOADED_TARGET_MS: '555',
          // shell webview target missing -> fall back to full
        },
      }),
    ).toEqual({
      windowVisibleTargetMs: 333,
      webviewLoadedTargetMs: 555,
      firstRenderTargetMs: 500,
      ttiTargetMs: 444,
    });
  });

  it('ignores invalid/zero targets and falls back to defaults', () => {
    expect(
      resolveDesktopStartupTargets({
        benchKind: 'full',
        mode: 'cold',
        env: {
          FORMULA_DESKTOP_COLD_WINDOW_VISIBLE_TARGET_MS: '0',
          FORMULA_DESKTOP_COLD_TTI_TARGET_MS: '-1',
          FORMULA_DESKTOP_COLD_FIRST_RENDER_TARGET_MS: 'wat',
          FORMULA_DESKTOP_WEBVIEW_LOADED_TARGET_MS: '',
        },
      }),
    ).toEqual({
      windowVisibleTargetMs: 500,
      webviewLoadedTargetMs: 800,
      firstRenderTargetMs: 500,
      ttiTargetMs: 1000,
    });
  });
});

describe('desktopStartupUtil resolveDesktopStartupMode', () => {
  it('defaults to cold', () => {
    expect(resolveDesktopStartupMode({ env: {} })).toBe('cold');
  });

  it('parses warm', () => {
    expect(resolveDesktopStartupMode({ env: { FORMULA_DESKTOP_STARTUP_MODE: 'warm' } })).toBe('warm');
  });

  it('throws on invalid mode', () => {
    expect(() => resolveDesktopStartupMode({ env: { FORMULA_DESKTOP_STARTUP_MODE: 'wat' } })).toThrow(
      /FORMULA_DESKTOP_STARTUP_MODE/,
    );
  });
});

describe('desktopStartupUtil resolveDesktopStartupBenchKind', () => {
  it('defaults to full when env is empty and CI is unset', () => {
    expect(resolveDesktopStartupBenchKind({ env: {} })).toBe('full');
  });

  it('defaults to shell when CI is set', () => {
    expect(resolveDesktopStartupBenchKind({ env: { CI: '1' } })).toBe('shell');
  });

  it('uses the provided default when env is empty', () => {
    expect(resolveDesktopStartupBenchKind({ env: {}, defaultKind: 'shell' })).toBe('shell');
  });

  it('parses full', () => {
    expect(resolveDesktopStartupBenchKind({ env: { FORMULA_DESKTOP_STARTUP_BENCH_KIND: 'full' }, defaultKind: 'shell' })).toBe(
      'full',
    );
  });

  it('throws on invalid bench kind', () => {
    expect(() =>
      resolveDesktopStartupBenchKind({ env: { FORMULA_DESKTOP_STARTUP_BENCH_KIND: 'wat' }, defaultKind: 'full' }),
    ).toThrow(/FORMULA_DESKTOP_STARTUP_BENCH_KIND/);
  });
});

describe('desktopStartupUtil resolveDesktopStartupArgv', () => {
  it('returns --startup-bench args for shell', () => {
    expect(resolveDesktopStartupArgv('shell')).toEqual(['--startup-bench']);
  });

  it('returns no args for full', () => {
    expect(resolveDesktopStartupArgv('full')).toEqual([]);
  });
});

describe('desktopStartupUtil buildDesktopStartupProfileRoot', () => {
  it('names profile roots consistently', () => {
    const perfHome = resolve('target', 'perf-home');
    expect(
      buildDesktopStartupProfileRoot({ perfHome, benchKind: 'shell', mode: 'cold', now: 123, pid: 456 }),
    ).toBe(resolve(perfHome, 'desktop-startup-shell-cold-123-456'));
  });
});
