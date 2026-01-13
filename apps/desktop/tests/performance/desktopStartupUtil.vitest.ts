import { describe, expect, it } from 'vitest';

import { parseStartupLine } from './desktopStartupUtil.ts';

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

