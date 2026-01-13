import { describe, expect, test } from 'vitest';

import { parseStartupLine } from './desktopStartupRunnerShared.ts';

describe('desktopStartupRunnerShared.parseStartupLine', () => {
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

