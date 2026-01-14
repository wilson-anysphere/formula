import { resolve } from 'node:path';

import { describe, expect, it } from 'vitest';

import { resolveDesktopMemoryBenchEnv } from './desktopStartupUtil.ts';
import { repoRoot } from './desktopStartupUtil.ts';

describe('desktopMemoryUtil resolveDesktopMemoryBenchEnv', () => {
  it('uses defaults when env is empty', () => {
    expect(resolveDesktopMemoryBenchEnv({ env: {} })).toEqual({
      runs: 10,
      timeoutMs: 20_000,
      settleMs: 5_000,
      targetMb: 100,
      enforce: false,
      binPath: null,
    });
  });

  it('parses overrides and resolves FORMULA_DESKTOP_BIN', () => {
    expect(
      resolveDesktopMemoryBenchEnv({
        env: {
          FORMULA_DESKTOP_MEMORY_RUNS: '5',
          FORMULA_DESKTOP_MEMORY_TIMEOUT_MS: '123',
          FORMULA_DESKTOP_MEMORY_SETTLE_MS: '0',
          FORMULA_DESKTOP_IDLE_RSS_TARGET_MB: '250',
          FORMULA_ENFORCE_DESKTOP_MEMORY_BENCH: '1',
          FORMULA_DESKTOP_BIN: '/tmp/formula-desktop',
        },
      }),
    ).toEqual({
      runs: 5,
      timeoutMs: 123,
      settleMs: 0,
      targetMb: 250,
      enforce: true,
      binPath: '/tmp/formula-desktop',
    });
  });

  it('resolves relative FORMULA_DESKTOP_BIN paths from the repo root', () => {
    expect(
      resolveDesktopMemoryBenchEnv({
        env: {
          FORMULA_DESKTOP_BIN: 'target/release/formula-desktop',
        },
      }).binPath,
    ).toBe(resolve(repoRoot, 'target/release/formula-desktop'));
  });

  it('falls back when values are invalid or non-positive', () => {
    expect(
      resolveDesktopMemoryBenchEnv({
        env: {
          FORMULA_DESKTOP_MEMORY_RUNS: '0',
          FORMULA_DESKTOP_MEMORY_TIMEOUT_MS: '-1',
          FORMULA_DESKTOP_MEMORY_SETTLE_MS: '-5',
          FORMULA_DESKTOP_IDLE_RSS_TARGET_MB: '-10',
          FORMULA_DESKTOP_MEMORY_TARGET_MB: '200',
          FORMULA_ENFORCE_DESKTOP_MEMORY_BENCH: '0',
          FORMULA_DESKTOP_BIN: '   ',
        },
      }),
    ).toEqual({
      runs: 10,
      timeoutMs: 20_000,
      settleMs: 5_000,
      targetMb: 200,
      enforce: false,
      binPath: null,
    });
  });
});
