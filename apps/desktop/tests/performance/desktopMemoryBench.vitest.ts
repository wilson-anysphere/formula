import { describe, expect, test } from 'vitest';

import { parseProcChildrenPids, parseProcStatusVmRssKb } from './linuxProcUtil.ts';

describe('linuxProcUtil /proc parsing', () => {
  test('parseProcChildrenPids parses an empty children file', () => {
    expect(parseProcChildrenPids('\n')).toEqual([]);
    expect(parseProcChildrenPids('   \n')).toEqual([]);
  });

  test('parseProcChildrenPids parses whitespace-separated pid lists', () => {
    expect(parseProcChildrenPids('123 456 789\n')).toEqual([123, 456, 789]);
    expect(parseProcChildrenPids('  1\t2   3  \n')).toEqual([1, 2, 3]);
  });

  test('parseProcStatusVmRssKb extracts VmRSS from /proc/<pid>/status', () => {
    const status = [
      'Name:\tformula-desktop',
      'Umask:\t0022',
      'State:\tR (running)',
      'VmRSS:\t   42420 kB',
      'Threads:\t12',
      '',
    ].join('\n');

    expect(parseProcStatusVmRssKb(status)).toBe(42420);
  });

  test('parseProcStatusVmRssKb returns null when VmRSS is missing', () => {
    const status = ['Name:\tformula-desktop', 'Threads:\t12', ''].join('\n');
    expect(parseProcStatusVmRssKb(status)).toBeNull();
  });
});
