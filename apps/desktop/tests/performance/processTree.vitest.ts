import { describe, expect, it, vi } from 'vitest';

import { buildTaskkillCommand, terminateProcessTree } from './processTree.ts';

describe('terminateProcessTree', () => {
  it('buildTaskkillCommand uses /PID <pid> /T /F', () => {
    expect(buildTaskkillCommand(1234)).toEqual({
      command: 'taskkill',
      args: ['/PID', '1234', '/T', '/F'],
    });
  });

  it('kills the POSIX process group (SIGTERM) in graceful mode', () => {
    const processKill = vi.fn();
    const childKill = vi.fn();

    terminateProcessTree({ pid: 4242, kill: childKill } as any, 'graceful', {
      platform: 'linux',
      processKill,
    });

    expect(processKill).toHaveBeenCalledTimes(1);
    expect(processKill).toHaveBeenCalledWith(-4242, 'SIGTERM');
    expect(childKill).not.toHaveBeenCalled();
  });

  it('kills the POSIX process group (SIGKILL) in force mode', () => {
    const processKill = vi.fn();

    terminateProcessTree({ pid: 4242, kill: vi.fn() } as any, 'force', {
      platform: 'darwin',
      processKill,
    });

    expect(processKill).toHaveBeenCalledTimes(1);
    expect(processKill).toHaveBeenCalledWith(-4242, 'SIGKILL');
  });

  it('falls back to killing just the root pid when group kill fails', () => {
    const calls: Array<[number, string]> = [];
    const processKill = vi.fn((pid: number, signal: any) => {
      calls.push([pid, signal]);
      if (pid < 0) throw new Error('no such process group');
    });

    terminateProcessTree({ pid: 4242, kill: vi.fn() } as any, 'force', {
      platform: 'linux',
      processKill,
    });

    expect(calls).toEqual([
      [-4242, 'SIGKILL'],
      [4242, 'SIGKILL'],
    ]);
  });

  it('uses taskkill on Windows in force mode (bounded timeout)', () => {
    const spawnSync = vi.fn(() => ({ status: 0 }) as any);
    const childKill = vi.fn();

    terminateProcessTree({ pid: 9001, kill: childKill } as any, 'force', {
      platform: 'win32',
      spawnSync,
      taskkillTimeoutMs: 1234,
    });

    expect(spawnSync).toHaveBeenCalledTimes(1);
    expect(spawnSync).toHaveBeenCalledWith(
      'taskkill',
      ['/PID', '9001', '/T', '/F'],
      expect.objectContaining({ stdio: 'ignore', windowsHide: true, timeout: 1234 }),
    );

    // If taskkill was invoked successfully, we shouldn't need to fall back to killing just the root.
    expect(childKill).not.toHaveBeenCalled();
  });

  it('does not use taskkill on Windows in graceful mode', () => {
    const spawnSync = vi.fn(() => ({ status: 0 }) as any);
    const childKill = vi.fn();

    terminateProcessTree({ pid: 9001, kill: childKill } as any, 'graceful', {
      platform: 'win32',
      spawnSync,
    });

    expect(spawnSync).not.toHaveBeenCalled();
    expect(childKill).toHaveBeenCalledTimes(1);
  });

  it('falls back to killing just the root pid on Windows if taskkill throws', () => {
    const spawnSync = vi.fn(() => {
      throw new Error('taskkill unavailable');
    });
    const childKill = vi.fn();

    terminateProcessTree({ pid: 9001, kill: childKill } as any, 'force', {
      platform: 'win32',
      spawnSync,
    });

    expect(spawnSync).toHaveBeenCalledTimes(1);
    expect(childKill).toHaveBeenCalledTimes(1);
  });

  it('falls back to killing just the root pid on Windows if taskkill exits non-zero', () => {
    const spawnSync = vi.fn(() => ({ status: 1, error: undefined }) as any);
    const childKill = vi.fn();

    terminateProcessTree({ pid: 9001, kill: childKill } as any, 'force', {
      platform: 'win32',
      spawnSync,
    });

    expect(spawnSync).toHaveBeenCalledTimes(1);
    expect(childKill).toHaveBeenCalledTimes(1);
  });

  it('no-ops when pid is missing', () => {
    const spawnSync = vi.fn();
    const processKill = vi.fn();

    terminateProcessTree({ pid: undefined, kill: vi.fn() } as any, 'force', {
      platform: 'win32',
      spawnSync,
      processKill,
    });

    expect(spawnSync).not.toHaveBeenCalled();
    expect(processKill).not.toHaveBeenCalled();
  });
});
