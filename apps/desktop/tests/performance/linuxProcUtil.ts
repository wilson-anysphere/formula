import { realpathSync } from 'node:fs';
import { readFile, readlink, readdir } from 'node:fs/promises';
import { resolve } from 'node:path';

export function parseProcChildrenPids(content: string): number[] {
  const trimmed = content.trim();
  if (!trimmed) return [];
  return trimmed
    .split(/\s+/g)
    .map((x) => Number(x))
    .filter((n) => Number.isInteger(n) && n > 0);
}

export function parseProcStatusVmRssKb(content: string): number | null {
  const match = content.match(/^VmRSS:\s+(\d+)\s+kB\s*$/m);
  if (!match) return null;
  const kb = Number(match[1]);
  if (!Number.isFinite(kb)) return null;
  return kb;
}

export async function readUtf8(path: string): Promise<string | null> {
  try {
    return await readFile(path, 'utf8');
  } catch (err) {
    const code = (err as NodeJS.ErrnoException).code;
    if (code === 'ENOENT' || code === 'ESRCH' || code === 'EACCES') return null;
    throw err;
  }
}

export async function readProcExeLinux(pid: number): Promise<string | null> {
  try {
    const target = await readlink(`/proc/${pid}/exe`);
    // If the binary was replaced/cleaned up mid-run, Linux appends " (deleted)".
    return target.replace(/ \(deleted\)$/, '');
  } catch (err) {
    const code = (err as NodeJS.ErrnoException).code;
    if (code === 'ENOENT' || code === 'ESRCH' || code === 'EACCES') return null;
    throw err;
  }
}

export async function getChildPidsLinux(pid: number): Promise<number[]> {
  // NOTE: `/proc/<pid>/task/<tid>/children` is per-thread, not per-process. A multi-threaded
  // process can fork from any thread, so union children across all tasks to avoid missing
  // descendants (e.g. WebKit WebView helper processes).
  let tids: string[];
  try {
    tids = await readdir(`/proc/${pid}/task`);
  } catch (err) {
    const code = (err as NodeJS.ErrnoException).code;
    if (code === 'ENOENT' || code === 'ESRCH' || code === 'EACCES') return [];
    throw err;
  }

  const out = new Set<number>();
  for (const tid of tids) {
    const content = await readUtf8(`/proc/${pid}/task/${tid}/children`);
    if (!content) continue;
    for (const child of parseProcChildrenPids(content)) {
      out.add(child);
    }
  }

  return [...out];
}

export async function collectProcessTreePidsLinux(rootPid: number): Promise<number[]> {
  const seen = new Set<number>();
  const stack: number[] = [rootPid];
  while (stack.length > 0) {
    const pid = stack.pop()!;
    if (seen.has(pid)) continue;
    seen.add(pid);
    const children = await getChildPidsLinux(pid);
    for (const child of children) {
      if (!seen.has(child)) stack.push(child);
    }
  }
  return [...seen];
}

async function sleep(ms: number, signal?: AbortSignal): Promise<void> {
  await new Promise<void>((resolvePromise, rejectPromise) => {
    const timer = setTimeout(() => {
      cleanup();
      resolvePromise();
    }, ms);
    const cleanup = () => {
      clearTimeout(timer);
      signal?.removeEventListener('abort', onAbort);
    };
    const onAbort = () => {
      cleanup();
      rejectPromise(new Error('aborted'));
    };
    if (signal) {
      if (signal.aborted) {
        onAbort();
        return;
      }
      signal.addEventListener('abort', onAbort);
    }
  }).catch(() => {
    // Ignore abort errors; callers treat abort as a best-effort early exit.
  });
}

export async function findPidForExecutableLinux(
  rootPid: number,
  binPath: string,
  timeoutMs: number,
  signal?: AbortSignal,
): Promise<number | null> {
  const binResolved = resolve(binPath);
  let binReal = binResolved;
  try {
    binReal = realpathSync(binResolved);
  } catch {
    // Best-effort; realpath can fail in some sandbox setups.
  }

  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    if (signal?.aborted) return null;
    const pids = await collectProcessTreePidsLinux(rootPid);
    for (const pid of pids) {
      const exe = await readProcExeLinux(pid);
      if (!exe) continue;
      if (exe === binReal || exe === binResolved) return pid;
    }
    await sleep(50, signal);
  }
  return null;
}

export async function getProcessRssKbLinux(pid: number): Promise<number | null> {
  const status = await readUtf8(`/proc/${pid}/status`);
  if (!status) return null;
  return parseProcStatusVmRssKb(status);
}

export async function getProcessRssMbLinux(pid: number): Promise<number | null> {
  const kb = await getProcessRssKbLinux(pid);
  if (kb == null) return null;
  return kb / 1024;
}

export async function getProcessRssBytesLinux(pid: number): Promise<number> {
  const kb = await getProcessRssKbLinux(pid);
  if (!kb) return 0;
  return kb * 1024;
}

export async function getProcessTreeRssBytesLinux(rootPid: number): Promise<number> {
  const pids = await collectProcessTreePidsLinux(rootPid);
  let total = 0;
  for (const pid of pids) {
    total += await getProcessRssBytesLinux(pid);
  }
  return total;
}

