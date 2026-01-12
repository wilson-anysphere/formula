import { promises as fs } from "node:fs";
import type { FileHandle } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

export type DataDirLockMetadata = {
  pid: number;
  startedAtMs: number;
  host?: string;
};

export type DataDirLockHandle = {
  lockPath: string;
  fd: FileHandle;
  release: () => Promise<void>;
};

export class DataDirLockError extends Error {
  readonly lockPath: string;
  readonly existingLock?: unknown;

  constructor(message: string, opts: { lockPath: string; existingLock?: unknown }) {
    super(message);
    this.name = "DataDirLockError";
    this.lockPath = opts.lockPath;
    this.existingLock = opts.existingLock;
  }
}

function isFiniteNumber(value: unknown): value is number {
  return typeof value === "number" && Number.isFinite(value);
}

function isPidRunning(pid: number): boolean {
  try {
    process.kill(pid, 0);
    return true;
  } catch (err) {
    const code = (err as NodeJS.ErrnoException).code;
    if (code === "ESRCH") return false;
    // EPERM and other errors indicate the PID exists but we can't signal it.
    return true;
  }
}

async function readExistingLock(lockPath: string): Promise<unknown | null> {
  try {
    const raw = await fs.readFile(lockPath, "utf8");
    return JSON.parse(raw);
  } catch {
    return null;
  }
}

function formatStartedAt(startedAtMs: number | undefined): string | undefined {
  if (startedAtMs === undefined) return undefined;
  try {
    return new Date(startedAtMs).toISOString();
  } catch {
    return undefined;
  }
}

export async function acquireDataDirLock(dataDir: string): Promise<DataDirLockHandle> {
  await fs.mkdir(dataDir, { recursive: true });

  const lockPath = path.join(dataDir, ".sync-server.lock");
  const metadata: DataDirLockMetadata = {
    pid: process.pid,
    startedAtMs: Date.now(),
    host: os.hostname(),
  };

  const createLockFile = async (): Promise<FileHandle> => {
    const fd = await fs.open(lockPath, "wx", 0o600);
    try {
      await fd.writeFile(`${JSON.stringify(metadata)}\n`, "utf8");
      return fd;
    } catch (err) {
      // Best-effort cleanup: if we fail to write lock metadata, don't leave a
      // zero-byte/partial lock file behind that would block subsequent startups.
      try {
        await fd.close();
      } catch {
        // ignore
      }
      await fs.unlink(lockPath).catch((unlinkErr) => {
        const code = (unlinkErr as NodeJS.ErrnoException).code;
        if (code !== "ENOENT") throw unlinkErr;
      });
      throw err;
    }
  };

  try {
    const fd = await createLockFile();
    return {
      lockPath,
      fd,
      async release() {
        try {
          await fd.close();
        } finally {
          await fs.unlink(lockPath).catch((err) => {
            const code = (err as NodeJS.ErrnoException).code;
            if (code !== "ENOENT") throw err;
          });
        }
      },
    };
  } catch (err) {
    const code = (err as NodeJS.ErrnoException).code;
    if (code !== "EEXIST") throw err;

    const existingLock = await readExistingLock(lockPath);
    const host =
      existingLock && typeof existingLock === "object" && "host" in existingLock
        ? (existingLock as { host?: unknown }).host
        : undefined;
    const existingHost =
      typeof host === "string" && host.trim().length > 0 ? host.trim() : undefined;

    const existingPid =
      existingLock && typeof existingLock === "object" && "pid" in existingLock
        ? (existingLock as { pid?: unknown }).pid
        : undefined;

    // Stale-lock cleanup is only safe when the lock was created on this host.
    // If the data dir is on a shared filesystem across machines, a lock from a
    // different host may still be active even if its PID doesn't exist locally.
    const sameHost = existingHost === undefined || existingHost === os.hostname();

    if (isFiniteNumber(existingPid) && sameHost && !isPidRunning(existingPid)) {
      // Stale lock: previous process crashed without cleaning up.
      await fs.unlink(lockPath).catch((err) => {
        const code = (err as NodeJS.ErrnoException).code;
        if (code !== "ENOENT") throw err;
      });

      try {
        const fd = await createLockFile();
        return {
          lockPath,
          fd,
          async release() {
            try {
              await fd.close();
            } finally {
              await fs.unlink(lockPath).catch((err) => {
                const code = (err as NodeJS.ErrnoException).code;
                if (code !== "ENOENT") throw err;
              });
            }
          },
        };
      } catch (err2) {
        const code2 = (err2 as NodeJS.ErrnoException).code;
        if (code2 !== "EEXIST") throw err2;
        // Someone else created the lock file while we were clearing it.
      }
    }

    const startedAtMs =
      existingLock && typeof existingLock === "object" && "startedAtMs" in existingLock
        ? (existingLock as { startedAtMs?: unknown }).startedAtMs
        : undefined;

    const startedAt = isFiniteNumber(startedAtMs) ? formatStartedAt(startedAtMs) : undefined;

    const details = [
      isFiniteNumber(existingPid) ? `pid=${existingPid}` : undefined,
      existingHost ? `host=${existingHost}` : undefined,
      startedAt ? `startedAt=${startedAt}` : undefined,
    ]
      .filter(Boolean)
      .join(" ");

    throw new DataDirLockError(
      [
        `Could not acquire sync-server data directory lock.`,
        `Lock file already exists: ${lockPath}.`,
        details
          ? `It appears another sync-server is running (${details}).`
          : `Another sync-server may already be running.`,
        !sameHost && existingHost
          ? `Lock was created on a different host (${existingHost}). Stale lock cleanup is local-only; stop the other host and delete the lock file manually.`
          : undefined,
        `Stop the other process or choose a different SYNC_SERVER_DATA_DIR.`,
        `If you are sure no other sync-server is using this directory, delete the lock file manually.`,
      ].join(" "),
      { lockPath, existingLock }
    );
  }
}
