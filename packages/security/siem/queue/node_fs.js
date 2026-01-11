import {
  mkdir,
  open as openFile,
  readdir,
  readFile,
  rename,
  rm,
  stat,
  unlink,
  writeFile,
} from "node:fs/promises";
import path from "node:path";

import { redactAuditEvent } from "../redaction.js";
import {
  SEGMENT_STATES,
  createSegmentBaseName,
  cursorFileName,
  lockFileName,
  parseSegmentFileName,
  segmentFileName,
} from "./segment.js";

const UUID_REGEX = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;

function assertUuid(id) {
  if (typeof id !== "string" || !UUID_REGEX.test(id)) {
    throw new Error("audit event id must be a UUID");
  }
}

function redactIfConfigured(event, options) {
  if (options?.redact === false) return event;
  return redactAuditEvent(event, options?.redactionOptions);
}

async function safeUnlink(filePath) {
  try {
    await unlink(filePath);
  } catch (error) {
    if (error?.code === "ENOENT") return;
    throw error;
  }
}

async function syncDir(dirPath) {
  try {
    const handle = await openFile(dirPath, "r");
    try {
      await handle.sync();
    } finally {
      await handle.close();
    }
  } catch (error) {
    if (error?.code === "ENOENT") return;
    // Some platforms/filesystems do not support fsync on directories.
    if (error?.code === "EINVAL" || error?.code === "EPERM" || error?.code === "EACCES") return;
    throw error;
  }
}

async function touchFile(filePath) {
  try {
    // Overwrite contents to bump mtime. This is intentionally lightweight; the lock file is advisory.
    await writeFile(filePath, JSON.stringify({ pid: process.pid, updatedAt: Date.now() }), "utf8");
  } catch (error) {
    if (error?.code === "ENOENT") return;
    throw error;
  }
}

async function safeReadJson(filePath) {
  try {
    const raw = await readFile(filePath, "utf8");
    return JSON.parse(raw);
  } catch (error) {
    if (error?.code === "ENOENT") return null;
    return null;
  }
}

function isPidAlive(pid) {
  if (!Number.isInteger(pid) || pid <= 0) return false;
  try {
    process.kill(pid, 0);
    return true;
  } catch (error) {
    if (error?.code === "ESRCH") return false;
    if (error?.code === "EPERM") return true;
    return false;
  }
}

async function atomicWriteJson(filePath, value) {
  const tmpPath = `${filePath}.tmp`;
  const payload = JSON.stringify(value);

  const handle = await openFile(tmpPath, "w");
  try {
    await handle.writeFile(payload, "utf8");
    await handle.sync();
  } finally {
    await handle.close();
  }

  await rename(tmpPath, filePath);
  await syncDir(path.dirname(filePath));
}

async function readCursor(filePath) {
  try {
    const raw = await readFile(filePath, "utf8");
    const parsed = JSON.parse(raw);
    const acked = Number(parsed?.acked);
    return Number.isFinite(acked) && acked >= 0 ? acked : 0;
  } catch (error) {
    if (error?.code === "ENOENT") return 0;
    return 0;
  }
}

async function readJsonlEvents(filePath) {
  let raw;
  try {
    raw = await readFile(filePath, "utf8");
  } catch (error) {
    if (error?.code === "ENOENT") return { events: [], lineCount: 0 };
    throw error;
  }

  if (!raw.trim()) return { events: [], lineCount: 0 };

  const lines = raw.split("\n");
  const events = [];
  let lineCount = 0;

  for (const line of lines) {
    if (!line) continue;
    try {
      events.push(JSON.parse(line));
      lineCount += 1;
    } catch {
      // Treat parse errors as a truncated tail record from a crash during write.
      break;
    }
  }

  return { events, lineCount };
}

async function calculateDirBytes(dirPath) {
  let entries;
  try {
    entries = await readdir(dirPath, { withFileTypes: true });
  } catch (error) {
    if (error?.code === "ENOENT") return 0;
    throw error;
  }

  let total = 0;
  for (const entry of entries) {
    if (!entry.isFile()) continue;
    const fullPath = path.join(dirPath, entry.name);
    try {
      const info = await stat(fullPath);
      total += info.size;
    } catch (error) {
      if (error?.code === "ENOENT") continue;
      throw error;
    }
  }
  return total;
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function acquireLock(
  lockPath,
  { staleMs = 5 * 60_000, timeoutMs = 30_000, busyMessage = "offline audit queue is currently locked" } = {}
) {
  const startedAt = Date.now();
  let delayMs = 25;

  while (true) {
    try {
      const handle = await openFile(lockPath, "wx");
      await handle.writeFile(JSON.stringify({ pid: process.pid, createdAt: Date.now() }), "utf8");
      await handle.close();
      return;
    } catch (error) {
      if (error?.code !== "EEXIST") throw error;

      try {
        const info = await stat(lockPath);
        if (Date.now() - info.mtimeMs > staleMs) {
          await safeUnlink(lockPath);
          continue;
        }
      } catch (statError) {
        if (statError?.code === "ENOENT") continue;
        throw statError;
      }

      if (Date.now() - startedAt > timeoutMs) {
        const locked = new Error(busyMessage);
        locked.code = "EQUEUELOCKED";
        throw locked;
      }

      await sleep(delayMs);
      delayMs = Math.min(1_000, Math.floor(delayMs * 1.5));
    }
  }
}

async function acquireFlushLock(lockPath, options) {
  return acquireLock(lockPath, { busyMessage: "offline audit queue is currently flushing", ...options });
}

async function acquireEnqueueLock(lockPath, options) {
  return acquireLock(lockPath, { busyMessage: "offline audit queue is currently busy", ...options });
}

export class NodeFsOfflineAuditQueue {
  constructor(options) {
    if (!options || !options.dirPath) throw new Error("NodeFsOfflineAuditQueue requires dirPath");

    this.dirPath = options.dirPath;
    this.segmentsDir = options.segmentsDir ?? path.join(this.dirPath, "segments");
    this.maxBytes = options.maxBytes ?? 50 * 1024 * 1024;
    this.maxSegmentBytes = options.maxSegmentBytes ?? 512 * 1024;
    this.maxSegmentAgeMs = options.maxSegmentAgeMs ?? 60_000;
    this.flushBatchSize = options.flushBatchSize ?? 250;
    this.redact = options.redact !== false;
    this.redactionOptions = options.redactionOptions;
    this.syncOnWrite = options.syncOnWrite !== false;

    this.currentSegment = null;
    this._mutex = Promise.resolve();
    this.flushPromise = null;

    this.lockPath = path.join(this.dirPath, "queue.flush.lock");
    this.enqueueLockPath = path.join(this.dirPath, "queue.enqueue.lock");
  }

  async ensureDir() {
    await mkdir(this.segmentsDir, { recursive: true });
  }

  _withMutex(fn) {
    const run = async () => fn();
    const result = this._mutex.then(run, run);
    this._mutex = result.then(
      () => undefined,
      () => undefined
    );
    return result;
  }

  async _createSegment() {
    await this.ensureDir();

    for (let attempt = 0; attempt < 5; attempt += 1) {
      const baseName = createSegmentBaseName();
      const openName = segmentFileName(baseName, SEGMENT_STATES.OPEN);
      const openPath = path.join(this.segmentsDir, openName);
      const lockPath = path.join(this.segmentsDir, lockFileName(baseName));

      try {
        await writeFile(lockPath, JSON.stringify({ pid: process.pid, createdAt: Date.now() }), { encoding: "utf8", flag: "wx" });
      } catch (error) {
        if (error?.code === "EEXIST") continue;
        throw error;
      }

      return {
        baseName,
        createdAtMs: Date.now(),
        bytes: 0,
        openPath,
        lockPath,
      };
    }

    throw new Error("failed to create offline audit segment lock");
  }

  async _sealCurrentSegment() {
    if (!this.currentSegment) return;
    if (this.currentSegment.bytes === 0) return;

    const pendingPath = path.join(this.segmentsDir, segmentFileName(this.currentSegment.baseName, SEGMENT_STATES.PENDING));
    try {
      await rename(this.currentSegment.openPath, pendingPath);
      await syncDir(this.segmentsDir);
    } catch (error) {
      if (error?.code === "ENOENT") {
        // Nothing to seal; allow segment to be recreated.
      } else {
        throw error;
      }
    } finally {
      if (this.currentSegment.lockPath) await safeUnlink(this.currentSegment.lockPath);
      this.currentSegment = null;
    }
  }

  async _getOrCreateOpenSegment() {
    if (!this.currentSegment) this.currentSegment = await this._createSegment();

    const ageMs = Date.now() - this.currentSegment.createdAtMs;
    if (this.currentSegment.bytes >= this.maxSegmentBytes || ageMs >= this.maxSegmentAgeMs) {
      await this._sealCurrentSegment();
      this.currentSegment = await this._createSegment();
    }

    return this.currentSegment;
  }

  async enqueue(event) {
    if (!event || typeof event !== "object") throw new Error("audit event must be an object");
    assertUuid(event.id);

    const safeEvent = redactIfConfigured(event, { redact: this.redact, redactionOptions: this.redactionOptions });
    const line = JSON.stringify(safeEvent) + "\n";
    const lineBytes = Buffer.byteLength(line, "utf8");

    return this._withMutex(async () => {
      await this.ensureDir();

      await acquireEnqueueLock(this.enqueueLockPath, { staleMs: 60_000, timeoutMs: 5_000 });
      try {
        const usage = await calculateDirBytes(this.segmentsDir);
        if (usage + lineBytes > this.maxBytes) {
          const error = new Error("offline audit queue is full");
          error.code = "EQUEUEFULL";
          throw error;
        }

        const segment = await this._getOrCreateOpenSegment();
        const createdNewFile = segment.bytes === 0;

        const handle = await openFile(segment.openPath, "a");
        try {
          await handle.writeFile(line, "utf8");
          if (this.syncOnWrite) await handle.sync();
        } finally {
          await handle.close();
        }
        if (createdNewFile) await syncDir(this.segmentsDir);

        segment.bytes += lineBytes;

        if (segment.bytes >= this.maxSegmentBytes) {
          await this._sealCurrentSegment();
        }
      } finally {
        await safeUnlink(this.enqueueLockPath);
      }
    });
  }

  async _listSegments() {
    await this.ensureDir();
    const entries = await readdir(this.segmentsDir, { withFileTypes: true });
    const segments = [];
    for (const entry of entries) {
      if (!entry.isFile()) continue;
      const parsed = parseSegmentFileName(entry.name);
      if (!parsed) continue;
      segments.push({
        ...parsed,
        fileName: entry.name,
        path: path.join(this.segmentsDir, entry.name),
        cursorPath: path.join(this.segmentsDir, cursorFileName(parsed.baseName)),
        lockPath: path.join(this.segmentsDir, lockFileName(parsed.baseName)),
      });
    }

    segments.sort((a, b) => a.createdAtMs - b.createdAtMs || a.fileName.localeCompare(b.fileName));
    return segments;
  }

  async readAll() {
    // Snapshot read; intended for tests and diagnostics.
    const segments = await this._listSegments();
    const events = [];

    for (const segment of segments) {
      if (segment.state === SEGMENT_STATES.ACKED) continue;

      const acked = await readCursor(segment.cursorPath);
      const { events: segmentEvents } = await readJsonlEvents(segment.path);
      events.push(...segmentEvents.slice(acked));
    }

    return events;
  }

  async clear() {
    await this._withMutex(async () => {
      this.currentSegment = null;
      try {
        await rm(this.segmentsDir, { recursive: true, force: true });
      } catch (error) {
        if (error?.code === "ENOENT") return;
        throw error;
      }

      await safeUnlink(this.lockPath);
      await safeUnlink(this.enqueueLockPath);
    });
  }

  async _gcAckedSegments() {
    const segments = await this._listSegments();
    let deleted = false;
    for (const segment of segments) {
      if (segment.state !== SEGMENT_STATES.ACKED) continue;
      await safeUnlink(segment.path);
      await safeUnlink(segment.cursorPath);
      await safeUnlink(segment.lockPath);
      deleted = true;
    }
    if (deleted) await syncDir(this.segmentsDir);
  }

  async flushToExporter(exporter) {
    if (!exporter || typeof exporter.sendBatch !== "function") {
      throw new Error("flushToExporter requires exporter.sendBatch(events)");
    }

    if (this.flushPromise) return this.flushPromise;

    this.flushPromise = (async () => {
      // Seal our open segment so the flusher only sees immutable files.
      await this._withMutex(async () => {
        await this.ensureDir();
        await this._sealCurrentSegment();
      });

      await acquireFlushLock(this.lockPath);
      try {
        await this._gcAckedSegments();

        let sent = 0;
        const segments = await this._listSegments();

        for (const segment of segments) {
          if (segment.state === SEGMENT_STATES.ACKED) continue;

          if (segment.state === SEGMENT_STATES.PENDING) {
            const inflightPath = path.join(this.segmentsDir, segmentFileName(segment.baseName, SEGMENT_STATES.INFLIGHT));
            try {
              await rename(segment.path, inflightPath);
              await syncDir(this.segmentsDir);
              segment.state = SEGMENT_STATES.INFLIGHT;
              segment.path = inflightPath;
            } catch (error) {
              if (error?.code !== "ENOENT") throw error;
              continue;
            }
          }

          if (segment.state !== SEGMENT_STATES.INFLIGHT && segment.state !== SEGMENT_STATES.OPEN) continue;

          const { events, lineCount } = await readJsonlEvents(segment.path);
          let acked = await readCursor(segment.cursorPath);
          if (acked > lineCount) acked = lineCount;

          if (acked >= lineCount) {
            if (segment.state === SEGMENT_STATES.OPEN) {
              const lock = await safeReadJson(segment.lockPath);
              const lockedPid = Number(lock?.pid);
              const orphan = Number.isFinite(lockedPid) && lockedPid > 0 ? !isPidAlive(lockedPid) : false;
              if (!orphan) continue;

              const ackedPath = path.join(this.segmentsDir, segmentFileName(segment.baseName, SEGMENT_STATES.ACKED));
              try {
                await rename(segment.path, ackedPath);
                await syncDir(this.segmentsDir);
              } catch (error) {
                if (error?.code !== "ENOENT") throw error;
              }
              await safeUnlink(segment.cursorPath);
              await safeUnlink(segment.lockPath);
              continue;
            }
            const ackedPath = path.join(this.segmentsDir, segmentFileName(segment.baseName, SEGMENT_STATES.ACKED));
            try {
              await rename(segment.path, ackedPath);
              await syncDir(this.segmentsDir);
            } catch (error) {
              if (error?.code !== "ENOENT") throw error;
            }
            await safeUnlink(segment.cursorPath);
            continue;
          }

          while (acked < events.length) {
            const batch = events.slice(acked, acked + this.flushBatchSize);
            if (batch.length === 0) break;

            await touchFile(this.lockPath);
            await exporter.sendBatch(batch);
            acked += batch.length;
            sent += batch.length;
            await atomicWriteJson(segment.cursorPath, { acked });
            await touchFile(this.lockPath);
          }

          if (acked >= lineCount) {
            if (segment.state === SEGMENT_STATES.OPEN) {
              const lock = await safeReadJson(segment.lockPath);
              const lockedPid = Number(lock?.pid);
              const orphan = Number.isFinite(lockedPid) && lockedPid > 0 ? !isPidAlive(lockedPid) : false;
              if (!orphan) continue;

              const ackedPath = path.join(this.segmentsDir, segmentFileName(segment.baseName, SEGMENT_STATES.ACKED));
              try {
                await rename(segment.path, ackedPath);
                await syncDir(this.segmentsDir);
              } catch (error) {
                if (error?.code !== "ENOENT") throw error;
              }
              await safeUnlink(segment.cursorPath);
              await safeUnlink(segment.lockPath);
              continue;
            }
            const ackedPath = path.join(this.segmentsDir, segmentFileName(segment.baseName, SEGMENT_STATES.ACKED));
            await rename(segment.path, ackedPath);
            await syncDir(this.segmentsDir);
            await safeUnlink(segment.cursorPath);
          }
        }

        await this._gcAckedSegments();
        return { sent };
      } finally {
        await safeUnlink(this.lockPath);
      }
    })();

    try {
      return await this.flushPromise;
    } finally {
      this.flushPromise = null;
    }
  }
}
