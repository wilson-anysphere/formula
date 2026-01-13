import { createHash } from "node:crypto";
import { createReadStream } from "node:fs";
import { promises as fs } from "node:fs";
import path from "node:path";
import readline from "node:readline";

import type { Logger } from "pino";

export type TombstoneRecord = {
  deletedAtMs: number;
};

/**
 * Legacy tombstone format (schemaVersion=1) stored all tombstones in a single
 * JSON file (`tombstones.json`). Any mutation rewrote the full JSON map, which
 * becomes increasingly expensive on high-churn deployments.
 *
 * New format (schemaVersion=2) is append-friendly:
 *  - `tombstones.snapshot.json` stores a compact snapshot of the current map.
 *    It is only rewritten during compaction.
 *  - `tombstones.log` is an append-only JSONL stream of mutations ("set"/"delete").
 *
 * Tradeoffs:
 *  - Pros: O(1) appends for most writes, no large JSON rewrites on every purge.
 *  - Cons: log grows over time and needs periodic compaction; recovery replays
 *    the log after loading the snapshot. (A sync-server dataDir lock already
 *    prevents multi-process writers in production; without that lock this file
 *    format is not safe for concurrent writers.)
 */

type TombstonesFileV1 = {
  schemaVersion: 1;
  tombstones: Record<string, TombstoneRecord>;
};

type TombstonesSnapshotV2 = {
  schemaVersion: 2;
  tombstones: Record<string, TombstoneRecord>;
};

type TombstoneLogRecord =
  | {
      op: "set";
      docKey: string;
      deletedAtMs: number;
    }
  | {
      op: "delete";
      docKey: string;
    };

async function atomicWriteFile(filePath: string, contents: string): Promise<void> {
  const tmpPath = `${filePath}.${process.pid}.${Date.now()}.tmp`;
  await fs.writeFile(tmpPath, contents, { encoding: "utf8", mode: 0o600 });
  try {
    await fs.rename(tmpPath, filePath);
  } catch (err) {
    const code = (err as NodeJS.ErrnoException).code;
    if (code === "EEXIST" || code === "EPERM") {
      await fs.rm(filePath, { force: true });
      await fs.rename(tmpPath, filePath);
      return;
    }
    throw err;
  }
}

export function docKeyFromDocName(docName: string): string {
  return createHash("sha256").update(docName).digest("hex");
}

export class TombstoneStore {
  private readonly legacyPath: string;
  private readonly snapshotPath: string;
  private readonly logPath: string;
  private initialized = false;
  private initPromise: Promise<void> | null = null;
  private readonly tombstones = new Map<string, TombstoneRecord>();
  private writeQueue: Promise<void> = Promise.resolve();
  private compactInFlight: Promise<void> | null = null;
  private bytesSinceCompaction = 0;
  private recordsSinceCompaction = 0;

  constructor(
    private readonly dataDir: string,
    private readonly logger: Logger
  ) {
    this.legacyPath = path.join(dataDir, "tombstones.json");
    this.snapshotPath = path.join(dataDir, "tombstones.snapshot.json");
    this.logPath = path.join(dataDir, "tombstones.log");
  }

  isInitialized(): boolean {
    return this.initialized;
  }

  async init(): Promise<void> {
    if (this.initialized) return;
    if (this.initPromise) return await this.initPromise;

    this.initPromise = this.initImpl();
    try {
      await this.initPromise;
    } finally {
      this.initPromise = null;
    }
  }

  private async initImpl(): Promise<void> {
    await fs.mkdir(this.dataDir, { recursive: true });

    const hasSnapshot = await this.fileExists(this.snapshotPath);
    const hasLog = await this.fileExists(this.logPath);
    const hasNewFormat = hasSnapshot || hasLog;

    // One-time migration from legacy `tombstones.json` to the new append-only format.
    // If the new format already exists, treat it as the source of truth.
    if (!hasNewFormat && (await this.fileExists(this.legacyPath))) {
      await this.loadLegacyJson(this.legacyPath);
      await this.writeSnapshot();
      await this.backupLegacyFile();
      // Note: we intentionally do not create `tombstones.log` here; it will be
      // created on first mutation with mode 0600.
    } else {
      if (hasSnapshot) {
        await this.loadSnapshot(this.snapshotPath);
      }
      if (hasLog) {
        await this.loadLog(this.logPath);
        // Keep approximate counters so compaction can trigger after restart if the
        // log is already large.
        await this.bootstrapCompactionCountersFromDisk();
      }
    }

    this.initialized = true;
  }

  count(): number {
    return this.tombstones.size;
  }

  has(docKey: string): boolean {
    return this.tombstones.has(docKey);
  }

  entries(): Array<[string, TombstoneRecord]> {
    return [...this.tombstones.entries()];
  }

  async set(docKey: string, deletedAtMs: number = Date.now()): Promise<void> {
    this.tombstones.set(docKey, { deletedAtMs });
    await this.appendLog({ op: "set", docKey, deletedAtMs });
  }

  async delete(docKey: string): Promise<void> {
    if (!this.tombstones.delete(docKey)) return;
    await this.appendLog({ op: "delete", docKey });
  }

  async sweepExpired(ttlMs: number, nowMs: number = Date.now()): Promise<{
    expiredDocKeys: string[];
  }> {
    const expiredDocKeys: string[] = [];
    if (ttlMs <= 0) return { expiredDocKeys };

    const cutoff = nowMs - ttlMs;
    for (const [docKey, record] of this.tombstones.entries()) {
      if (record.deletedAtMs <= cutoff) {
        this.tombstones.delete(docKey);
        expiredDocKeys.push(docKey);
      }
    }

    if (expiredDocKeys.length > 0) {
      await this.appendLog(expiredDocKeys.map((docKey) => ({ op: "delete", docKey })));
    }
    return { expiredDocKeys };
  }

  private async appendLog(record: TombstoneLogRecord | TombstoneLogRecord[]): Promise<void> {
    const records = Array.isArray(record) ? record : [record];
    if (records.length === 0) return;

    this.writeQueue = this.writeQueue
      .catch(() => {
        // keep queue alive
      })
      .then(async () => {
        await fs.mkdir(this.dataDir, { recursive: true });

        // Using a file handle ensures we can set a restrictive mode on create.
        // (appendFile() only applies `mode` when the file is created.)
        const fh = await fs.open(this.logPath, "a", 0o600);
        try {
          const payload = `${records.map((r) => JSON.stringify(r)).join("\n")}\n`;
          // For string writes, FileHandle.write accepts an optional numeric position.
          // In append mode ("a") the position is ignored, so omit it.
          await fh.write(payload, undefined, "utf8");
          this.bytesSinceCompaction += Buffer.byteLength(payload, "utf8");
          this.recordsSinceCompaction += records.length;
        } finally {
          await fh.close();
        }

        await this.maybeCompactAfterAppend();
      });

    return await this.writeQueue;
  }

  private async maybeCompactAfterAppend(): Promise<void> {
    // Avoid stat() on every write: only consider compaction periodically.
    const MAX_LOG_BYTES_BEFORE_COMPACTION = 16 * 1024 * 1024; // 16 MiB
    const MIN_RECORDS_BEFORE_COMPACTION = 10_000;

    if (
      this.bytesSinceCompaction < MAX_LOG_BYTES_BEFORE_COMPACTION &&
      this.recordsSinceCompaction < MIN_RECORDS_BEFORE_COMPACTION
    ) {
      return;
    }

    // If the log already exists and is larger than our threshold, compact.
    let stSize = 0;
    try {
      const st = await fs.stat(this.logPath);
      stSize = st.size;
    } catch (err) {
      const code = (err as NodeJS.ErrnoException).code;
      if (code !== "ENOENT") throw err;
      return;
    }

    if (stSize < MAX_LOG_BYTES_BEFORE_COMPACTION) return;

    if (this.compactInFlight) {
      await this.compactInFlight;
      return;
    }

    // Run compaction in the existing writeQueue to preserve ordering with
    // subsequent appends.
    this.compactInFlight = this.writeSnapshot()
      .then(() => this.truncateLog())
      .catch((err) => {
        // Compaction failures should not break tombstone writes (we can recover by
        // replaying the log). Log and continue.
        this.logger.warn({ err }, "tombstones_compaction_failed");
      })
      .finally(() => {
        this.compactInFlight = null;
      });

    await this.compactInFlight;
  }

  private async truncateLog(): Promise<void> {
    try {
      await fs.truncate(this.logPath, 0);
    } catch (err) {
      const code = (err as NodeJS.ErrnoException).code;
      if (code !== "ENOENT") throw err;
    }
    this.bytesSinceCompaction = 0;
    this.recordsSinceCompaction = 0;
  }

  private async writeSnapshot(): Promise<void> {
    const sorted = [...this.tombstones.entries()].sort(([a], [b]) => a.localeCompare(b));
    const json: TombstonesSnapshotV2 = {
      schemaVersion: 2,
      tombstones: Object.fromEntries(sorted),
    };
    await atomicWriteFile(this.snapshotPath, `${JSON.stringify(json)}\n`);
  }

  private async loadSnapshot(snapshotPath: string): Promise<void> {
    try {
      const raw = await fs.readFile(snapshotPath, "utf8");
      const parsed: unknown = JSON.parse(raw);
      if (
        !parsed ||
        typeof parsed !== "object" ||
        (parsed as { schemaVersion?: unknown }).schemaVersion !== 2 ||
        !(parsed as { tombstones?: unknown }).tombstones ||
        typeof (parsed as { tombstones?: unknown }).tombstones !== "object"
      ) {
        this.logger.warn({ filePath: snapshotPath }, "tombstones_invalid_snapshot");
        return;
      }
      const tombstones = (parsed as TombstonesSnapshotV2).tombstones;
      for (const [docKey, record] of Object.entries(tombstones)) {
        if (
          typeof docKey === "string" &&
          record &&
          typeof record === "object" &&
          typeof (record as TombstoneRecord).deletedAtMs === "number"
        ) {
          this.tombstones.set(docKey, { deletedAtMs: (record as TombstoneRecord).deletedAtMs });
        }
      }
    } catch (err) {
      const code = (err as NodeJS.ErrnoException).code;
      if (code !== "ENOENT") throw err;
    }
  }

  private async loadLog(logPath: string): Promise<void> {
    try {
      const rl = readline.createInterface({
        input: createReadStream(logPath, { encoding: "utf8" }),
        crlfDelay: Infinity,
      });
      for await (const line of rl) {
        const trimmed = line.trim();
        if (!trimmed) continue;
        let parsed: unknown;
        try {
          parsed = JSON.parse(trimmed);
        } catch {
          // Most commonly caused by a partial last-line write after a crash.
          this.logger.warn({ filePath: logPath }, "tombstones_invalid_log_line");
          continue;
        }

        if (!parsed || typeof parsed !== "object") continue;
        const op = (parsed as { op?: unknown }).op;
        const docKey = (parsed as { docKey?: unknown }).docKey;
        if (op !== "set" && op !== "delete") continue;
        if (typeof docKey !== "string" || docKey.length === 0) continue;

        if (op === "set") {
          const deletedAtMs = (parsed as { deletedAtMs?: unknown }).deletedAtMs;
          if (typeof deletedAtMs !== "number") continue;
          this.tombstones.set(docKey, { deletedAtMs });
        } else {
          this.tombstones.delete(docKey);
        }
      }
    } catch (err) {
      const code = (err as NodeJS.ErrnoException).code;
      if (code !== "ENOENT") throw err;
    }
  }

  private async loadLegacyJson(legacyPath: string): Promise<void> {
    try {
      const raw = await fs.readFile(legacyPath, "utf8");
      const parsed: unknown = JSON.parse(raw);
      if (
        !parsed ||
        typeof parsed !== "object" ||
        (parsed as { schemaVersion?: unknown }).schemaVersion !== 1 ||
        !(parsed as { tombstones?: unknown }).tombstones ||
        typeof (parsed as { tombstones?: unknown }).tombstones !== "object"
      ) {
        this.logger.warn({ filePath: legacyPath }, "tombstones_invalid_legacy_file");
        return;
      }

      const tombstones = (parsed as TombstonesFileV1).tombstones;
      for (const [docKey, record] of Object.entries(tombstones)) {
        if (
          typeof docKey === "string" &&
          record &&
          typeof record === "object" &&
          typeof (record as TombstoneRecord).deletedAtMs === "number"
        ) {
          this.tombstones.set(docKey, { deletedAtMs: (record as TombstoneRecord).deletedAtMs });
        }
      }
    } catch (err) {
      const code = (err as NodeJS.ErrnoException).code;
      if (code !== "ENOENT") throw err;
    }
  }

  private async backupLegacyFile(): Promise<void> {
    // Keep legacy tombstones.json as a backup for manual rollback/debugging, but
    // rename it so it won't be re-imported on subsequent startups.
    if (!(await this.fileExists(this.legacyPath))) return;
    const base = `${this.legacyPath}.bak`;
    let backupPath = base;
    if (await this.fileExists(backupPath)) {
      backupPath = `${base}.${Date.now()}`;
    }
    try {
      await fs.rename(this.legacyPath, backupPath);
      this.logger.info({ from: this.legacyPath, to: backupPath }, "tombstones_migrated");
    } catch (err) {
      const code = (err as NodeJS.ErrnoException).code;
      if (code === "ENOENT") return;
      throw err;
    }
  }

  private async fileExists(filePath: string): Promise<boolean> {
    try {
      await fs.access(filePath);
      return true;
    } catch (err) {
      const code = (err as NodeJS.ErrnoException).code;
      if (code === "ENOENT") return false;
      throw err;
    }
  }

  private async bootstrapCompactionCountersFromDisk(): Promise<void> {
    try {
      const st = await fs.stat(this.logPath);
      this.bytesSinceCompaction = st.size;
      // recordsSinceCompaction is only a heuristic; we'll start from 0. The byte
      // threshold is sufficient to trigger compaction if needed.
      this.recordsSinceCompaction = 0;
    } catch (err) {
      const code = (err as NodeJS.ErrnoException).code;
      if (code !== "ENOENT") throw err;
    }
  }
}
