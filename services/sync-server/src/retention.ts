import type { Logger } from "pino";

export const LAST_SEEN_META_KEY = "lastSeenMs";
export const LAST_FLUSHED_META_KEY = "lastFlushedMs";

export const DEFAULT_LAST_SEEN_THROTTLE_MS = 60_000;

export type LeveldbPersistenceLike = {
  getAllDocNames: () => Promise<string[]>;
  clearDocument: (docName: string) => Promise<void>;
  getMeta: (docName: string, metaKey: string) => Promise<unknown>;
  setMeta: (docName: string, metaKey: string, value: unknown) => Promise<void>;
};

export type RetentionSweepError = { docName: string; message: string };

export type RetentionSweepResult = {
  scanned: number;
  purged: number;
  skippedActive: number;
  skippedNoMeta: number;
  errors: RetentionSweepError[];
};

export class DocConnectionTracker {
  private readonly counts = new Map<string, number>();

  register(docName: string) {
    this.counts.set(docName, (this.counts.get(docName) ?? 0) + 1);
  }

  unregister(docName: string) {
    const current = this.counts.get(docName);
    if (current === undefined) return;
    if (current <= 1) this.counts.delete(docName);
    else this.counts.set(docName, current - 1);
  }

  isActive(docName: string): boolean {
    return (this.counts.get(docName) ?? 0) > 0;
  }
}

function toErrorMessage(err: unknown): string {
  if (err instanceof Error) return err.message;
  return typeof err === "string" ? err : JSON.stringify(err);
}

function coerceFiniteNumber(value: unknown): number | null {
  if (typeof value !== "number") return null;
  return Number.isFinite(value) ? value : null;
}

export class LeveldbRetentionManager {
  private readonly lastSeenWriteMs = new Map<string, number>();
  private readonly purgingDocs = new Set<string>();
  private lastSeenWriteSweepAtMs = 0;

  constructor(
    private readonly ldb: LeveldbPersistenceLike,
    private readonly docConnections: DocConnectionTracker,
    private readonly logger: Logger,
    private readonly ttlMs: number,
    private readonly throttleMs: number = DEFAULT_LAST_SEEN_THROTTLE_MS
  ) {}

  isPurging(docName: string): boolean {
    return this.purgingDocs.has(docName);
  }

  async markSeen(
    docName: string,
    { nowMs = Date.now(), force = false }: { nowMs?: number; force?: boolean } = {}
  ): Promise<void> {
    if (this.purgingDocs.has(docName)) return;

    // Opportunistically clean up the throttling map so large numbers of one-off
    // documents don't grow memory usage without bound.
    const sweepIntervalMs = Math.max(this.throttleMs, 30_000);
    const maxEntriesBeforeSweep = 10_000;
    if (
      this.lastSeenWriteMs.size > 0 &&
      (this.lastSeenWriteMs.size > maxEntriesBeforeSweep ||
        nowMs - this.lastSeenWriteSweepAtMs > sweepIntervalMs)
    ) {
      const staleAfterMs = Math.max(this.throttleMs * 2, sweepIntervalMs);
      for (const [key, lastWrite] of this.lastSeenWriteMs) {
        if (nowMs - lastWrite > staleAfterMs) {
          this.lastSeenWriteMs.delete(key);
        }
      }
      this.lastSeenWriteSweepAtMs = nowMs;
    }

    const lastWrite = this.lastSeenWriteMs.get(docName) ?? 0;
    if (!force && nowMs - lastWrite < this.throttleMs) return;
    this.lastSeenWriteMs.set(docName, nowMs);

    try {
      await this.ldb.setMeta(docName, LAST_SEEN_META_KEY, nowMs);
    } catch (err) {
      this.logger.warn({ err, docName }, "retention_set_last_seen_failed");
    }
  }

  async markFlushed(
    docName: string,
    { nowMs = Date.now() }: { nowMs?: number } = {}
  ): Promise<void> {
    if (this.purgingDocs.has(docName)) return;

    try {
      await this.ldb.setMeta(docName, LAST_FLUSHED_META_KEY, nowMs);
    } catch (err) {
      this.logger.warn({ err, docName }, "retention_set_last_flushed_failed");
    }
  }

  async sweep({
    nowMs = Date.now(),
    maxErrors = 20,
  }: { nowMs?: number; maxErrors?: number } = {}): Promise<RetentionSweepResult> {
    if (this.ttlMs <= 0) {
      throw new Error("Retention is disabled (SYNC_SERVER_RETENTION_TTL_MS is unset or 0).");
    }

    const cutoffMs = nowMs - this.ttlMs;
    const docNames = await this.ldb.getAllDocNames();

    let scanned = 0;
    let purged = 0;
    let skippedActive = 0;
    let skippedNoMeta = 0;

    const errors: RetentionSweepError[] = [];

    for (const docName of docNames) {
      scanned += 1;

      if (this.docConnections.isActive(docName) || this.purgingDocs.has(docName)) {
        skippedActive += 1;
        continue;
      }

      let lastSeenMs: number | null;
      try {
        lastSeenMs = coerceFiniteNumber(
          await this.ldb.getMeta(docName, LAST_SEEN_META_KEY)
        );
      } catch (err) {
        if (errors.length < maxErrors) {
          errors.push({ docName, message: toErrorMessage(err) });
        }
        continue;
      }

      if (lastSeenMs === null) {
        skippedNoMeta += 1;
        continue;
      }

      if (lastSeenMs >= cutoffMs) continue;

      // Tombstone guard: prevent new `lastSeenMs` writes (and optionally new
      // websocket connections) while a purge is in-flight.
      this.purgingDocs.add(docName);

      try {
        if (this.docConnections.isActive(docName)) {
          skippedActive += 1;
          continue;
        }

        await this.ldb.clearDocument(docName);
        this.lastSeenWriteMs.delete(docName);
        purged += 1;
      } catch (err) {
        if (errors.length < maxErrors) {
          errors.push({ docName, message: toErrorMessage(err) });
        }
      } finally {
        this.purgingDocs.delete(docName);
      }
    }

    return { scanned, purged, skippedActive, skippedNoMeta, errors };
  }
}
