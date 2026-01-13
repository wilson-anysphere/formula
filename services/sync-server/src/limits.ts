export class TokenBucketRateLimiter {
  private buckets = new Map<string, { tokens: number; lastRefillMs: number }>();
  private lastSweepMs = 0;

  constructor(
    private readonly capacity: number,
    private readonly refillMs: number
  ) {}

  consume(key: string, nowMs: number = Date.now()): boolean {
    // Treat invalid/disabled configurations as "no limit" (aligns with the
    // SlidingWindowRateLimiter behavior).
    if (this.capacity <= 0 || this.refillMs <= 0) return true;

    // Opportunistically sweep stale entries so a large number of one-off keys
    // (e.g. connection attempts from many unique IPs) doesn't grow this map
    // without bound.
    const sweepIntervalMs = Math.max(this.refillMs, 30_000);
    const maxEntriesBeforeSweep = 10_000;
    if (
      this.buckets.size > 0 &&
      (this.buckets.size > maxEntriesBeforeSweep ||
        nowMs - this.lastSweepMs > sweepIntervalMs)
    ) {
      const staleAfterMs = Math.max(this.refillMs, 1);
      for (const [bucketKey, bucket] of this.buckets) {
        if (nowMs - bucket.lastRefillMs > staleAfterMs) {
          this.buckets.delete(bucketKey);
        }
      }
      this.lastSweepMs = nowMs;
    }

    const existing = this.buckets.get(key);
    if (!existing) {
      this.buckets.set(key, { tokens: this.capacity - 1, lastRefillMs: nowMs });
      return true;
    }

    const elapsed = nowMs - existing.lastRefillMs;
    if (elapsed > 0) {
      const refillTokens = (elapsed / this.refillMs) * this.capacity;
      if (refillTokens >= 1) {
        existing.tokens = Math.min(
          this.capacity,
          existing.tokens + Math.floor(refillTokens)
        );
        existing.lastRefillMs = nowMs;
      }
    }

    if (existing.tokens <= 0) return false;
    existing.tokens -= 1;
    return true;
  }
}

/**
 * Exact sliding-window rate limiter.
 *
 * For each key, tracks timestamps of accepted events within `windowMs` and
 * rejects once more than `maxEvents` occur in that window.
 */
export class SlidingWindowRateLimiter {
  private readonly windows = new Map<
    string,
    { timestamps: number[]; startIndex: number }
  >();
  private lastSweepMs = 0;

  constructor(
    private readonly maxEvents: number,
    private readonly windowMs: number
  ) {}

  consume(key: string, nowMs: number = Date.now()): boolean {
    if (this.maxEvents <= 0 || this.windowMs <= 0) return true;

    const cutoff = nowMs - this.windowMs;

    // Opportunistically sweep stale entries so a large number of one-off keys
    // (e.g. message floods from many unique IPs) doesn't grow this map without
    // bound.
    const sweepIntervalMs = Math.max(this.windowMs, 30_000);
    const maxEntriesBeforeSweep = 10_000;
    if (
      this.windows.size > 0 &&
      (this.windows.size > maxEntriesBeforeSweep ||
        nowMs - this.lastSweepMs > sweepIntervalMs)
    ) {
      for (const [windowKey, window] of this.windows) {
        const { timestamps } = window;
        const lastTimestamp = timestamps[timestamps.length - 1];
        // If the most recent accepted event is outside the window (or the
        // timestamps array was compacted down to empty), this key has no effect
        // on future decisions and can be evicted.
        if (lastTimestamp === undefined || lastTimestamp <= cutoff) {
          this.windows.delete(windowKey);
        }
      }
      this.lastSweepMs = nowMs;
    }

    const existing = this.windows.get(key) ?? { timestamps: [], startIndex: 0 };
    const { timestamps } = existing;

    while (existing.startIndex < timestamps.length) {
      if (timestamps[existing.startIndex]! > cutoff) break;
      existing.startIndex += 1;
    }

    const inWindowCount = timestamps.length - existing.startIndex;
    if (inWindowCount >= this.maxEvents) {
      this.windows.set(key, existing);
      return false;
    }

    timestamps.push(nowMs);

    // Compaction to avoid unbounded growth from a large number of old entries.
    if (existing.startIndex > 0 && existing.startIndex * 2 >= timestamps.length) {
      existing.timestamps = timestamps.slice(existing.startIndex);
      existing.startIndex = 0;
    }

    if (existing.timestamps.length === 0) this.windows.delete(key);
    else this.windows.set(key, existing);
    return true;
  }

  reset(key: string): void {
    this.windows.delete(key);
  }
}

export class ConnectionTracker {
  private total = 0;
  private perIp = new Map<string, number>();

  constructor(
    private readonly maxTotal: number,
    private readonly maxPerIp: number
  ) {}

  snapshot() {
    return {
      total: this.total,
      uniqueIps: this.perIp.size,
    };
  }

  tryRegister(ip: string): { ok: true } | { ok: false; reason: string } {
    const currentPerIp = this.perIp.get(ip) ?? 0;
    if (this.maxTotal > 0 && this.total >= this.maxTotal) {
      return { ok: false, reason: "max_connections_exceeded" };
    }
    if (this.maxPerIp > 0 && currentPerIp >= this.maxPerIp) {
      return { ok: false, reason: "max_connections_per_ip_exceeded" };
    }

    this.total += 1;
    this.perIp.set(ip, currentPerIp + 1);
    return { ok: true };
  }

  unregister(ip: string) {
    const current = this.perIp.get(ip);
    if (current === undefined) return;

    if (current <= 1) this.perIp.delete(ip);
    else this.perIp.set(ip, current - 1);

    if (this.total > 0) this.total -= 1;
  }
}
