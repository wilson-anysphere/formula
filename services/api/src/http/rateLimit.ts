import crypto from "node:crypto";

export type RateLimitResult = { ok: true; retryAfterMs: 0 } | { ok: false; retryAfterMs: number };

export function sha256Hex(input: string): string {
  return crypto.createHash("sha256").update(input, "utf8").digest("hex");
}

export class TokenBucketRateLimiter {
  private readonly capacity: number;
  private readonly refillMs: number;
  private readonly maxEntries: number;
  private readonly state = new Map<string, { tokens: number; updatedAt: number }>();
  private lastPrunedAt = 0;

  constructor(options: { capacity: number; refillMs: number; maxEntries?: number }) {
    this.capacity = options.capacity;
    this.refillMs = options.refillMs;
    this.maxEntries = options.maxEntries ?? 10_000;
  }

  private prune(now: number): void {
    const pruneEveryMs = this.refillMs;
    const maxAgeMs = this.refillMs * 10;
    if (this.state.size <= this.maxEntries && now - this.lastPrunedAt < pruneEveryMs) return;
    this.lastPrunedAt = now;

    for (const [key, entry] of this.state.entries()) {
      if (now - entry.updatedAt > maxAgeMs) this.state.delete(key);
    }

    // Hard cap: if an attacker forces too many unique keys, drop the entire map to avoid OOM.
    if (this.state.size > this.maxEntries * 2) this.state.clear();
  }

  take(key: string): RateLimitResult {
    if (!Number.isFinite(this.capacity) || this.capacity <= 0) return { ok: true, retryAfterMs: 0 };

    const now = Date.now();
    this.prune(now);

    const existing = this.state.get(key) ?? { tokens: this.capacity, updatedAt: now };
    const elapsed = now - existing.updatedAt;
    const refill = (elapsed / this.refillMs) * this.capacity;
    const tokens = Math.min(this.capacity, existing.tokens + refill);

    if (tokens < 1) {
      this.state.set(key, { tokens, updatedAt: now });
      const refillRatePerMs = this.capacity / this.refillMs;
      const missing = Math.max(0, 1 - tokens);
      const retryAfterMs =
        Number.isFinite(refillRatePerMs) && refillRatePerMs > 0
          ? Math.ceil(missing / refillRatePerMs)
          : this.refillMs;
      return { ok: false, retryAfterMs };
    }

    this.state.set(key, { tokens: tokens - 1, updatedAt: now });
    return { ok: true, retryAfterMs: 0 };
  }
}

