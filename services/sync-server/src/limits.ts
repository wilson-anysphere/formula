export class TokenBucketRateLimiter {
  private buckets = new Map<string, { tokens: number; lastRefillMs: number }>();

  constructor(
    private readonly capacity: number,
    private readonly refillMs: number
  ) {}

  consume(key: string, nowMs: number = Date.now()): boolean {
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
    if (this.total >= this.maxTotal) {
      return { ok: false, reason: "max_connections_exceeded" };
    }
    if (currentPerIp >= this.maxPerIp) {
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

