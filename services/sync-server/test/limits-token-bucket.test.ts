import assert from "node:assert/strict";
import test from "node:test";

import { TokenBucketRateLimiter } from "../src/limits.js";

test("TokenBucketRateLimiter sweeps stale buckets to avoid unbounded growth", () => {
  const limiter = new TokenBucketRateLimiter(2, 1_000);

  limiter.consume("ip-a", 0);
  limiter.consume("ip-b", 0);

  assert.equal((limiter as any).buckets.size, 2);

  // Advance time past the sweep interval and trigger a new consume. Old buckets
  // should have refilled to full capacity, so evicting them is safe.
  limiter.consume("ip-c", 31_000);

  assert.equal((limiter as any).buckets.size, 1);
});

test("TokenBucketRateLimiter disables limiting when capacity is <= 0", () => {
  const limiter = new TokenBucketRateLimiter(0, 1_000);
  assert.equal(limiter.consume("ip", 0), true);
  assert.equal((limiter as any).buckets.size, 0);
});

test("TokenBucketRateLimiter disables limiting when refill window is <= 0", () => {
  const limiter = new TokenBucketRateLimiter(10, 0);
  assert.equal(limiter.consume("ip", 0), true);
  assert.equal((limiter as any).buckets.size, 0);
});
