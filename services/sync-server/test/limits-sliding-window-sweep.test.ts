import assert from "node:assert/strict";
import test from "node:test";

import { SlidingWindowRateLimiter } from "../src/limits.js";

test("SlidingWindowRateLimiter sweeps stale keys to avoid unbounded growth", () => {
  const limiter = new SlidingWindowRateLimiter(1, 1_000);

  for (let i = 0; i < 1_000; i += 1) {
    limiter.consume(`ip-${i}`, 0);
  }

  assert.equal((limiter as any).windows.size, 1_000);

  // Advance time beyond both the limiter window and the opportunistic sweep
  // interval (which has a minimum of 30s). Triggering a consume should sweep all
  // fully stale keys.
  limiter.consume("ip-trigger", 31_000);

  assert.equal((limiter as any).windows.size, 1);
});

