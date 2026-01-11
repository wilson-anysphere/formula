function now() {
  // `performance.now()` is monotonic and available in modern browsers + Node.
  // Fall back to `Date.now()` for older environments.
  return globalThis.performance?.now ? globalThis.performance.now() : Date.now();
}

export function createAbortError(message = "The operation was aborted") {
  // DOMException is the standard AbortSignal error type in browsers.
  // Node also exposes DOMException.
  if (typeof globalThis.DOMException === "function") {
    return new DOMException(message, "AbortError");
  }
  const err = new Error(message);
  // Match DOMException shape so callers can do `err.name === "AbortError"`.
  err.name = "AbortError";
  return err;
}

export function throwIfAborted(signal) {
  if (signal?.aborted) throw createAbortError();
}

export function defaultYieldScheduler() {
  return new Promise((resolve) => {
    // Prefer setImmediate when available (Node) – it yields without clamping.
    if (typeof setImmediate === "function") {
      setImmediate(resolve);
      return;
    }
    setTimeout(resolve, 0);
  });
}

/**
 * Cooperative scheduler for long-running loops.
 *
 * Call `await slicer.checkpoint()` regularly inside tight loops to:
 * - throw promptly when `signal` is aborted;
 * - yield to the host event loop when we exceed `timeBudgetMs`.
 *
 * Yielding is based on *elapsed time* rather than fixed iteration counts, so it
 * adapts to varying per-iteration costs.
 */
export function createTimeSlicer({
  signal,
  timeBudgetMs = 10,
  // Checking the clock every iteration is surprisingly expensive. We still
  // yield based on elapsed time, but only sample the clock periodically.
  checkEvery = 256,
  scheduler = defaultYieldScheduler,
} = {}) {
  let lastYield = now();
  let ticks = 0;

  return {
    async checkpoint() {
      throwIfAborted(signal);
      if (timeBudgetMs <= 0) return;

      ticks++;
      if (checkEvery > 1 && ticks % checkEvery !== 0) return;

      const elapsed = now() - lastYield;
      if (elapsed < timeBudgetMs) return;

      await scheduler();
      throwIfAborted(signal);
      lastYield = now();
    },
  };
}

