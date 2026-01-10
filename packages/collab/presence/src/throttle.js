export function throttle(fn, waitMs, options = {}) {
  const now = options.now ?? (() => Date.now());
  const setTimeoutFn = options.setTimeout ?? globalThis.setTimeout;
  const clearTimeoutFn = options.clearTimeout ?? globalThis.clearTimeout;

  let lastInvokeAt = -Infinity;
  let timeoutId = null;

  const invoke = () => {
    lastInvokeAt = now();
    fn();
  };

  const throttled = () => {
    const timeUntilNextInvoke = waitMs - (now() - lastInvokeAt);

    if (timeUntilNextInvoke <= 0) {
      if (timeoutId !== null) {
        clearTimeoutFn(timeoutId);
        timeoutId = null;
      }
      invoke();
      return;
    }

    if (timeoutId !== null) return;

    timeoutId = setTimeoutFn(() => {
      timeoutId = null;
      invoke();
    }, timeUntilNextInvoke);
  };

  throttled.cancel = () => {
    if (timeoutId === null) return;
    clearTimeoutFn(timeoutId);
    timeoutId = null;
  };

  return throttled;
}

