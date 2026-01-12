export type DebouncedFn<TArgs extends any[]> = {
  (...args: TArgs): void;
  /**
   * Immediately invoke the pending callback (if any) and clear the timer.
   */
  flush: () => void;
  /**
   * Cancel any pending callback.
   */
  cancel: () => void;
  /**
   * Returns true if a callback is currently scheduled.
   */
  pending: () => boolean;
};

/**
 * Tiny debounce helper (lodash-style) with `flush()` support.
 *
 * - Used by the command palette to avoid rescoring large lists on every keystroke.
 * - `flush()` is important so pressing Enter/Arrow keys immediately after typing
 *   runs against the latest query (prevents stale selection).
 */
export function debounce<TArgs extends any[]>(
  fn: (...args: TArgs) => void,
  waitMs: number,
): DebouncedFn<TArgs> {
  const wait = Math.max(0, Math.floor(waitMs));

  let timer: ReturnType<typeof setTimeout> | null = null;
  let lastArgs: TArgs | null = null;

  const invoke = () => {
    if (!lastArgs) return;
    const args = lastArgs;
    lastArgs = null;
    fn(...args);
  };

  const debounced = ((...args: TArgs) => {
    lastArgs = args;
    if (timer) clearTimeout(timer);
    timer = setTimeout(() => {
      timer = null;
      invoke();
    }, wait);
  }) as DebouncedFn<TArgs>;

  debounced.flush = () => {
    if (!timer) return;
    clearTimeout(timer);
    timer = null;
    invoke();
  };

  debounced.cancel = () => {
    if (timer) clearTimeout(timer);
    timer = null;
    lastArgs = null;
  };

  debounced.pending = () => timer != null;

  return debounced;
}

