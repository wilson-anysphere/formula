import { setTimeout as delay } from 'node:timers/promises';

/**
 * Promise-based sleep that optionally supports aborting via `AbortSignal`.
 *
 * Node's built-in `timers/promises.setTimeout` rejects with a DOM-style AbortError; normalize that
 * into a plain `Error('aborted')` so callers can treat aborts consistently.
 */
export async function sleep(ms: number, signal?: AbortSignal): Promise<void> {
  try {
    if (signal) {
      await delay(ms, undefined, { signal });
    } else {
      await delay(ms);
    }
  } catch (err) {
    const anyErr = err as { name?: string; code?: string };
    if (anyErr?.name === 'AbortError' || anyErr?.code === 'ABORT_ERR') {
      throw new Error('aborted');
    }
    throw err;
  }
}
