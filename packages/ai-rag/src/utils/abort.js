/**
 * Abort helpers shared across ai-rag.
 *
 * We intentionally implement AbortError as a plain Error with `name = "AbortError"`
 * (instead of DOMException) so behavior is consistent across Node, browsers, and
 * test environments.
 */

/**
 * @param {string} [message]
 */
export function createAbortError(message = "Aborted") {
  const err = new Error(message);
  err.name = "AbortError";
  return err;
}

/**
 * @param {AbortSignal | undefined} signal
 */
export function throwIfAborted(signal) {
  if (signal?.aborted) throw createAbortError();
}

/**
 * Await a promise but reject early if the AbortSignal is triggered.
 *
 * This cannot cancel underlying work (e.g. an embedder call), but it ensures callers can
 * stop waiting promptly when a request is canceled.
 *
 * @template T
 * @param {Promise<T> | T} promise
 * @param {AbortSignal | undefined} signal
 * @returns {Promise<T>}
 */
export function awaitWithAbort(promise, signal) {
  if (!signal) return Promise.resolve(promise);
  if (signal.aborted) return Promise.reject(createAbortError());

  return new Promise((resolve, reject) => {
    const onAbort = () => reject(createAbortError());
    signal.addEventListener("abort", onAbort, { once: true });

    Promise.resolve(promise).then(
      (value) => {
        signal.removeEventListener("abort", onAbort);
        resolve(value);
      },
      (error) => {
        signal.removeEventListener("abort", onAbort);
        reject(error);
      }
    );
  });
}

