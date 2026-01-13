/**
 * Create an AbortError compatible with DOM-style cancellation.
 *
 * Node's built-in abort errors are typically `DOMException` instances, but this
 * package intentionally uses a plain `Error` with `name="AbortError"` so callers
 * can depend on a stable shape across runtimes.
 *
 * @param {string} [message]
 * @returns {Error & { name: "AbortError" }}
 */
export function createAbortError(message = "Aborted") {
  const err = new Error(message);
  err.name = "AbortError";
  return /** @type {Error & { name: "AbortError" }} */ (err);
}

/**
 * Throw synchronously when the signal is already aborted.
 *
 * @param {AbortSignal | undefined} [signal]
 */
export function throwIfAborted(signal) {
  if (signal?.aborted) throw createAbortError();
}

/**
 * Await a promise but reject early if the AbortSignal is triggered.
 *
 * This cannot cancel underlying work, but it ensures callers stop waiting
 * promptly when a request is canceled.
 *
 * @template T
 * @param {Promise<T> | T} promise
 * @param {AbortSignal | undefined} [signal]
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

