export type AbortError = Error & { name: "AbortError" };

/**
 * Create an AbortError compatible with DOM-style cancellation.
 *
 * This package uses a plain `Error` with `name="AbortError"` so callers can rely
 * on a stable shape across runtimes (Node vs browser).
 */
export function createAbortError(message?: string): AbortError;

/**
 * Throw synchronously when the signal is already aborted.
 */
export function throwIfAborted(signal?: AbortSignal): void;

/**
 * Await a promise but reject early if the AbortSignal is triggered.
 *
 * Note: this does not cancel the underlying work, it only stops awaiting it.
 */
export function awaitWithAbort<T>(promise: Promise<T> | T, signal?: AbortSignal): Promise<T>;

