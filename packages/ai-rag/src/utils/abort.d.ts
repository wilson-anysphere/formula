/**
 * Abort helpers shared across ai-rag.
 *
 * We intentionally implement AbortError as a plain Error with `name = "AbortError"`
 * (instead of DOMException) so behavior is consistent across Node, browsers, and
 * test environments.
 */

export function createAbortError(message?: string): Error;

export function throwIfAborted(signal: AbortSignal | undefined): void;

/**
 * Await a promise but reject early if the AbortSignal is triggered.
 *
 * This cannot cancel underlying work, but it ensures callers can stop waiting promptly.
 */
export function awaitWithAbort<T>(promise: Promise<T> | T, signal: AbortSignal | undefined): Promise<T>;

