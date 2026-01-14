import type { CollabSession } from "@formula/collab-session";

type FlushableCollabSession = Pick<CollabSession, "flushLocalPersistence">;

type Logger = {
  warn: (message: string) => void;
};

function nextMicrotask(): Promise<void> {
  return new Promise((resolve) => {
    if (typeof queueMicrotask === "function") {
      queueMicrotask(resolve);
      return;
    }
    // Extremely old JS runtimes may not provide `queueMicrotask`; fall back to a
    // resolved promise turn.
    void Promise.resolve()
      .then(resolve)
      .catch(() => {
        // Best-effort: avoid unhandled rejections if the microtask callback throws.
      });
  });
}

class QuitFlushTimeoutError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "QuitFlushTimeoutError";
  }
}

async function withTimeout<T>(promise: Promise<T>, timeoutMs: number, message: string): Promise<T> {
  if (!Number.isFinite(timeoutMs) || timeoutMs <= 0) return promise;

  let timeoutId: ReturnType<typeof setTimeout> | null = null;
  let didTimeout = false;
  const timeoutPromise = new Promise<never>((_resolve, reject) => {
    timeoutId = setTimeout(() => {
      didTimeout = true;
      reject(new QuitFlushTimeoutError(message));
    }, timeoutMs);
    // Avoid keeping Node test processes alive due to a quit-timeout timer.
    (timeoutId as any).unref?.();
  });

  try {
    return await Promise.race([promise, timeoutPromise]);
  } finally {
    if (timeoutId != null) clearTimeout(timeoutId);
    if (didTimeout) {
      // If the underlying operation continues in the background and later rejects,
      // ensure it doesn't surface as an unhandled rejection.
      promise.catch(() => {});
    }
  }
}

/**
 * Best-effort flush of collaborative local persistence (IndexedDB/File) before
 * the desktop shell hard-exits.
 *
 * This intentionally never throws: quitting should proceed even if persistence
 * is unavailable, slow, or fails.
 */
export async function flushCollabLocalPersistenceBestEffort(options: {
  session: FlushableCollabSession | null | undefined;
  /**
   * Optional "wait for app idle" hook. In the desktop renderer this is typically
   * `() => app.whenIdle()`, giving binder propagation + microtask-batched UI work
   * a chance to run before we snapshot the Yjs document.
   */
  whenIdle?: (() => Promise<void>) | null;
  /**
   * Max time to wait for the flush operation before continuing the quit anyway.
   */
  flushTimeoutMs?: number;
  /**
   * Max time to wait for `whenIdle()` before continuing the flush anyway.
   */
  idleTimeoutMs?: number;
  logger?: Logger;
}): Promise<void> {
  const session = options.session ?? null;
  if (!session || typeof (session as any).flushLocalPersistence !== "function") return;

  const flushTimeoutMs =
    typeof options.flushTimeoutMs === "number" && Number.isFinite(options.flushTimeoutMs) && options.flushTimeoutMs > 0
      ? options.flushTimeoutMs
      : 2_000;

  const idleTimeoutMs =
    typeof options.idleTimeoutMs === "number" && Number.isFinite(options.idleTimeoutMs) && options.idleTimeoutMs > 0
      ? options.idleTimeoutMs
      : 500;

  const logger: Logger = options.logger ?? console;

  try {
    // Allow the DocumentController->Yjs binder (and any other microtask-batched UI
    // propagation) a chance to apply edits before we snapshot the doc into local
    // persistence.
    await nextMicrotask();
    const whenIdle = options.whenIdle;
    if (whenIdle) {
      await withTimeout(
        Promise.resolve()
          .then(() => whenIdle())
          .catch(() => {
            // Best-effort; ignore idle failures and still attempt flush.
          }),
        idleTimeoutMs,
        "Timed out waiting for app idle",
      ).catch(() => {
        // Best-effort; proceed to flush even if idle never settles.
      });
    }
    await nextMicrotask();
  } catch {
    // Best-effort; proceed to flush.
  }

  try {
    await withTimeout(
      // Prefer a lightweight snapshot flush on quit. IndexedDB persistence defaults
      // to compaction, which can be slower and increases the chance that we time out
      // before the process hard-exits.
      Promise.resolve().then(() => session.flushLocalPersistence({ compact: false })),
      flushTimeoutMs,
      "Timed out flushing collab persistence",
    );
  } catch (err) {
    try {
      // Avoid printing any potentially sensitive state (tokens, cell contents). We log only
      // the error message and continue quitting.
      const message = err instanceof Error ? err.message : String(err);
      logger.warn(`[formula][desktop] Failed to flush collab local persistence before quit: ${message}`);
    } catch {
      // ignore
    }
  }
}
