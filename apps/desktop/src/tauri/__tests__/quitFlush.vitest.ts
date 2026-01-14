import { afterEach, describe, expect, it, vi } from "vitest";

import { flushCollabLocalPersistenceBestEffort } from "../quitFlush";

describe("flushCollabLocalPersistenceBestEffort", () => {
  afterEach(() => {
    vi.useRealTimers();
    vi.restoreAllMocks();
  });

  it("does nothing when there is no session", async () => {
    await expect(flushCollabLocalPersistenceBestEffort({ session: null })).resolves.toBeUndefined();
  });

  it("flushes local persistence when a session is provided", async () => {
    const flushLocalPersistence = vi.fn().mockResolvedValue(undefined);
    await flushCollabLocalPersistenceBestEffort({
      session: { flushLocalPersistence },
      whenIdle: vi.fn().mockResolvedValue(undefined),
      flushTimeoutMs: 100,
      idleTimeoutMs: 100,
    });
    expect(flushLocalPersistence).toHaveBeenCalledTimes(1);
    expect(flushLocalPersistence).toHaveBeenCalledWith({ compact: false });
  });

  it("awaits the idle hook before flushing", async () => {
    const calls: string[] = [];
    const whenIdle = vi.fn(async () => {
      calls.push("idle");
    });
    const flushLocalPersistence = vi.fn(async () => {
      calls.push("flush");
    });

    await flushCollabLocalPersistenceBestEffort({
      session: { flushLocalPersistence },
      whenIdle,
      flushTimeoutMs: 100,
      idleTimeoutMs: 100,
    });

    expect(calls).toEqual(["idle", "flush"]);
    expect(flushLocalPersistence).toHaveBeenCalledWith({ compact: false });
  });

  it("continues quitting even if flushLocalPersistence throws", async () => {
    const flushLocalPersistence = vi.fn(async () => {
      throw new Error("boom");
    });
    const warn = vi.fn();

    await expect(
      flushCollabLocalPersistenceBestEffort({
        session: { flushLocalPersistence },
        logger: { warn },
        flushTimeoutMs: 100,
        idleTimeoutMs: 100,
      }),
    ).resolves.toBeUndefined();

    expect(flushLocalPersistence).toHaveBeenCalledTimes(1);
    expect(flushLocalPersistence).toHaveBeenCalledWith({ compact: false });
    expect(warn).toHaveBeenCalledTimes(1);
    expect(warn.mock.calls[0]?.[0]).toMatch(/Failed to flush collab local persistence/i);
  });

  it("continues quitting even if flushLocalPersistence times out", async () => {
    vi.useFakeTimers();

    const flushLocalPersistence = vi.fn(
      () =>
        new Promise<void>(() => {
          // never resolve
        }),
    );
    const warn = vi.fn();

    const promise = flushCollabLocalPersistenceBestEffort({
      session: { flushLocalPersistence },
      logger: { warn },
      flushTimeoutMs: 250,
      idleTimeoutMs: 0,
    });

    await vi.advanceTimersByTimeAsync(250);
    await expect(promise).resolves.toBeUndefined();

    expect(flushLocalPersistence).toHaveBeenCalledTimes(1);
    expect(flushLocalPersistence).toHaveBeenCalledWith({ compact: false });
    expect(warn).toHaveBeenCalledTimes(1);
    expect(warn.mock.calls[0]?.[0]).toMatch(/Timed out flushing collab persistence/i);
  });

  it("does not emit an unhandled rejection if flushLocalPersistence rejects after the timeout", async () => {
    vi.useFakeTimers();

    const flushLocalPersistence = vi.fn(
      () =>
        new Promise<void>((_resolve, reject) => {
          setTimeout(() => reject(new Error("late failure")), 100);
        }),
    );
    const warn = vi.fn();

    const unhandled: unknown[] = [];
    const onUnhandled = (reason: unknown) => {
      unhandled.push(reason);
    };
    process.on("unhandledRejection", onUnhandled);

    try {
      const promise = flushCollabLocalPersistenceBestEffort({
        session: { flushLocalPersistence },
        logger: { warn },
        flushTimeoutMs: 50,
        idleTimeoutMs: 0,
      });

      await vi.advanceTimersByTimeAsync(50);
      await expect(promise).resolves.toBeUndefined();

      // Let the underlying promise reject after we've already timed out.
      await vi.advanceTimersByTimeAsync(100);
      // Allow Node a turn to emit any unhandled rejection events.
      await Promise.resolve();

      expect(unhandled).toHaveLength(0);
      expect(flushLocalPersistence).toHaveBeenCalledWith({ compact: false });
    } finally {
      process.off("unhandledRejection", onUnhandled);
    }
  });
});
