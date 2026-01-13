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
    expect(warn).toHaveBeenCalledTimes(1);
    expect(warn.mock.calls[0]?.[0]).toMatch(/Timed out flushing collab persistence/i);
  });
});

