// @vitest-environment jsdom
import React, { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it } from "vitest";
import * as Y from "yjs";

import { CollabVersionHistoryPanel } from "../version-history/CollabVersionHistoryPanel.js";
import { CollabBranchManagerPanel } from "../branch-manager/CollabBranchManagerPanel.js";

function flushPromises() {
  return new Promise<void>((resolve) => setTimeout(resolve, 0));
}

async function waitFor(condition: () => boolean, timeoutMs = 1_000) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    if (condition()) return;
    // eslint-disable-next-line no-await-in-loop
    await flushPromises();
  }
  throw new Error("Timed out waiting for condition");
}

class FakeBrowserWebSocket {
  private readonly listeners = new Map<string, Set<(ev: any) => void>>();

  addEventListener(type: string, cb: (ev: any) => void) {
    let set = this.listeners.get(type);
    if (!set) {
      set = new Set();
      this.listeners.set(type, set);
    }
    set.add(cb);
  }

  removeEventListener(type: string, cb: (ev: any) => void) {
    this.listeners.get(type)?.delete(cb);
  }

  emitClose(code: number, reason: string) {
    const ev = { code, reason };
    for (const cb of Array.from(this.listeners.get("close") ?? [])) {
      cb(ev);
    }
  }
}

class FakeProvider {
  ws: any;
  private readonly listeners = new Map<string, Set<(...args: any[]) => void>>();
  connectCalls = 0;

  constructor(ws: any) {
    this.ws = ws;
  }

  connect() {
    this.connectCalls += 1;
  }

  on(event: string, cb: (...args: any[]) => void) {
    let set = this.listeners.get(event);
    if (!set) {
      set = new Set();
      this.listeners.set(event, set);
    }
    set.add(cb);
  }

  off(event: string, cb: (...args: any[]) => void) {
    this.listeners.get(event)?.delete(cb);
  }

  emit(event: string, ...args: any[]) {
    for (const cb of Array.from(this.listeners.get(event) ?? [])) {
      cb(...args);
    }
  }
}

class FakeNodeWs {
  private readonly listeners = new Map<string, Set<(...args: any[]) => void>>();

  on(event: string, cb: (...args: any[]) => void) {
    let set = this.listeners.get(event);
    if (!set) {
      set = new Set();
      this.listeners.set(event, set);
    }
    set.add(cb);
  }

  off(event: string, cb: (...args: any[]) => void) {
    this.listeners.get(event)?.delete(cb);
  }

  emitClose(code: number, reason: any) {
    for (const cb of Array.from(this.listeners.get("close") ?? [])) {
      cb(code, reason);
    }
  }
}

afterEach(() => {
  document.body.innerHTML = "";
});

describe("sync-server reserved root guard disconnect UX", () => {
  it("shows a persistent error banner and disables version history mutations", async () => {
    (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

    const ws = new FakeBrowserWebSocket();
    const provider = new FakeProvider(ws);
    const session = { doc: new Y.Doc({ guid: "doc-1" }), provider, presence: null } as any;

    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);

    await act(async () => {
      root.render(<CollabVersionHistoryPanel session={session} />);
    });

    await act(async () => {
      await waitFor(() => container.querySelector(".collab-version-history__input") instanceof HTMLInputElement);
    });

    const nameInputBefore = container.querySelector(".collab-version-history__input") as HTMLInputElement | null;
    expect(nameInputBefore).toBeInstanceOf(HTMLInputElement);
    expect(nameInputBefore?.disabled).toBe(false);

    await act(async () => {
      ws.emitClose(1008, "reserved root mutation");
      await flushPromises();
    });

    await act(async () => {
      await waitFor(() => container.textContent?.includes("SYNC_SERVER_RESERVED_ROOT_GUARD_ENABLED") ?? false);
    });

    expect(container.textContent).toContain("reserved root guard");
    expect(container.textContent).toContain("ApiVersionStore");
    expect(container.textContent).toContain("SQLite");

    const nameInputAfter = container.querySelector(".collab-version-history__input") as HTMLInputElement | null;
    expect(nameInputAfter).toBeInstanceOf(HTMLInputElement);
    expect(nameInputAfter?.disabled).toBe(true);

    await act(async () => {
      root.unmount();
    });

    // Re-mounting the panel with the same provider should still surface the banner
    // (even though we are not re-emitting a close event).
    const container2 = document.createElement("div");
    document.body.appendChild(container2);
    const root2 = createRoot(container2);
    await act(async () => {
      root2.render(<CollabVersionHistoryPanel session={session} />);
    });
    await act(async () => {
      await waitFor(() => container2.textContent?.includes("SYNC_SERVER_RESERVED_ROOT_GUARD_ENABLED") ?? false);
    });
    await act(async () => {
      root2.unmount();
    });
  });

  it("shows a persistent error banner and disables branch manager mutations", async () => {
    (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

    const ws = new FakeBrowserWebSocket();
    const provider = new FakeProvider(ws);
    const session = { doc: new Y.Doc({ guid: "doc-2" }), provider, presence: null } as any;

    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);

    await act(async () => {
      root.render(<CollabBranchManagerPanel session={session} sheetNameResolver={null} />);
    });

    await act(async () => {
      await waitFor(() => container.querySelector("input") instanceof HTMLInputElement);
    });

    const inputBefore = container.querySelector("input") as HTMLInputElement | null;
    expect(inputBefore).toBeInstanceOf(HTMLInputElement);
    expect(inputBefore?.disabled).toBe(false);

    await act(async () => {
      ws.emitClose(1008, "reserved root mutation");
      await flushPromises();
    });

    await act(async () => {
      await waitFor(() => container.textContent?.includes("SYNC_SERVER_RESERVED_ROOT_GUARD_ENABLED") ?? false);
    });

    const inputAfter = container.querySelector("input") as HTMLInputElement | null;
    expect(inputAfter).toBeInstanceOf(HTMLInputElement);
    expect(inputAfter?.disabled).toBe(true);

    await act(async () => {
      root.unmount();
    });

    // Banner should persist across re-mounts for the same provider instance.
    const container2 = document.createElement("div");
    document.body.appendChild(container2);
    const root2 = createRoot(container2);
    await act(async () => {
      root2.render(<CollabBranchManagerPanel session={session} sheetNameResolver={null} />);
    });
    await act(async () => {
      await waitFor(() => container2.textContent?.includes("SYNC_SERVER_RESERVED_ROOT_GUARD_ENABLED") ?? false);
    });
    await act(async () => {
      root2.unmount();
    });
  });

  it("detects provider 'connection-close' events (no ws)", async () => {
    (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

    const provider = new FakeProvider(undefined);
    const session = { doc: new Y.Doc({ guid: "doc-3" }), provider, presence: null } as any;

    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);

    await act(async () => {
      root.render(<CollabVersionHistoryPanel session={session} />);
    });

    await act(async () => {
      await waitFor(() => container.querySelector(".collab-version-history__input") instanceof HTMLInputElement);
    });

    await act(async () => {
      provider.emit("connection-close", { code: 1008, reason: "reserved root mutation" });
      await flushPromises();
    });

    await act(async () => {
      await waitFor(() => container.textContent?.includes("SYNC_SERVER_RESERVED_ROOT_GUARD_ENABLED") ?? false);
    });

    await act(async () => {
      root.unmount();
    });
  });

  it("detects Node/ws close events (ws.on('close', ...))", async () => {
    (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

    const ws = new FakeNodeWs();
    const provider = new FakeProvider(ws);
    const session = { doc: new Y.Doc({ guid: "doc-4" }), provider, presence: null } as any;

    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);

    await act(async () => {
      root.render(<CollabBranchManagerPanel session={session} sheetNameResolver={null} />);
    });

    await act(async () => {
      await waitFor(() => container.querySelector("input") instanceof HTMLInputElement);
    });

    await act(async () => {
      const reason = typeof Buffer !== "undefined" ? Buffer.from("reserved root mutation") : new TextEncoder().encode("reserved root mutation");
      ws.emitClose(1008, reason);
      await flushPromises();
    });

    await act(async () => {
      await waitFor(() => container.textContent?.includes("SYNC_SERVER_RESERVED_ROOT_GUARD_ENABLED") ?? false);
    });

    await act(async () => {
      root.unmount();
    });
  });

  it("records the disconnect even if the panel is unmounted when the close event fires", async () => {
    (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

    const ws = new FakeBrowserWebSocket();
    const provider = new FakeProvider(ws);
    const session = { doc: new Y.Doc({ guid: "doc-5" }), provider, presence: null } as any;

    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);

    await act(async () => {
      root.render(<CollabVersionHistoryPanel session={session} />);
    });

    await act(async () => {
      await waitFor(() => container.querySelector(".collab-version-history__input") instanceof HTMLInputElement);
    });

    await act(async () => {
      root.unmount();
    });

    // Emit after unmount: the reserved-root-guard monitor should still cache the error.
    ws.emitClose(1008, "reserved root mutation");

    const container2 = document.createElement("div");
    document.body.appendChild(container2);
    const root2 = createRoot(container2);
    await act(async () => {
      root2.render(<CollabVersionHistoryPanel session={session} />);
    });

    await act(async () => {
      await waitFor(() => container2.textContent?.includes("SYNC_SERVER_RESERVED_ROOT_GUARD_ENABLED") ?? false);
    });

    await act(async () => {
      root2.unmount();
    });
  });

  it("tracks provider.ws replacement via provider 'status' events", async () => {
    (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

    const ws1 = new FakeBrowserWebSocket();
    const ws2 = new FakeBrowserWebSocket();
    const provider = new FakeProvider(ws1);
    const session = { doc: new Y.Doc({ guid: "doc-6" }), provider, presence: null } as any;

    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);

    await act(async () => {
      root.render(<CollabVersionHistoryPanel session={session} />);
    });

    await act(async () => {
      await waitFor(() => container.querySelector(".collab-version-history__input") instanceof HTMLInputElement);
    });

    await act(async () => {
      provider.ws = ws2;
      provider.emit("status", { status: "connected" });
      await flushPromises();
    });

    await act(async () => {
      ws2.emitClose(1008, "reserved root mutation");
      await flushPromises();
    });

    await act(async () => {
      await waitFor(() => container.textContent?.includes("SYNC_SERVER_RESERVED_ROOT_GUARD_ENABLED") ?? false);
    });

    await act(async () => {
      root.unmount();
    });
  });

  it("allows clearing the lockout via Retry", async () => {
    (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

    const ws = new FakeBrowserWebSocket();
    const provider = new FakeProvider(ws);
    const session = { doc: new Y.Doc({ guid: "doc-7" }), provider, presence: null } as any;

    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);

    await act(async () => {
      root.render(<CollabVersionHistoryPanel session={session} />);
    });

    await act(async () => {
      await waitFor(() => container.querySelector(".collab-version-history__input") instanceof HTMLInputElement);
    });

    await act(async () => {
      ws.emitClose(1008, "reserved root mutation");
      await flushPromises();
    });

    await act(async () => {
      await waitFor(() => container.textContent?.includes("SYNC_SERVER_RESERVED_ROOT_GUARD_ENABLED") ?? false);
    });

    const retryBtn = container.querySelector('[data-testid="reserved-root-guard-retry"]') as HTMLButtonElement | null;
    expect(retryBtn).toBeInstanceOf(HTMLButtonElement);

    await act(async () => {
      retryBtn?.click();
      await flushPromises();
    });

    expect(provider.connectCalls).toBe(1);

    await act(async () => {
      await waitFor(() => (container.textContent?.includes("SYNC_SERVER_RESERVED_ROOT_GUARD_ENABLED") ?? false) === false);
    });

    // The panel should eventually return to the interactive UI state.
    await act(async () => {
      await waitFor(
        () => {
          const input = container.querySelector(".collab-version-history__input") as HTMLInputElement | null;
          return Boolean(input && !input.disabled);
        },
        5_000,
      );
    });

    await act(async () => {
      root.unmount();
    });
  });
});
