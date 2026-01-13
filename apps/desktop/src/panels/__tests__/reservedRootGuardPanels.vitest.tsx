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

  constructor(ws: any) {
    this.ws = ws;
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
      await waitFor(() => container.textContent?.includes("Create checkpoint") ?? false);
    });

    const createBtnBefore = Array.from(container.querySelectorAll("button")).find((b) => b.textContent === "Create checkpoint");
    expect(createBtnBefore).toBeInstanceOf(HTMLButtonElement);
    expect((createBtnBefore as HTMLButtonElement).disabled).toBe(false);

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

    const createBtnAfter = Array.from(container.querySelectorAll("button")).find((b) => b.textContent === "Create checkpoint");
    expect(createBtnAfter).toBeInstanceOf(HTMLButtonElement);
    expect((createBtnAfter as HTMLButtonElement).disabled).toBe(true);

    await act(async () => {
      root.unmount();
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
  });
});

