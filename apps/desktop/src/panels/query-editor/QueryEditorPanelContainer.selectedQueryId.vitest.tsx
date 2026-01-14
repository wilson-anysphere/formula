/**
 * @vitest-environment jsdom
 */

import React, { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import type { Query } from "@formula/power-query";

// React 18 relies on this flag to suppress act() warnings in test runners.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

async function flushMicrotasks(count = 10): Promise<void> {
  for (let i = 0; i < count; i += 1) await Promise.resolve();
}

function createInMemoryLocalStorage(): Storage {
  const store = new Map<string, string>();
  return {
    getItem: (key: string) => (store.has(key) ? store.get(key)! : null),
    setItem: (key: string, value: string) => {
      store.set(String(key), String(value));
    },
    removeItem: (key: string) => {
      store.delete(String(key));
    },
    clear: () => {
      store.clear();
    },
    key: (index: number) => Array.from(store.keys())[index] ?? null,
    get length() {
      return store.size;
    },
  } as Storage;
}

const mocks = vi.hoisted(() => {
  const captured: { props: any } = { props: null };

  const baseQuery = (id: string): Query =>
    ({
      id,
      name: `Query ${id}`,
      source: { type: "range", range: { values: [], hasHeaders: false } },
      steps: [],
      refreshPolicy: { type: "manual" },
    }) as any;

  const queries: Query[] = [baseQuery("q1"), baseQuery("q2")];

  const service: any = {
    ready: Promise.resolve(),
    engine: {},
    engineError: null,
    getQueries: vi.fn(() => queries),
    getQuery: vi.fn((id: string) => queries.find((q) => q.id === id) ?? null),
    onEvent: vi.fn(() => () => {}),
  };

  const QueryEditorPanel = (props: any) => {
    captured.props = props;
    return null;
  };

  return { captured, service, QueryEditorPanel };
});

vi.mock("./QueryEditorPanel.js", () => ({
  QueryEditorPanel: mocks.QueryEditorPanel,
}));

vi.mock("../../power-query/service.js", () => ({
  getDesktopPowerQueryService: () => mocks.service,
  onDesktopPowerQueryServiceChanged: () => () => {},
  DesktopPowerQueryService: class {},
}));

vi.mock("../../power-query/engine.js", () => ({
  getContextForDocument: () => ({}),
}));

vi.mock("../../tauri/api", () => ({
  hasTauri: () => false,
  getTauriDialogOpenOrNull: () => null,
}));

vi.mock("../../tauri/nativeDialogs.js", () => ({}));

vi.mock("../../extensions/ui.js", () => ({
  showInputBox: vi.fn(async () => null),
}));

import { QueryEditorPanelContainer } from "./QueryEditorPanelContainer";

describe("QueryEditorPanelContainer", () => {
  let host: HTMLDivElement | null = null;
  let root: ReturnType<typeof createRoot> | null = null;
  let localStorageInstance: Storage | null = null;

  beforeEach(() => {
    mocks.captured.props = null;
    host = document.createElement("div");
    document.body.appendChild(host);
    root = createRoot(host);

    localStorageInstance = createInMemoryLocalStorage();
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: localStorageInstance });
  });

  afterEach(async () => {
    if (root) {
      await act(async () => {
        root?.unmount();
      });
    }
    host?.remove();
    host = null;
    root = null;
    localStorageInstance = null;
    vi.restoreAllMocks();
  });

  it("loads the selected query id from localStorage and trims whitespace", async () => {
    const workbookId = "workbook-1";
    const key = `formula.desktop.powerQuery.selectedQuery:${workbookId}`;
    localStorage!.setItem(key, "  q2  ");

    await act(async () => {
      root?.render(<QueryEditorPanelContainer workbookId={workbookId} getDocumentController={() => ({})} />);
      await flushMicrotasks(20);
    });

    expect(mocks.captured.props?.query?.id).toBe("q2");
    expect(localStorage!.getItem(key)).toBe("q2");
  });
});

