// @vitest-environment jsdom
import React, { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it, vi } from "vitest";

import { SheetTabStrip } from "../SheetTabStrip";
import { WorkbookSheetStore } from "../workbookSheetStore";

const originalScrollIntoView = (HTMLElement.prototype as any).scrollIntoView;

afterEach(() => {
  document.body.innerHTML = "";
  // React 18 act env flag is set per-test in `renderSheetTabStrip`.
  delete (globalThis as any).IS_REACT_ACT_ENVIRONMENT;
  vi.restoreAllMocks();
  if (originalScrollIntoView) {
    (HTMLElement.prototype as any).scrollIntoView = originalScrollIntoView;
  } else {
    delete (HTMLElement.prototype as any).scrollIntoView;
  }
});

function renderSheetTabStrip(opts: { store: WorkbookSheetStore; onRenameSheet: (sheetId: string, name: string) => unknown }) {
  (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;
  // JSDOM doesn't implement scrollIntoView; SheetTabStrip uses it in an effect.
  if (typeof (HTMLElement.prototype as any).scrollIntoView !== "function") {
    (HTMLElement.prototype as any).scrollIntoView = () => {};
  }

  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);

  act(() => {
    root.render(
      React.createElement(SheetTabStrip, {
        store: opts.store,
        activeSheetId: opts.store.listVisible()[0]?.id ?? "",
        onActivateSheet: () => {},
        onAddSheet: () => {},
        onRenameSheet: opts.onRenameSheet,
      }),
    );
  });

  return { container, root };
}

describe("SheetTabStrip rename", () => {
  it("commits the latest input value even when React state updates are batched", async () => {
    const store = new WorkbookSheetStore([{ id: "s1", name: "Sheet1", visibility: "visible" }]);
    const onRenameSheet = vi.fn().mockResolvedValue(undefined);

    const { container, root } = renderSheetTabStrip({ store, onRenameSheet });

    const tab = container.querySelector<HTMLButtonElement>('[data-testid="sheet-tab-s1"]');
    expect(tab).toBeInstanceOf(HTMLButtonElement);

    act(() => {
      tab!.dispatchEvent(new MouseEvent("dblclick", { bubbles: true, cancelable: true }));
    });

    const input = tab!.querySelector<HTMLInputElement>("input.sheet-tab__input");
    expect(input).toBeInstanceOf(HTMLInputElement);

    await act(async () => {
      // Simulate React 18 batched updates: fire the "input" event and the commit
      // shortcut (Enter) in the same tick so `draftName` state hasn't flushed yet.
      (input as HTMLInputElement).value = "Budget 2024";
      input!.dispatchEvent(new Event("input", { bubbles: true }));
      input!.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true, cancelable: true }));

      await Promise.resolve();
    });

    expect(onRenameSheet).toHaveBeenCalledTimes(1);
    expect(onRenameSheet).toHaveBeenCalledWith("s1", "Budget 2024");

    act(() => root.unmount());
  });
});

