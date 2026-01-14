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

function renderSheetTabStrip(opts: { disableMutations: boolean }) {
  (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;
  // JSDOM doesn't implement scrollIntoView; SheetTabStrip uses it in an effect.
  if (typeof (HTMLElement.prototype as any).scrollIntoView !== "function") {
    (HTMLElement.prototype as any).scrollIntoView = () => {};
  }

  const store = new WorkbookSheetStore([{ id: "s1", name: "Sheet1", visibility: "visible" }]);
  const onAddSheet = vi.fn();
  const onRenameSheet = vi.fn();
  const onError = vi.fn();

  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);

  act(() => {
    root.render(
      React.createElement(SheetTabStrip, {
        store,
        activeSheetId: "s1",
        disableMutations: opts.disableMutations,
        onActivateSheet: () => {},
        onAddSheet,
        onRenameSheet,
        onError,
      }),
    );
  });

  return { container, root, onAddSheet, onRenameSheet, onError };
}

describe("SheetTabStrip disableMutations", () => {
  it("disables sheet mutations and surfaces an edit-mode message when renaming", () => {
    const { container, root, onAddSheet, onRenameSheet, onError } = renderSheetTabStrip({ disableMutations: true });

    const add = container.querySelector<HTMLButtonElement>('[data-testid="sheet-add"]');
    expect(add).toBeInstanceOf(HTMLButtonElement);
    expect(add?.disabled).toBe(true);

    act(() => {
      add!.click();
    });
    expect(onAddSheet).not.toHaveBeenCalled();

    const tab = container.querySelector<HTMLButtonElement>('[data-testid="sheet-tab-s1"]');
    expect(tab).toBeInstanceOf(HTMLButtonElement);

    act(() => {
      tab!.dispatchEvent(new MouseEvent("dblclick", { bubbles: true, cancelable: true }));
    });

    expect(onRenameSheet).not.toHaveBeenCalled();
    expect(onError).toHaveBeenCalledWith("Finish editing to modify sheets.");
    expect(tab!.querySelector("input.sheet-tab__input")).toBeNull();

    act(() => root.unmount());
  });
});

