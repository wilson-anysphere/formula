// @vitest-environment jsdom
import { describe, expect, it } from "vitest";

import { registerFindReplaceShortcuts } from "./shortcuts.js";

describe("find/replace shortcuts", () => {
  it("binds Replace to Cmd+Option+F on macOS (and avoids Cmd+H)", () => {
    // JSDOM doesn't provide requestAnimationFrame by default.
    // `registerFindReplaceShortcuts` uses it to focus/select the first input.
    (globalThis as any).requestAnimationFrame ??= (cb: FrameRequestCallback) => {
      cb(0);
      return 0;
    };

    const controller = {
      scope: "sheet",
      lookIn: "values",
      query: "",
      replacement: "",
      matchCase: false,
      matchEntireCell: false,
      async findNext() {
        return null;
      },
      async findAll() {
        return [];
      },
      async replaceNext() {
        return null;
      },
      async replaceAll() {
        return null;
      },
    };

    const mount = document.createElement("div");
    document.body.appendChild(mount);

    const { findDialog, replaceDialog, goToDialog } = registerFindReplaceShortcuts({
      controller: controller as any,
      workbook: {} as any,
      getCurrentSheetName: () => "Sheet1",
      setActiveCell: () => {},
      selectRange: () => {},
      mount,
    });

    // JSDOM doesn't fully implement <dialog>. Patch in a minimal `showModal` / `close`
    // so the shortcut handler can toggle the `open` state.
    for (const dialog of [findDialog, replaceDialog, goToDialog]) {
      (dialog as any).showModal ??= () => {
        (dialog as any).open = true;
        dialog.setAttribute("open", "");
      };
      (dialog as any).close ??= () => {
        (dialog as any).open = false;
        dialog.removeAttribute("open");
      };
    }

    expect(replaceDialog.hasAttribute("open")).toBe(false);

    // Cmd+H (macOS Hide) should no longer open Replace.
    window.dispatchEvent(new KeyboardEvent("keydown", { key: "h", metaKey: true, bubbles: true }));
    expect(replaceDialog.hasAttribute("open")).toBe(false);

    // Cmd+Option+F should open Replace (and should not open Find).
    window.dispatchEvent(
      new KeyboardEvent("keydown", { key: "f", code: "KeyF", metaKey: true, altKey: true, bubbles: true }),
    );
    expect(replaceDialog.hasAttribute("open")).toBe(true);
    expect(findDialog.hasAttribute("open")).toBe(false);

    // Ctrl+H continues to work for Windows/Linux-style replace.
    (replaceDialog as any).close();
    window.dispatchEvent(new KeyboardEvent("keydown", { key: "h", ctrlKey: true, bubbles: true }));
    expect(replaceDialog.hasAttribute("open")).toBe(true);
  });
});

