// @vitest-environment jsdom
import { describe, expect, it } from "vitest";

import { builtinKeybindings } from "../../commands/builtinKeybindings.js";
import { CommandRegistry } from "../../extensions/commandRegistry.js";
import { ContextKeyService } from "../../extensions/contextKeys.js";
import { KeybindingService } from "../../extensions/keybindingService.js";
import { registerFindReplaceShortcuts } from "./shortcuts.js";

async function flushMicrotasks(times = 6): Promise<void> {
  for (let i = 0; i < times; i++) {
    await new Promise<void>((resolve) => queueMicrotask(resolve));
  }
}

describe("find/replace shortcuts", () => {
  it("binds Replace to Cmd+Option+F on macOS (and avoids Cmd+H)", async () => {
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

    const commandRegistry = new CommandRegistry();
    const contextKeys = new ContextKeyService();
    contextKeys.set("workbench.commandPaletteOpen", false);
    const keybindingService = new KeybindingService({ commandRegistry, contextKeys, platform: "mac" });
    keybindingService.setBuiltinKeybindings(builtinKeybindings);
    const disposeKeybindings = keybindingService.installWindowListener(window);

    function showDialogAndFocus(dialog: HTMLDialogElement): void {
      if (!dialog.open) {
        dialog.showModal();
      }

      const focusInput = () => {
        const input = dialog.querySelector<HTMLInputElement | HTMLTextAreaElement>("input, textarea");
        if (!input) return;
        input.focus();
        input.select?.();
      };

      requestAnimationFrame(focusInput);
    }

    commandRegistry.registerBuiltinCommand("edit.find", "Find", () => showDialogAndFocus(findDialog as any));
    commandRegistry.registerBuiltinCommand("edit.replace", "Replace", () => showDialogAndFocus(replaceDialog as any));

    expect(replaceDialog.hasAttribute("open")).toBe(false);
    expect(findDialog.hasAttribute("open")).toBe(false);

    // Cmd+F should open Find (and not Replace).
    window.dispatchEvent(new KeyboardEvent("keydown", { key: "f", code: "KeyF", metaKey: true, bubbles: true }));
    await flushMicrotasks();
    expect(findDialog.hasAttribute("open")).toBe(true);
    expect(replaceDialog.hasAttribute("open")).toBe(false);
    (findDialog as any).close();

    // Ctrl+F should also open Find on macOS as a fallback (Windows/Linux-style shortcut).
    window.dispatchEvent(new KeyboardEvent("keydown", { key: "f", ctrlKey: true, bubbles: true }));
    await flushMicrotasks();
    expect(findDialog.hasAttribute("open")).toBe(true);
    expect(replaceDialog.hasAttribute("open")).toBe(false);
    (findDialog as any).close();

    // Cmd+H (macOS Hide) should no longer open Replace.
    window.dispatchEvent(new KeyboardEvent("keydown", { key: "h", metaKey: true, bubbles: true }));
    await flushMicrotasks();
    expect(replaceDialog.hasAttribute("open")).toBe(false);

    // Cmd+Option+F should open Replace (and should not open Find).
    window.dispatchEvent(
      new KeyboardEvent("keydown", { key: "f", code: "KeyF", metaKey: true, altKey: true, bubbles: true }),
    );
    await flushMicrotasks();
    expect(replaceDialog.hasAttribute("open")).toBe(true);
    expect(findDialog.hasAttribute("open")).toBe(false);

    // Ctrl+H continues to work for Windows/Linux-style replace.
    (replaceDialog as any).close();
    window.dispatchEvent(new KeyboardEvent("keydown", { key: "h", ctrlKey: true, bubbles: true }));
    await flushMicrotasks();
    expect(replaceDialog.hasAttribute("open")).toBe(true);

    disposeKeybindings();
    mount.remove();
  });
});
