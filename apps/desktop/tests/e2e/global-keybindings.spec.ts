import { expect, test } from "@playwright/test";

import { gotoDesktop, openExtensionsPanel } from "./helpers";

test.describe("global keybindings", () => {
  test("routes command palette, find/replace, and extension bindings through a single handler with input focus scoping", async ({
    page,
  }) => {
    await page.addInitScript(() => {
      // Avoid permission prompt flakiness in this suite; other e2e tests cover the prompt UI.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (globalThis as any).__formulaPermissionPrompt = async () => true;
    });
    await gotoDesktop(page);

    const primary = process.platform === "darwin" ? "Meta" : "Control";

    // Pre-grant permissions needed for the built-in Sample Hello extension to activate so the
    // extension keybinding can be exercised without an interactive modal prompt.
    await page.evaluate(() => {
      const extensionId = "formula.sample-hello";
      const key = "formula.extensionHost.permissions";
      const existing = (() => {
        try {
          const raw = localStorage.getItem(key);
          return raw ? JSON.parse(raw) : {};
        } catch {
          return {};
        }
      })();

      existing[extensionId] = {
        ...(existing[extensionId] ?? {}),
        "ui.commands": true,
        "cells.read": true,
        "cells.write": true,
      };

      localStorage.setItem(key, JSON.stringify(existing));
    });

    // Ensure the grid has focus.
    await page.locator("#grid").click();

    // Ctrl/Cmd+F opens Find.
    await page.keyboard.press(`${primary}+F`);
    await expect(page.getByTestId("find-dialog")).toBeVisible();
    await page.keyboard.press("Escape");
    await expect(page.getByTestId("find-dialog")).not.toBeVisible();

    // Ctrl/Cmd+G opens Go To.
    await page.keyboard.press(`${primary}+G`);
    await expect(page.getByTestId("goto-dialog")).toBeVisible();
    await page.keyboard.press("Escape");
    await expect(page.getByTestId("goto-dialog")).not.toBeVisible();

    // Cmd+H is reserved for the system "Hide" shortcut on macOS and should never open Replace.
    // Dispatch directly to avoid browser-level interception.
    await page.evaluate(() => {
      const target = document.activeElement ?? window;
      target.dispatchEvent(
        new KeyboardEvent("keydown", {
          key: "h",
          metaKey: true,
          bubbles: true,
          cancelable: true,
        }),
      );
    });
    await page.waitForTimeout(100);
    await expect(page.getByTestId("replace-dialog")).not.toBeVisible();

    // Ctrl+H (Win/Linux) / Cmd+Option+F (macOS) opens Replace.
    // Avoid using Playwright's `keyboard.press()` here since Ctrl+H / Cmd+Option+F may be
    // intercepted by the browser shell (History, toolbar focus, etc.) in some environments.
    // Dispatching a keydown event on the active element still exercises our focus scoping logic.
    await page.evaluate((isMac) => {
      const target = document.activeElement ?? window;
      target.dispatchEvent(
        new KeyboardEvent("keydown", {
          key: isMac ? "f" : "h",
          metaKey: isMac,
          altKey: isMac,
          ctrlKey: !isMac,
          bubbles: true,
          cancelable: true,
        }),
      );
    }, process.platform === "darwin");
    await expect(page.getByTestId("replace-dialog")).toBeVisible();
    await page.keyboard.press("Escape");
    await expect(page.getByTestId("replace-dialog")).not.toBeVisible();

    // Ctrl/Cmd+Shift+P opens the command palette.
    await page.keyboard.press(`${primary}+Shift+P`);
    await expect(page.getByTestId("command-palette")).toBeVisible();
    await page.keyboard.press("Escape");
    await expect(page.getByTestId("command-palette")).not.toBeVisible();

    // Extension keybinding still works.
    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = window.__formulaApp as any;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();

      doc.setCellValue(sheetId, { row: 0, col: 0 }, 1);
      doc.setCellValue(sheetId, { row: 0, col: 1 }, 2);
      doc.setCellValue(sheetId, { row: 1, col: 0 }, 3);
      doc.setCellValue(sheetId, { row: 1, col: 1 }, 4);

      app.selectRange({
        sheetId,
        range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
      });
    });

    await openExtensionsPanel(page);
    await expect(page.getByTestId("run-command-sampleHello.sumSelection")).toBeVisible();

    await page.keyboard.press(`${primary}+Shift+Y`);
    await expect(page.getByTestId("toast-root")).toContainText("Sum: 10");

    // Clear toasts before the "while typing in input" assertions.
    await page.evaluate(() => {
      document.getElementById("toast-root")?.replaceChildren();
    });

    // Focus the formula bar editor (which focuses the hidden textarea) and ensure input focus scoping works:
    // - Find/Go To/Replace should *not* open while typing.
    // - Command palette should still open.
    // - Extension keybindings should not fire.
    await page.getByTestId("formula-highlight").click();
    await expect(page.getByTestId("formula-input")).toBeFocused();

    // Ctrl/Cmd+F should not open our Find dialog while typing in an input.
    await page.evaluate((isMac) => {
      const target = document.activeElement ?? window;
      target.dispatchEvent(
        new KeyboardEvent("keydown", {
          key: "f",
          metaKey: isMac,
          ctrlKey: !isMac,
          bubbles: true,
          cancelable: true,
        }),
      );
    }, process.platform === "darwin");
    await page.waitForTimeout(100);
    await expect(page.getByTestId("find-dialog")).not.toBeVisible();

    // Ctrl/Cmd+G should not open Go To while typing in an input.
    await page.evaluate((isMac) => {
      const target = document.activeElement ?? window;
      target.dispatchEvent(
        new KeyboardEvent("keydown", {
          key: "g",
          metaKey: isMac,
          ctrlKey: !isMac,
          bubbles: true,
          cancelable: true,
        }),
      );
    }, process.platform === "darwin");
    await page.waitForTimeout(100);
    await expect(page.getByTestId("goto-dialog")).not.toBeVisible();

    // Replace shortcut should not open Replace while typing in an input.
    await page.evaluate((isMac) => {
      const target = document.activeElement ?? window;
      target.dispatchEvent(
        new KeyboardEvent("keydown", {
          key: isMac ? "f" : "h",
          metaKey: isMac,
          altKey: isMac,
          ctrlKey: !isMac,
          bubbles: true,
          cancelable: true,
        }),
      );
    }, process.platform === "darwin");
    await page.waitForTimeout(100);
    await expect(page.getByTestId("replace-dialog")).not.toBeVisible();

    // Ctrl/Cmd+Shift+P *should* open our command palette while typing in an input.
    // Dispatch directly to avoid any environment-specific shortcut interception.
    await page.evaluate((isMac) => {
      const target = document.activeElement ?? window;
      target.dispatchEvent(
        new KeyboardEvent("keydown", {
          key: "p",
          metaKey: isMac,
          ctrlKey: !isMac,
          shiftKey: true,
          bubbles: true,
          cancelable: true,
        }),
      );
    }, process.platform === "darwin");
    await expect(page.getByTestId("command-palette")).toBeVisible();
    await page.keyboard.press("Escape");
    await expect(page.getByTestId("command-palette")).not.toBeVisible();

    // Focus should be restored to something sensible (grid or formula bar).
    await expect
      .poll(async () => {
        return page.evaluate(() => {
          const active = document.activeElement;
          if (!active) return false;

          const formulaInput = document.querySelector('[data-testid="formula-input"]');
          if (active === formulaInput) return true;

          const grid = document.getElementById("grid");
          return grid ? grid.contains(active) : false;
        });
      })
      .toBe(true);

    // Ensure we're still "typing in an input" for the remaining assertions.
    // (The formula bar textarea is rendered on top of the highlight <pre> in edit mode, so
    // re-focus the textarea directly instead of clicking the highlight.)
    const formulaInput = page.getByTestId("formula-input");
    await formulaInput.focus();
    await expect(formulaInput).toBeFocused();

    // Comments shortcuts should not fire while typing in an input.
    await page.evaluate(() => {
      const target = document.activeElement ?? window;
      target.dispatchEvent(
        new KeyboardEvent("keydown", {
          key: "F2",
          shiftKey: true,
          bubbles: true,
          cancelable: true,
        }),
      );
    });
    await page.waitForTimeout(100);
    await expect(page.getByTestId("comments-panel")).not.toBeVisible();

    await page.evaluate((isMac) => {
      const target = document.activeElement ?? window;
      target.dispatchEvent(
        new KeyboardEvent("keydown", {
          key: "M",
          code: "KeyM",
          ctrlKey: !isMac,
          metaKey: isMac,
          shiftKey: true,
          bubbles: true,
          cancelable: true,
        }),
      );
    }, process.platform === "darwin");
    await page.waitForTimeout(100);
    await expect(page.getByTestId("comments-panel")).not.toBeVisible();

    // Some remote/VM keyboard setups emit both Ctrl+Meta on a single chord. Ensure the
    // fallback `Ctrl+Cmd+Shift+M` binding also does not fire while typing in an input.
    await page.evaluate(() => {
      const target = document.activeElement ?? window;
      target.dispatchEvent(
        new KeyboardEvent("keydown", {
          key: "M",
          code: "KeyM",
          ctrlKey: true,
          metaKey: true,
          shiftKey: true,
          bubbles: true,
          cancelable: true,
        }),
      );
    });
    await page.waitForTimeout(100);
    await expect(page.getByTestId("comments-panel")).not.toBeVisible();

    // Extension keybinding should not fire while typing in an input.
    await page.evaluate((isMac) => {
      const target = document.activeElement ?? window;
      target.dispatchEvent(
        new KeyboardEvent("keydown", {
          key: "y",
          metaKey: isMac,
          ctrlKey: !isMac,
          shiftKey: true,
          bubbles: true,
          cancelable: true,
        }),
      );
    }, process.platform === "darwin");
    await page.waitForTimeout(250);
    await expect(page.getByTestId("toast")).toHaveCount(0);
  });
});
