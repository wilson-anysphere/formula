import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function waitForIdle(page: import("@playwright/test").Page): Promise<void> {
  await page.evaluate(() => (window as any).__formulaApp.whenIdle());
}

test.describe("keybinding barriers", () => {
  test("command palette isolates Ctrl/Cmd+B (does not toggle bold)", async ({ page }) => {
    await gotoDesktop(page);

    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const sheetId = app.getCurrentSheetId();
      const doc = app.getDocument();
      doc.setRangeFormat(sheetId, "A1", { font: { bold: false }, numberFormat: null }, { label: "Reset" });
      app.selectRange({ range: { startRow: 0, endRow: 0, startCol: 0, endCol: 0 } });
      app.focus();
    });
    await waitForIdle(page);

    // Open command palette.
    await page.keyboard.press("ControlOrMeta+Shift+P");
    await expect(page.getByTestId("command-palette")).toBeVisible();

    // Move focus off the input so this assertion is not relying on "ignore text inputs" heuristics.
    await page.keyboard.press("Tab");
    await expect(page.getByTestId("command-palette-list")).toBeFocused();

    // Ctrl/Cmd+B should not leak into the grid while the palette is open.
    await page.keyboard.press("ControlOrMeta+B");
    await waitForIdle(page);

    await expect
      .poll(async () => {
        await waitForIdle(page);
        return await page.evaluate(() => {
          const app = (window as any).__formulaApp;
          const sheetId = app.getCurrentSheetId();
          const doc = app.getDocument();
          return Boolean(doc.getCellFormat(sheetId, "A1").font?.bold);
        });
      })
      .toBe(false);
  });

  test("context menu isolates Ctrl/Cmd+PageUp/PageDown (does not switch sheets)", async ({ page }) => {
    await gotoDesktop(page);

    // Lazily create Sheet2 so sheet navigation is observable.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.getDocument().setCellValue("Sheet2", "A1", "Hello from Sheet2");
    });

    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();
    await page.getByTestId("sheet-tab-Sheet1").click();
    await expect(page.getByTestId("sheet-tab-Sheet1")).toHaveAttribute("data-active", "true");

    // Focus the grid then open the context menu with Shift+F10.
    await page.click("#grid", { position: { x: 80, y: 40 } });
    await page.keyboard.press("Shift+F10");
    const menu = page.getByTestId("context-menu");
    await expect(menu).toBeVisible();
    // Ensure keyboard events originate from within the menu (some WebView environments
    // may briefly keep focus on the grid right after opening via Shift+F10).
    const firstItem = menu.locator(".context-menu__item:not(:disabled)").first();
    await firstItem.focus();
    await expect(firstItem).toBeFocused();

    const isMac = process.platform === "darwin";
    const dispatch = async (key: "PageUp" | "PageDown") => {
      await page.evaluate(
        ({ key, isMac }) => {
          const target = document.activeElement ?? window;
          target.dispatchEvent(
            new KeyboardEvent("keydown", {
              key,
              metaKey: isMac,
              ctrlKey: !isMac,
              bubbles: true,
              cancelable: true,
            }),
          );
        },
        { key, isMac },
      );
    };

    await dispatch("PageDown");
    await expect(page.getByTestId("sheet-tab-Sheet1")).toHaveAttribute("data-active", "true");

    await dispatch("PageUp");
    await expect(page.getByTestId("sheet-tab-Sheet1")).toHaveAttribute("data-active", "true");
  });
});
