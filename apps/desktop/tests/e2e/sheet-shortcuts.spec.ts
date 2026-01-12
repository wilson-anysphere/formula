import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("sheet navigation shortcuts", () => {
  test("Ctrl+PageDown / Ctrl+PageUp switches the active sheet (wraps)", async ({ page }) => {
    await gotoDesktop(page);
    await expect(page.getByTestId("sheet-tab-Sheet1")).toBeVisible();

    // Ensure the grid has focus by clicking the center of A1 once the layout is ready.
    await page.waitForFunction(() => {
      const app = (window as any).__formulaApp;
      const rect = app?.getCellRectA1?.("A1");
      return rect && rect.width > 0 && rect.height > 0;
    });
    const a1 = (await page.evaluate(() => (window as any).__formulaApp.getCellRectA1("A1"))) as {
      x: number;
      y: number;
      width: number;
      height: number;
    };
    await page
      .locator("#grid")
      .click({ position: { x: a1.x + a1.width / 2, y: a1.y + a1.height / 2 } });
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    // Lazily create Sheet2 by writing a value into it.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.getDocument().setCellValue("Sheet2", "A1", "Hello from Sheet2");
    });
    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();

    const ctrlKey = process.platform !== "darwin";
    const metaKey = process.platform === "darwin";

    const dispatch = async (key: "PageUp" | "PageDown") => {
      await page.evaluate(
        ({ key, ctrlKey, metaKey }) => {
          const grid = document.getElementById("grid");
          if (!grid) throw new Error("Missing #grid");
          grid.focus();
          const evt = new KeyboardEvent("keydown", { key, ctrlKey, metaKey, bubbles: true, cancelable: true });
          grid.dispatchEvent(evt);
        },
        { key, ctrlKey, metaKey },
      );
    };

    // Next sheet.
    await dispatch("PageDown");
    await expect(page.getByTestId("sheet-tab-Sheet2")).toHaveAttribute("data-active", "true");
    await expect(page.getByTestId("active-cell")).toHaveText("A1");
    await expect(page.getByTestId("active-value")).toHaveText("Hello from Sheet2");

    // Previous sheet.
    await dispatch("PageUp");
    await expect(page.getByTestId("sheet-tab-Sheet1")).toHaveAttribute("data-active", "true");
    await expect(page.getByTestId("active-cell")).toHaveText("A1");
    await expect(page.getByTestId("active-value")).toHaveText("Seed");

    // Wrap-around at the start.
    await dispatch("PageUp");
    await expect(page.getByTestId("sheet-tab-Sheet2")).toHaveAttribute("data-active", "true");

    // Wrap-around at the end.
    await dispatch("PageDown");
    await expect(page.getByTestId("sheet-tab-Sheet1")).toHaveAttribute("data-active", "true");
  });
});
