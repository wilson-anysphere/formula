import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("sheet tabs", () => {
  test("switching sheets updates the visible cell values", async ({ page }) => {
    await gotoDesktop(page);

    // Ensure A1 is active before switching sheets so the status bar reflects A1 values.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.activateCell({ row: 0, col: 0 });
    });
    await expect(page.getByTestId("sheet-tab-Sheet1")).toBeVisible();
    await expect(page.getByTestId("sheet-position")).toHaveText("Sheet 1 of 1");

    // Lazily create Sheet2 by writing a value into it.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.getDocument().setCellValue("Sheet2", "A1", "Hello from Sheet2");
    });

    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();
    await expect(page.getByTestId("sheet-position")).toHaveText("Sheet 1 of 2");

    await page.getByTestId("sheet-tab-Sheet2").click();
    await expect(page.getByTestId("active-cell")).toHaveText("A1");
    await expect(page.getByTestId("active-value")).toHaveText("Hello from Sheet2");
    await expect(page.getByTestId("sheet-position")).toHaveText("Sheet 2 of 2");

    // Switching back restores the original Sheet1 value.
    await page.getByTestId("sheet-tab-Sheet1").click();
    await expect(page.getByTestId("active-value")).toHaveText("Seed");
    await expect(page.getByTestId("sheet-position")).toHaveText("Sheet 1 of 2");
  });

  test("add sheet button creates and activates the next SheetN tab", async ({ page }) => {
    await gotoDesktop(page);

    // Ensure A1 is active so the status bar is deterministic after the sheet switch.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.activateCell({ row: 0, col: 0 });
    });

    const nextSheetId = await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const ids = app.getDocument().getSheetIds();
      const existing = new Set((ids.length > 0 ? ids : ["Sheet1"]) as string[]);
      let n = 1;
      while (existing.has(`Sheet${n}`)) n += 1;
      return `Sheet${n}`;
    });

    await page.getByTestId("sheet-add").click();

    const newTab = page.getByTestId(`sheet-tab-${nextSheetId}`);
    await expect(newTab).toBeVisible();
    await expect(newTab).toHaveAttribute("data-active", "true");

    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCurrentSheetId())).toBe(nextSheetId);

    // Sheet activation should return focus to the grid so keyboard shortcuts keep working.
    await expect
      .poll(() => page.evaluate(() => (document.activeElement as HTMLElement | null)?.id))
      .toBe("grid");

    // Verify the new sheet is functional by writing a value into A1 and observing the status bar update.
    await page.evaluate((sheetId) => {
      const app = (window as any).__formulaApp;
      app.getDocument().setCellValue(sheetId, "A1", `Hello from ${sheetId}`);
    }, nextSheetId);

    await expect(page.getByTestId("active-cell")).toHaveText("A1");
    await expect(page.getByTestId("active-value")).toHaveText(`Hello from ${nextSheetId}`);
  });

  test("drag reordering sheet tabs updates Ctrl+PgUp/PgDn navigation order", async ({ page }) => {
    await gotoDesktop(page);

    // Ensure A1 is active so Ctrl+PgUp/PgDn starts from a deterministic sheet/cell.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.activateCell({ row: 0, col: 0 });
    });

    // Create Sheet2 + Sheet3 via the UI.
    await page.getByTestId("sheet-add").click();
    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();
    await page.getByTestId("sheet-add").click();
    await expect(page.getByTestId("sheet-tab-Sheet3")).toBeVisible();
    await expect(page.getByTestId("sheet-tab-Sheet3")).toHaveAttribute("data-active", "true");

    // Move Sheet3 before Sheet1.
    try {
      await page
        .getByTestId("sheet-tab-Sheet3")
        .dragTo(page.getByTestId("sheet-tab-Sheet1"), { targetPosition: { x: 1, y: 1 } });
    } catch {
      // Ignore; we'll fall back to a synthetic drop below.
    }

    // Playwright drag/drop can be flaky with HTML5 DataTransfer. If the order doesn't match,
    // dispatch a synthetic drop event that exercises the sheet tab DnD plumbing.
    const desiredOrder = ["Sheet3", "Sheet1", "Sheet2"];
    const orderKey = (order: Array<string | null>) => order.filter(Boolean).slice(0, 3).join(",");
    const didReorder = orderKey(
      await page.evaluate(() =>
        Array.from(document.querySelectorAll("#sheet-tabs .sheet-tabs [data-sheet-id]")).map((el) =>
          (el as HTMLElement).getAttribute("data-sheet-id"),
        ),
      ),
    );

    if (didReorder !== desiredOrder.join(",")) {
      await page.evaluate(() => {
        const fromId = "Sheet3";
        const target = document.querySelector('[data-testid="sheet-tab-Sheet1"]') as HTMLElement | null;
        if (!target) throw new Error("Missing Sheet1 tab");
        const rect = target.getBoundingClientRect();

        const dt = new DataTransfer();
        dt.setData("text/sheet-id", fromId);

        const drop = new DragEvent("drop", {
          bubbles: true,
          cancelable: true,
          clientX: rect.left + 1,
          clientY: rect.top + rect.height / 2,
        });
        Object.defineProperty(drop, "dataTransfer", { value: dt });
        target.dispatchEvent(drop);
      });
    }

    await expect.poll(() =>
      page.evaluate(() =>
        Array.from(document.querySelectorAll("#sheet-tabs .sheet-tabs [data-sheet-id]")).map((el) =>
          (el as HTMLElement).getAttribute("data-sheet-id"),
        ),
      ),
    ).toEqual(desiredOrder);

    // Ctrl+PgDn should follow the new order. We dispatch the key event directly to avoid
    // platform/browser-specific tab switching behavior.
    await page.evaluate(() => {
      const evt = new KeyboardEvent("keydown", { key: "PageDown", ctrlKey: true, bubbles: true, cancelable: true });
      window.dispatchEvent(evt);
    });
    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCurrentSheetId())).toBe("Sheet1");

    await page.evaluate(() => {
      const evt = new KeyboardEvent("keydown", { key: "PageDown", ctrlKey: true, bubbles: true, cancelable: true });
      window.dispatchEvent(evt);
    });
    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCurrentSheetId())).toBe("Sheet2");

    await page.evaluate(() => {
      const evt = new KeyboardEvent("keydown", { key: "PageDown", ctrlKey: true, bubbles: true, cancelable: true });
      window.dispatchEvent(evt);
    });
    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCurrentSheetId())).toBe("Sheet3");
  });
});
