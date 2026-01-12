import { expect, test, type Page } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function getVisibleSheetTabOrder(page: Page): Promise<string[]> {
  return await page.evaluate(() => {
    const strip = document.querySelector(".sheet-tabs");
    if (!strip) throw new Error("Missing .sheet-tabs");
    return Array.from(strip.children)
      .map((child) => (child as HTMLElement).dataset.sheetId ?? "")
      .filter(Boolean);
  });
}

async function getSheetSwitcherOptionValues(page: Page): Promise<string[]> {
  return await page.evaluate(() => {
    const select = document.querySelector('[data-testid="sheet-switcher"]') as HTMLSelectElement | null;
    if (!select) throw new Error("Missing sheet switcher");
    return Array.from(select.querySelectorAll("option")).map((opt) => (opt as HTMLOptionElement).value);
  });
}

test.describe("sheet tabs", () => {
  test("drag-and-drop reorders tabs and marks the document dirty", async ({ page }) => {
    await gotoDesktop(page);

    await expect(page.getByTestId("sheet-tab-Sheet1")).toBeVisible();

    // Create 3 sheets via the "+" button (Sheet1 is seeded by default).
    await page.getByTestId("sheet-add").click();
    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();
    await page.getByTestId("sheet-add").click();
    await expect(page.getByTestId("sheet-tab-Sheet3")).toBeVisible();

    await expect
      .poll(async () => await getVisibleSheetTabOrder(page), { timeout: 5_000 })
      .toEqual(["Sheet1", "Sheet2", "Sheet3"]);

    await expect
      .poll(async () => await getSheetSwitcherOptionValues(page), { timeout: 5_000 })
      .toEqual(["Sheet1", "Sheet2", "Sheet3"]);

    // Ensure the dirty assertion is specifically caused by sheet reordering.
    await page.evaluate(() => {
      (window as any).__formulaApp.getDocument().markSaved();
    });
    expect(await page.evaluate(() => (window as any).__formulaApp.getDocument().isDirty)).toBe(false);

    // Simulate a typical user flow: focus the dragged tab before starting the drag gesture.
    await page.getByTestId("sheet-tab-Sheet3").focus();
    await expect(page.getByTestId("sheet-tab-Sheet3")).toBeFocused();
    await expect(page.getByTestId("sheet-tab-Sheet3")).toHaveAttribute("data-active", "true");

    // Drag Sheet3 onto Sheet1.
    await page.dragAndDrop('[data-testid="sheet-tab-Sheet3"]', '[data-testid="sheet-tab-Sheet1"]', {
      // Drop near the left edge so the SheetTabStrip interprets this as "insert before".
      targetPosition: { x: 5, y: 5 },
    });

    // Wait for the UI to re-render in the new order. If Playwright's dragAndDrop
    // fails to plumb dataTransfer, fall back to a synthetic drop event.
    try {
      await expect
        .poll(async () => await getVisibleSheetTabOrder(page), { timeout: 2_000 })
        .toEqual(["Sheet3", "Sheet1", "Sheet2"]);
    } catch {
      await page.evaluate(() => {
        const target = document.querySelector('[data-testid="sheet-tab-Sheet1"]');
        if (!target) throw new Error("Missing sheet-tab-Sheet1");
        const dt = new DataTransfer();
        dt.setData("text/plain", "Sheet3");
        const drop = new DragEvent("drop", { bubbles: true, cancelable: true });
        Object.defineProperty(drop, "dataTransfer", { value: dt });
        target.dispatchEvent(drop);
      });

      await expect
        .poll(async () => await getVisibleSheetTabOrder(page), { timeout: 2_000 })
        .toEqual(["Sheet3", "Sheet1", "Sheet2"]);
    }

    // Sheet switcher ordering should match the visible tab strip ordering (store-driven).
    await expect
      .poll(async () => await getSheetSwitcherOptionValues(page), { timeout: 2_000 })
      .toEqual(["Sheet3", "Sheet1", "Sheet2"]);

    // Reordering should not switch the active sheet; the dragged sheet remains active.
    await expect(page.getByTestId("sheet-tab-Sheet3")).toHaveAttribute("data-active", "true");

    // Reordering should restore focus back to the grid (Excel-like flow).
    await expect(page.locator("#grid")).toBeFocused();

    // Dirty tracking: sheet reorder is a metadata edit and should still prompt for saving.
    expect(await page.evaluate(() => (window as any).__formulaApp.getDocument().isDirty)).toBe(true);
  });
});
