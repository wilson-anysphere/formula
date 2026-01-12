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

async function getDocumentSheetIdOrder(page: Page): Promise<string[]> {
  return await page.evaluate(() => {
    return (window as any).__formulaApp.getDocument().getSheetIds();
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

    // Status bar should reflect the active sheet (new sheets are activated on creation).
    await expect(page.getByTestId("sheet-position")).toHaveText("Sheet 3 of 3");

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
    //
    // We intentionally dispatch a synthetic `drop` event instead of using Playwright's
    // dragAndDrop helper. The desktop shell can be flaky under load when plumbing
    // `DataTransfer` through real pointer gestures, which can lead to hung/slow tests
    // unrelated to sheet reorder correctness.
    await page.evaluate(() => {
      const target = document.querySelector('[data-testid="sheet-tab-Sheet1"]') as HTMLElement | null;
      if (!target) throw new Error("Missing sheet-tab-Sheet1");
      const rect = target.getBoundingClientRect();

      const dt = new DataTransfer();
      dt.setData("text/sheet-id", "Sheet3");
      dt.setData("text/plain", "Sheet3");

      const drop = new DragEvent("drop", {
        bubbles: true,
        cancelable: true,
        // Drop near the left edge so the SheetTabStrip interprets this as "insert before".
        clientX: rect.left + 1,
        clientY: rect.top + rect.height / 2,
      });
      Object.defineProperty(drop, "dataTransfer", { value: dt });
      target.dispatchEvent(drop);
    });

    await expect
      .poll(async () => await getVisibleSheetTabOrder(page), { timeout: 2_000 })
      .toEqual(["Sheet3", "Sheet1", "Sheet2"]);

    // Sheet switcher ordering should match the visible tab strip ordering (store-driven).
    await expect
      .poll(async () => await getSheetSwitcherOptionValues(page), { timeout: 2_000 })
      .toEqual(["Sheet3", "Sheet1", "Sheet2"]);

    // DocumentController sheet iteration order should follow the tab reorder (used by versioning).
    await expect
      .poll(async () => await getDocumentSheetIdOrder(page), { timeout: 2_000 })
      .toEqual(["Sheet3", "Sheet1", "Sheet2"]);

    // Status bar should update to match the new position of the active sheet.
    await expect(page.getByTestId("sheet-position")).toHaveText("Sheet 1 of 3");

    // Reordering should not switch the active sheet; the dragged sheet remains active.
    await expect(page.getByTestId("sheet-tab-Sheet3")).toHaveAttribute("data-active", "true");

    // Reordering should restore focus back to the grid (Excel-like flow).
    await expect(page.locator("#grid")).toBeFocused();

    // Dirty tracking: sheet reorder is a metadata edit and should still prompt for saving.
    expect(await page.evaluate(() => (window as any).__formulaApp.getDocument().isDirty)).toBe(true);
  });
});
