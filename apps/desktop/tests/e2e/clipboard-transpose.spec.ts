import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function waitForGridFocus(page: import("@playwright/test").Page): Promise<void> {
  await page.waitForFunction(() => (document.activeElement as HTMLElement | null)?.id === "grid", null, { timeout: 5_000 });
}

test.describe("clipboard transpose", () => {
  test("Paste → Transpose transposes the copied grid and updates the selection", async ({ page }) => {
    await page.context().grantPermissions(["clipboard-read", "clipboard-write"]);
    await gotoDesktop(page);

    const ribbon = page.getByTestId("ribbon-root");
    await ribbon.getByRole("tab", { name: "Home" }).click();

    // Seed a 2x3 rectangle (A1:C2) with distinct values.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      doc.beginBatch({ label: "Seed transpose clipboard cells" });
      doc.setCellValue(sheetId, "A1", "A1");
      doc.setCellValue(sheetId, "B1", "B1");
      doc.setCellValue(sheetId, "C1", "C1");
      doc.setCellValue(sheetId, "A2", "A2");
      doc.setCellValue(sheetId, "B2", "B2");
      doc.setCellValue(sheetId, "C2", "C2");
      doc.endBatch();
      app.refresh();
    });

    // Select A1:C2 and copy.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const sheetId = app.getCurrentSheetId();
      app.selectRange({
        sheetId,
        range: { startRow: 0, startCol: 0, endRow: 1, endCol: 2 }, // A1:C2
      });
    });
    await expect(page.getByTestId("selection-range")).toHaveText("A1:C2");

    await ribbon.getByRole("button", { name: "Copy" }).click();
    await waitForGridFocus(page);

    // Paste transpose into E1.
    await page.click("#grid", { position: { x: 460, y: 40 } }); // E1
    await expect(page.getByTestId("active-cell")).toHaveText("E1");

    await ribbon.getByTestId("ribbon-paste").click();
    await ribbon.getByRole("menuitem", { name: "Transpose" }).click();
    await waitForGridFocus(page);

    // Selection should match the transposed 3x2 output rectangle (E1:F3).
    await expect(page.getByTestId("selection-range")).toHaveText("E1:F3");

    // Values should be transposed (rows/cols swapped).
    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCellValueA1("E1"))).toBe("A1");
    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCellValueA1("F1"))).toBe("A2");
    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCellValueA1("E2"))).toBe("B1");
    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCellValueA1("F2"))).toBe("B2");
    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCellValueA1("E3"))).toBe("C1");
    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCellValueA1("F3"))).toBe("C2");
  });
});

