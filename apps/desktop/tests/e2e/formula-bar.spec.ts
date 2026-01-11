import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("formula bar editing + range insertion", () => {
  test("type formula, drag range, commit stores formula in the active cell", async ({ page }) => {
    await gotoDesktop(page);

    // Seed numeric inputs in A1 and A2 (so SUM has a visible result).
    // Click within the first grid cell (accounting for row/column headers).
    await page.click("#grid", { position: { x: 53, y: 29 } });
    await page.keyboard.press("F2");
    const cellEditor = page.locator("textarea.cell-editor");
    await cellEditor.fill("1");
    await page.keyboard.press("Enter"); // commits and moves to A2

    await page.keyboard.press("F2");
    await cellEditor.fill("2");
    await page.keyboard.press("Enter");

    // Select C1.
    // Account for the row/column headers rendered inside the grid canvas.
    await page.click("#grid", { position: { x: 260, y: 40 } });
    await expect(page.getByTestId("active-cell")).toHaveText("C1");

    // Start editing in the formula bar.
    await page.getByTestId("formula-highlight").click();
    const input = page.getByTestId("formula-input");
    await expect(input).toBeVisible();

    await input.fill("=SUM(");

    // Drag select A1:A2 to insert a range reference.
    const gridBox = await page.locator("#grid").boundingBox();
    if (!gridBox) throw new Error("Missing grid bounding box");

    // Drag from A1 to A2.
    await page.mouse.move(gridBox.x + 60, gridBox.y + 40);
    await page.mouse.down();
    await page.mouse.move(gridBox.x + 60, gridBox.y + 64);
    await page.mouse.up();

    await expect(input).toHaveValue("=SUM(A1:A2");

    await page.keyboard.type(")");
    await page.keyboard.press("Enter");

    // Ensure the numeric inputs landed where expected.
    const { a1Value, a2Value, c1Formula } = await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const doc = app.getDocument();
      return {
        a1Value: doc.getCell("Sheet1", "A1").value,
        a2Value: doc.getCell("Sheet1", "A2").value,
        c1Formula: doc.getCell("Sheet1", "C1").formula,
      };
    });
    expect(a1Value).toBe(1);
    expect(a2Value).toBe(2);
    expect(c1Formula).toBe("=SUM(A1:A2)");

    const c1Value = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("C1"));
    expect(c1Value).toBe("3");
  });

  test("picking a range on another sheet inserts a sheet-qualified reference and commits to the original edit cell", async ({ page }) => {
    await page.goto("/");
    await page.waitForFunction(() => (window as any).__formulaApp != null);

    // Lazily create Sheet2 and seed A1.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.getDocument().setCellValue("Sheet2", "A1", 7);
    });
    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();

    // Start editing on Sheet1!C1.
    await page.click("#grid", { position: { x: 260, y: 40 } });
    await expect(page.getByTestId("active-cell")).toHaveText("C1");

    await page.getByTestId("formula-highlight").click();
    const input = page.getByTestId("formula-input");
    await expect(input).toBeVisible();
    await input.fill("=");

    // Switch to Sheet2 while still editing and pick A1.
    await page.getByTestId("sheet-tab-Sheet2").click();
    await expect(page.getByTestId("sheet-tab-Sheet2")).toHaveAttribute("data-active", "true");
    await page.click("#grid", { position: { x: 53, y: 29 } });
    await expect(input).toHaveValue("=Sheet2!A1");

    // Commit should apply to the original edit cell (Sheet1!C1) and restore the sheet.
    await input.focus();
    await page.keyboard.press("Enter");
    await expect(page.getByTestId("sheet-tab-Sheet1")).toHaveAttribute("data-active", "true");
    await expect(page.getByTestId("active-cell")).toHaveText("C1");

    const { sheet1Formula, sheet2Formula, sheet2Value } = await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const doc = app.getDocument();
      return {
        sheet1Formula: doc.getCell("Sheet1", "C1").formula,
        sheet2Formula: doc.getCell("Sheet2", "A1").formula,
        sheet2Value: doc.getCell("Sheet2", "A1").value,
      };
    });

    expect(sheet1Formula).toBe("=Sheet2!A1");
    expect(sheet2Formula).toBeNull();
    expect(sheet2Value).toBe(7);

    const c1Value = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("C1"));
    expect(c1Value).toBe("7");
  });

  test("canceling after switching sheets restores the original edit location without applying edits", async ({ page }) => {
    await page.goto("/");
    await page.waitForFunction(() => (window as any).__formulaApp != null);

    // Lazily create Sheet2 and seed A1.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.getDocument().setCellValue("Sheet2", "A1", 7);
    });
    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();

    // Start editing on Sheet1!C1.
    await page.click("#grid", { position: { x: 260, y: 40 } });
    await expect(page.getByTestId("active-cell")).toHaveText("C1");

    await page.getByTestId("formula-highlight").click();
    const input = page.getByTestId("formula-input");
    await expect(input).toBeVisible();
    await input.fill("=");

    // Switch to Sheet2 and pick A1 to insert a reference.
    await page.getByTestId("sheet-tab-Sheet2").click();
    await expect(page.getByTestId("sheet-tab-Sheet2")).toHaveAttribute("data-active", "true");
    await page.click("#grid", { position: { x: 53, y: 29 } });
    await expect(input).toHaveValue("=Sheet2!A1");

    // Cancel should restore Sheet1!C1 and leave the cell unchanged.
    await input.focus();
    await page.keyboard.press("Escape");
    await expect(page.getByTestId("sheet-tab-Sheet1")).toHaveAttribute("data-active", "true");
    await expect(page.getByTestId("active-cell")).toHaveText("C1");

    const sheet1Formula = await page.evaluate(() => (window as any).__formulaApp.getDocument().getCell("Sheet1", "C1").formula);
    expect(sheet1Formula).toBeNull();
  });

  test("shows friendly error explanation for #DIV/0!", async ({ page }) => {
    await gotoDesktop(page);

    // Seed A1 = 0.
    await page.click("#grid", { position: { x: 53, y: 29 } });
    await page.keyboard.press("F2");
    const cellEditor = page.locator("textarea.cell-editor");
    await cellEditor.fill("0");
    await page.keyboard.press("Enter");

    // Select B1.
    await page.click("#grid", { position: { x: 160, y: 40 } });
    await expect(page.getByTestId("active-cell")).toHaveText("B1");

    // Enter a division-by-zero formula.
    await page.getByTestId("formula-highlight").click();
    const input = page.getByTestId("formula-input");
    await expect(input).toBeVisible();
    await input.fill("=1/A1");
    await page.keyboard.press("Enter");

    // Error button should appear and panel should explain.
    const errorButton = page.getByTestId("formula-error-button");
    await expect(errorButton).toBeVisible();
    await errorButton.click();
    await expect(page.getByTestId("formula-error-panel")).toBeVisible();
    await expect(page.getByTestId("formula-error-panel")).toContainText("Division by zero");
  });
});
