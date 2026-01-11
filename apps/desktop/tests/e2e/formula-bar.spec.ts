import { expect, test } from "@playwright/test";

test.describe("formula bar editing + range insertion", () => {
  test("type formula, drag range, commit stores formula in the active cell", async ({ page }) => {
    await page.goto("/");

    // Seed numeric inputs in A1 and A2 (so SUM has a visible result).
    await page.click("#grid", { position: { x: 5, y: 5 } });
    await page.keyboard.press("F2");
    const cellEditor = page.locator("textarea.cell-editor");
    await cellEditor.fill("1");
    await page.keyboard.press("Enter"); // commits and moves to A2

    await page.keyboard.press("F2");
    await cellEditor.fill("2");
    await page.keyboard.press("Enter");

    // Select C1.
    await page.click("#grid", { position: { x: 253, y: 29 } });
    await expect(page.getByTestId("active-cell")).toHaveText("C1");

    // Start editing in the formula bar.
    await page.getByTestId("formula-highlight").click();
    const input = page.getByTestId("formula-input");
    await expect(input).toBeVisible();

    await input.fill("=SUM(");

    // Drag select A1:A2 to insert a range reference.
    const gridBox = await page.locator("#grid").boundingBox();
    if (!gridBox) throw new Error("Missing grid bounding box");

    // Drag from A1 to A2 inside the cell region (below the column header).
    await page.mouse.move(gridBox.x + 53, gridBox.y + 29);
    await page.mouse.down();
    await page.mouse.move(gridBox.x + 53, gridBox.y + 53);
    await page.mouse.up();

    await expect(input).toHaveValue("=SUM(A1:A2");

    await page.keyboard.type(")");
    await page.keyboard.press("Enter");

    const c1Value = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("C1"));
    expect(c1Value).toBe("3");
  });

  test("shows friendly error explanation for #DIV/0!", async ({ page }) => {
    await page.goto("/");

    // Seed A1 = 0.
    await page.click("#grid", { position: { x: 5, y: 5 } });
    await page.keyboard.press("F2");
    const cellEditor = page.locator("textarea.cell-editor");
    await cellEditor.fill("0");
    await page.keyboard.press("Enter");

    // Select B1.
    await page.click("#grid", { position: { x: 153, y: 29 } });
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
