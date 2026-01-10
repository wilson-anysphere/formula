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
    await page.click("#grid", { position: { x: 205, y: 5 } });
    await expect(page.getByTestId("active-cell")).toHaveText("C1");

    // Start editing in the formula bar.
    await page.getByTestId("formula-highlight").click();
    const input = page.getByTestId("formula-input");
    await expect(input).toBeVisible();

    await input.fill("=SUM(");

    // Drag select A1:A2 to insert a range reference.
    const gridBox = await page.locator("#grid").boundingBox();
    if (!gridBox) throw new Error("Missing grid bounding box");

    await page.mouse.move(gridBox.x + 5, gridBox.y + 5);
    await page.mouse.down();
    await page.mouse.move(gridBox.x + 5, gridBox.y + 29);
    await page.mouse.up();

    await expect(input).toHaveValue("=SUM(A1:A2");

    await page.keyboard.type(")");
    await page.keyboard.press("Enter");

    const c1Value = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("C1"));
    expect(c1Value).toBe("3");
  });
});
