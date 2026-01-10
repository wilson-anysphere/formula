import { expect, test } from "@playwright/test";

test.describe("grid keyboard navigation + in-place editing", () => {
  test("arrow navigation updates active cell", async ({ page }) => {
    await page.goto("/");

    // Focus + select A1 (top-left).
    await page.click("#grid", { position: { x: 5, y: 5 } });

    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    await page.keyboard.press("ArrowRight");
    await page.keyboard.press("ArrowRight");
    await expect(page.getByTestId("active-cell")).toHaveText("C1");

    await page.keyboard.press("ArrowDown");
    await expect(page.getByTestId("active-cell")).toHaveText("C2");
  });

  test("Ctrl+End jumps to last used cell", async ({ page, browserName }) => {
    // WebKit on Linux can be flaky with End/Home in headless; keep the coverage
    // focused on chromium/firefox in CI.
    test.skip(browserName === "webkit", "End/Home key handling is unreliable in webkit headless");

    await page.goto("/");
    await page.click("#grid", { position: { x: 5, y: 5 } });

    await page.keyboard.press("Control+End");
    await expect(page.getByTestId("active-cell")).toHaveText("D5");
  });

  test("F2 edit commits with Enter and cancels with Escape", async ({ page }) => {
    await page.goto("/");
    await page.click("#grid", { position: { x: 5, y: 5 } });

    // Edit A1.
    await page.keyboard.press("F2");
    const editor = page.locator("textarea.cell-editor");
    await expect(editor).toBeVisible();

    await editor.fill("Hello");
    await page.keyboard.press("Enter");

    // Commit moves down.
    await expect(page.getByTestId("active-cell")).toHaveText("A2");

    const a1Value = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"));
    expect(a1Value).toBe("Hello");

    // Start editing A2 but cancel.
    await page.keyboard.press("F2");
    await expect(editor).toBeVisible();
    await editor.fill("ShouldNotCommit");
    await page.keyboard.press("Escape");

    const a2Value = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A2"));
    expect(a2Value).toBe("");
    await expect(page.getByTestId("active-cell")).toHaveText("A2");
  });
});
