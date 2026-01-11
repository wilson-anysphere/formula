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

  test("typing begins in-place edit and commits", async ({ page }) => {
    await page.goto("/");
    await page.click("#grid", { position: { x: 5, y: 5 } });

    const recalcBefore = await page.evaluate(() => (window as any).__formulaApp.getRecalcCount());

    // Type directly without pressing F2 first.
    await page.keyboard.press("h");
    const editor = page.locator("textarea.cell-editor");
    await expect(editor).toBeVisible();

    await page.keyboard.type("ello");
    await page.keyboard.press("Enter");

    await expect(page.getByTestId("active-cell")).toHaveText("A2");
    const a1Value = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"));
    expect(a1Value).toBe("hello");

    const recalcAfter = await page.evaluate(() => (window as any).__formulaApp.getRecalcCount());
    expect(recalcAfter).toBeGreaterThan(recalcBefore);
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

    const recalcBefore = await page.evaluate(() => (window as any).__formulaApp.getRecalcCount());

    // Edit A1.
    await page.keyboard.press("F2");
    const editor = page.locator("textarea.cell-editor");
    await expect(editor).toBeVisible();

    await editor.fill("Hello");
    await page.keyboard.press("Enter");

    // Commit moves down.
    await expect(page.getByTestId("active-cell")).toHaveText("A2");

    const recalcAfterCommit = await page.evaluate(() => (window as any).__formulaApp.getRecalcCount());
    expect(recalcAfterCommit).toBeGreaterThan(recalcBefore);

    const a1Value = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"));
    expect(a1Value).toBe("Hello");

    // Start editing A2 but cancel.
    await page.keyboard.press("F2");
    await expect(editor).toBeVisible();
    await editor.fill("ShouldNotCommit");
    await page.keyboard.press("Escape");

    const recalcAfterCancel = await page.evaluate(() => (window as any).__formulaApp.getRecalcCount());
    expect(recalcAfterCancel).toBe(recalcAfterCommit);

    const a2Value = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A2"));
    expect(a2Value).toBe("A");
    await expect(page.getByTestId("active-cell")).toHaveText("A2");

    // Delete clears cell contents.
    await page.keyboard.press("ArrowUp");
    await expect(page.getByTestId("active-cell")).toHaveText("A1");
    const recalcBeforeDelete = await page.evaluate(() => (window as any).__formulaApp.getRecalcCount());
    await page.keyboard.press("Delete");
    const a1Cleared = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"));
    expect(a1Cleared).toBe("");
    const recalcAfterDelete = await page.evaluate(() => (window as any).__formulaApp.getRecalcCount());
    expect(recalcAfterDelete).toBeGreaterThan(recalcBeforeDelete);
  });

  test("Ctrl/Cmd+Z undo + Ctrl/Cmd+Shift+Z redo update the document", async ({ page }) => {
    await page.goto("/");
    await page.click("#grid", { position: { x: 5, y: 5 } });

    // Clear any seeded value so undo returns to an empty cell.
    await page.keyboard.press("Delete");

    // Edit A1.
    await page.keyboard.press("F2");
    const editor = page.locator("textarea.cell-editor");
    await expect(editor).toBeVisible();
    await editor.fill("Hello");
    await page.keyboard.press("Enter");

    const modifier = process.platform === "darwin" ? "Meta" : "Control";

    await page.keyboard.press(`${modifier}+Z`);
    const a1AfterUndo = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"));
    expect(a1AfterUndo).toBe("");

    await page.keyboard.press(`${modifier}+Shift+Z`);
    const a1AfterRedo = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"));
    expect(a1AfterRedo).toBe("Hello");
  });
});
