import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function waitForIdle(page: import("@playwright/test").Page): Promise<void> {
  await page.waitForFunction(() => Boolean((window as any).__formulaApp?.whenIdle), null, { timeout: 10_000 });
  await page.evaluate(() => (window as any).__formulaApp.whenIdle());
}

async function toggleShowFormulasShortcut(page: import("@playwright/test").Page): Promise<void> {
  const modifier = process.platform === "darwin" ? "Meta" : "Control";
  await page.keyboard.down(modifier);
  await page.keyboard.press("Backquote");
  await page.keyboard.up(modifier);
}

test.describe("ribbon view toggles", () => {
  test("Show Formulas toggle works from Ribbon and stays in sync with Ctrl/Cmd+`", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);

    // Seed a simple formula in C1 (SUM(A1:A2)) so we can verify render output.
    await page.click("#grid", { position: { x: 53, y: 29 } });
    await page.keyboard.press("F2");
    const cellEditor = page.locator("textarea.cell-editor");
    await cellEditor.fill("1");
    await page.keyboard.press("Enter"); // commits and moves to A2
    await waitForIdle(page);

    await page.keyboard.press("F2");
    await cellEditor.fill("2");
    await page.keyboard.press("Enter");
    await waitForIdle(page);

    // Create C1 = SUM(A1:A2).
    await page.click("#grid", { position: { x: 260, y: 40 } });
    await page.keyboard.press("F2");
    await cellEditor.fill("=SUM(A1:A2)");
    await page.keyboard.press("Enter");
    await waitForIdle(page);

    // Default: shows computed value.
    const viewTab = page.getByRole("tab", { name: "View", exact: true });
    await expect(viewTab).toBeVisible();
    await viewTab.click();
    const showFormulasToggle = page.getByTestId("ribbon-show-formulas");
    await expect(showFormulasToggle).toHaveAttribute("aria-pressed", "false");

    const defaultRenderText = await page.evaluate(() => (window as any).__formulaApp.getCellDisplayTextForRenderA1("C1"));
    expect(defaultRenderText).toBe("3");

    // Toggle via Ribbon: shows formula text and pressed state updates.
    await showFormulasToggle.click();
    await expect(showFormulasToggle).toHaveAttribute("aria-pressed", "true");

    const formulaRenderText = await page.evaluate(() => (window as any).__formulaApp.getCellDisplayTextForRenderA1("C1"));
    expect(formulaRenderText).toBe("=SUM(A1:A2)");

    // Toggle via keyboard shortcut: pressed state syncs back off.
    await page.click("#grid", { position: { x: 260, y: 40 } });
    await toggleShowFormulasShortcut(page);
    await expect(showFormulasToggle).toHaveAttribute("aria-pressed", "false");
    const toggledBackText = await page.evaluate(() => (window as any).__formulaApp.getCellDisplayTextForRenderA1("C1"));
    expect(toggledBackText).toBe("3");
  });

  test("Performance Stats toggle reflects app state (shared grid mode)", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");
    await waitForIdle(page);

    await page.getByRole("tab", { name: "View", exact: true }).click();
    const perfStatsToggle = page.getByTestId("ribbon-perf-stats");

    const initialEnabled = await page.evaluate(() => Boolean((window as any).__formulaApp.getGridPerfStats()?.enabled));
    await expect(perfStatsToggle).toHaveAttribute("aria-pressed", initialEnabled ? "true" : "false");

    await perfStatsToggle.click();
    await expect(perfStatsToggle).toHaveAttribute("aria-pressed", initialEnabled ? "false" : "true");
    const enabled = await page.evaluate(() => Boolean((window as any).__formulaApp.getGridPerfStats()?.enabled));
    expect(enabled).toBe(!initialEnabled);

    await perfStatsToggle.click();
    await expect(perfStatsToggle).toHaveAttribute("aria-pressed", initialEnabled ? "true" : "false");
    const disabled = await page.evaluate(() => Boolean((window as any).__formulaApp.getGridPerfStats()?.enabled));
    expect(disabled).toBe(initialEnabled);
  });
});
