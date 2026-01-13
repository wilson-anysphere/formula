import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function waitForIdle(page: import("@playwright/test").Page): Promise<void> {
  await page.waitForFunction(() => Boolean((window.__formulaApp as any)?.whenIdle), null, { timeout: 10_000 });
  await page.evaluate(() => (window.__formulaApp as any).whenIdle());
}

async function toggleShowFormulasShortcut(page: import("@playwright/test").Page): Promise<void> {
  const modifier = process.platform === "darwin" ? "Meta" : "Control";
  await page.keyboard.down(modifier);
  await page.keyboard.press("Backquote");
  await page.keyboard.up(modifier);
}

test.describe("ribbon view toggles", () => {
  test("clicking a ribbon control commits an in-progress in-cell edit without stealing focus", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");
    await waitForIdle(page);

    // Focus/select A1.
    await page.click("#grid", { position: { x: 80, y: 40 } });
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    // Start editing A1 but do not press Enter.
    await page.keyboard.press("h");
    const editor = page.locator("textarea.cell-editor");
    await expect(editor).toBeVisible();
    await page.keyboard.type("ello");
    await expect(editor).toHaveValue("hello");

    const ribbon = page.getByTestId("ribbon-root");
    const viewTab = ribbon.getByRole("tab", { name: "View", exact: true });
    await expect(viewTab).toBeVisible();

    // Clicking the ribbon should blur/commit the edit, and focus should remain on the ribbon
    // control (no focus ping-pong back to the grid).
    await viewTab.click();
    await expect(viewTab).toBeFocused();
    await waitForIdle(page);

    expect(await page.evaluate(() => (window.__formulaApp as any).getCellValueA1("A1"))).toBe("hello");
    await expect(editor).toBeHidden();
    await expect(viewTab).toBeFocused();
  });

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
    const showFormulasToggle = page.getByTestId("ribbon-root").getByTestId("ribbon-show-formulas");
    await expect(showFormulasToggle).toHaveAttribute("aria-pressed", "false");

    const defaultRenderText = await page.evaluate(() => (window.__formulaApp as any).getCellDisplayTextForRenderA1("C1"));
    expect(defaultRenderText).toBe("3");

    // Toggle via Ribbon: shows formula text and pressed state updates.
    await showFormulasToggle.click();
    await expect(showFormulasToggle).toHaveAttribute("aria-pressed", "true");

    const formulaRenderText = await page.evaluate(() => (window.__formulaApp as any).getCellDisplayTextForRenderA1("C1"));
    expect(formulaRenderText).toBe("=SUM(A1:A2)");

    // Formulas tab has its own Show Formulas toggle; it should stay in sync with the View tab toggle.
    const formulasTab = page.getByRole("tab", { name: "Formulas", exact: true });
    await expect(formulasTab).toBeVisible();
    await formulasTab.click();
    const formulasShowFormulasToggle = page
      .getByRole("tabpanel", { name: "Formulas", exact: true })
      .getByRole("button", { name: "Show Formulas", exact: true });
    await expect(formulasShowFormulasToggle).toHaveAttribute("aria-pressed", "true");

    // Toggle via the Formulas tab control: should hide formulas and sync back to the View tab toggle.
    await formulasShowFormulasToggle.click();
    await waitForIdle(page);
    await expect(formulasShowFormulasToggle).toHaveAttribute("aria-pressed", "false");
    const ribbonToggledOffText = await page.evaluate(() => (window.__formulaApp as any).getCellDisplayTextForRenderA1("C1"));
    expect(ribbonToggledOffText).toBe("3");

    await viewTab.click();
    await expect(showFormulasToggle).toHaveAttribute("aria-pressed", "false");

    // Toggle via keyboard shortcut: pressed state syncs back on.
    await page.click("#grid", { position: { x: 260, y: 40 } });
    await toggleShowFormulasShortcut(page);
    await waitForIdle(page);
    await expect(showFormulasToggle).toHaveAttribute("aria-pressed", "true");
    await formulasTab.click();
    await expect(formulasShowFormulasToggle).toHaveAttribute("aria-pressed", "true");
    const toggledBackText = await page.evaluate(() => (window.__formulaApp as any).getCellDisplayTextForRenderA1("C1"));
    expect(toggledBackText).toBe("=SUM(A1:A2)");
  });

  test("Performance Stats toggle reflects app state (shared grid mode)", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");
    await waitForIdle(page);

    await page.getByRole("tab", { name: "View", exact: true }).click();
    const perfStatsToggle = page.getByTestId("ribbon-root").getByTestId("ribbon-perf-stats");

    const initialEnabled = await page.evaluate(() => Boolean((window.__formulaApp as any).getGridPerfStats()?.enabled));
    await expect(perfStatsToggle).toHaveAttribute("aria-pressed", initialEnabled ? "true" : "false");

    await perfStatsToggle.click();
    await expect(perfStatsToggle).toHaveAttribute("aria-pressed", initialEnabled ? "false" : "true");
    const enabled = await page.evaluate(() => Boolean((window.__formulaApp as any).getGridPerfStats()?.enabled));
    expect(enabled).toBe(!initialEnabled);

    await perfStatsToggle.click();
    await expect(perfStatsToggle).toHaveAttribute("aria-pressed", initialEnabled ? "true" : "false");
    const disabled = await page.evaluate(() => Boolean((window.__formulaApp as any).getGridPerfStats()?.enabled));
    expect(disabled).toBe(initialEnabled);
  });
});
