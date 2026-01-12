import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function waitForIdle(page: import("@playwright/test").Page): Promise<void> {
  // Vite may occasionally trigger a one-time full reload after dependency optimization.
  // Retry once if the execution context is destroyed mid-wait.
  for (let attempt = 0; attempt < 2; attempt += 1) {
    try {
      await page.waitForFunction(() => Boolean((window.__formulaApp as any)?.whenIdle), null, { timeout: 10_000 });
      await page.evaluate(() => (window.__formulaApp as any).whenIdle());
      return;
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      if (attempt === 0 && message.includes("Execution context was destroyed")) {
        await page.waitForLoadState("domcontentloaded");
        continue;
      }
      throw err;
    }
  }
}

async function toggleShowFormulas(page: import("@playwright/test").Page): Promise<void> {
  const modifier = process.platform === "darwin" ? "Meta" : "Control";
  await page.keyboard.down(modifier);
  await page.keyboard.press("Backquote");
  await page.keyboard.up(modifier);
}

test.describe("show formulas", () => {
  test("renders computed values by default and toggles formula text via Ctrl/Cmd+`", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);

    // Seed A1=1 and A2=2.
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
    await expect(page.getByTestId("active-cell")).toHaveText("C1");

    await page.keyboard.press("F2");
    await cellEditor.fill("=SUM(A1:A2)");
    await page.keyboard.press("Enter");
    await waitForIdle(page);

    const computedValue = await page.evaluate(() => (window.__formulaApp as any).getCellDisplayValueA1("C1"));
    expect(computedValue).toBe("3");

    const defaultRenderText = await page.evaluate(() => (window.__formulaApp as any).getCellDisplayTextForRenderA1("C1"));
    expect(defaultRenderText).toBe("3");

    await toggleShowFormulas(page);
    const formulaRenderText = await page.evaluate(() => (window.__formulaApp as any).getCellDisplayTextForRenderA1("C1"));
    expect(formulaRenderText).toBe("=SUM(A1:A2)");

    await toggleShowFormulas(page);
    const toggledBackText = await page.evaluate(() => (window.__formulaApp as any).getCellDisplayTextForRenderA1("C1"));
    expect(toggledBackText).toBe("3");
  });

  test("selection renderer keeps drawing ranges when endpoints are offscreen", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);

    await page.evaluate(() => {
      const app = window.__formulaApp as any;
      app.selectRange({
        range: {
          startRow: 0,
          startCol: 0,
          endRow: 500,
          endCol: 100,
        },
      });
    });

    const drawn = await page.evaluate(() => (window.__formulaApp as any).getLastSelectionDrawn());
    expect(drawn).toBeTruthy();
    expect(drawn.ranges.length).toBeGreaterThan(0);
    expect(drawn.ranges[0].rect.width).toBeGreaterThan(0);
    expect(drawn.ranges[0].rect.height).toBeGreaterThan(0);
  });
});
