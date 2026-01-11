import { expect, test } from "@playwright/test";

test.describe("formula reference colors", () => {
  test("colors each reference in the formula bar and renders matching grid overlays", async ({ page }) => {
    await page.goto("/");

    // Select C1 (avoid overlapping the referenced cells).
    await page.click("#grid", { position: { x: 260, y: 40 } });
    await expect(page.getByTestId("active-cell")).toHaveText("C1");

    await page.getByTestId("formula-highlight").click();
    const input = page.getByTestId("formula-input");
    await expect(input).toBeVisible();

    await input.fill("=A1+B1");

    const refs = page.locator('[data-testid="formula-highlight"] span[data-kind="reference"]');
    await expect(refs).toHaveCount(2);

    const colors = await refs.evaluateAll((els) => els.map((el) => getComputedStyle(el).color));
    expect(colors[0]).not.toBe(colors[1]);

    await page.waitForFunction(() => (window as any).__formulaApp.getReferenceHighlightCount() === 2);
  });
});

