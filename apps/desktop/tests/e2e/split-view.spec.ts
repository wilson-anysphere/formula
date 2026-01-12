import { expect, test } from "@playwright/test";

test.describe("split view / shared grid zoom", () => {
  test("Ctrl/Cmd+wheel zoom changes grid geometry", async ({ page }) => {
    await page.goto("/?grid=shared");

    const grid = page.locator("#grid");

    await page.waitForFunction(() => {
      const app = (window as any).__formulaApp;
      const rect = app?.getCellRectA1?.("B1");
      return rect && typeof rect.x === "number" && rect.width > 0 && rect.height > 0;
    });

    const rectsBefore = await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      return {
        a1: app.getCellRectA1("A1"),
        b1: app.getCellRectA1("B1"),
      };
    });

    expect(rectsBefore.a1).toBeTruthy();
    expect(rectsBefore.b1).toBeTruthy();

    const a1Before = rectsBefore.a1 as { x: number; y: number; width: number; height: number };
    const b1Before = rectsBefore.b1 as { x: number; y: number; width: number; height: number };

    // Hover inside the row header so the zoom gesture doesn't anchor to the scrollable quadrant.
    await grid.hover({
      position: { x: Math.max(1, a1Before.x / 2), y: a1Before.y + a1Before.height / 2 },
    });

    const primary = process.platform === "darwin" ? "Meta" : "Control";
    await page.keyboard.down(primary);
    await page.mouse.wheel(0, -100);
    await page.keyboard.up(primary);

    await expect
      .poll(async () => {
        const rect = await page.evaluate(() => (window as any).__formulaApp.getCellRectA1("B1"));
        return rect?.x ?? null;
      })
      .toBeGreaterThan(b1Before.x);
  });
});

