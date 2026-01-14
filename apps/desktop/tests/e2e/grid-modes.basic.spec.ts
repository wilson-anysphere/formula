import { expect, test, type Page } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function whenIdle(page: Page): Promise<void> {
  // Vite may occasionally trigger a one-time full reload after dependency optimization.
  // Retry once if the execution context is destroyed mid-wait.
  for (let attempt = 0; attempt < 2; attempt += 1) {
    try {
      await page.evaluate(async () => {
        const app = (window as any).__formulaApp;
        if (app && typeof app.whenIdle === "function") {
          await app.whenIdle();
        }
      });
      return;
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      if (
        attempt === 0 &&
        (message.includes("Execution context was destroyed") ||
          message.includes("net::ERR_ABORTED") ||
          message.includes("net::ERR_NETWORK_CHANGED") ||
          message.includes("frame was detached"))
      ) {
        await page.waitForLoadState("domcontentloaded").catch(() => {});
        continue;
      }
      throw err;
    }
  }
}

const GRID_MODES = ["legacy", "shared"] as const;

for (const mode of GRID_MODES) {
  test.describe(`${mode} grid basics`, () => {
    test("keyboard navigation, edit, and scroll", async ({ page }) => {
      await gotoDesktop(page, `/?grid=${mode}`);

      // Focus grid.
      await page.click("#grid", { position: { x: 60, y: 40 } });
      await expect(page.getByTestId("active-cell")).toHaveText("A1");

      await page.keyboard.press("ArrowRight");
      await page.keyboard.press("ArrowRight");
      await page.keyboard.press("ArrowDown");
      await expect(page.getByTestId("active-cell")).toHaveText("C2");

      // Start typing edits the active cell.
      await page.keyboard.press("h");
      const editor = page.locator("textarea.cell-editor");
      await expect(editor).toBeVisible();
      await page.keyboard.type("ello");
      await page.keyboard.press("Enter");
      await whenIdle(page);

      const c2Value = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("C2"));
      expect(c2Value).toBe("hello");
      await expect(page.getByTestId("active-cell")).toHaveText("C3");

      // Wheel scrolling should update the reported scroll position.
      const scrollBefore = await page.evaluate(() => (window as any).__formulaApp.getScroll().y);
      const gridBox = await page.locator("#grid").boundingBox();
      if (!gridBox) throw new Error("Missing grid bounding box");
      await page.mouse.move(gridBox.x + gridBox.width / 2, gridBox.y + gridBox.height / 2);
      await page.mouse.wheel(0, 1600);
      const scrollAfter = await page.evaluate(() => (window as any).__formulaApp.getScroll().y);
      expect(scrollAfter).toBeGreaterThan(scrollBefore);
    });
  });
}

test.describe("shared grid resize", () => {
  test("column resize updates layout", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");

    const gridBox = await page.locator("#grid").boundingBox();
    if (!gridBox) throw new Error("Missing grid bounding box");

    const before = await page.evaluate(() => (window as any).__formulaApp.getCellRectA1("B1"));
    if (!before) throw new Error("Missing B1 rect");

    // Drag the boundary between columns A and B in the header row to make column A wider.
    const boundaryX = before.x;
    const boundaryY = before.y / 2;

    await page.mouse.move(gridBox.x + boundaryX, gridBox.y + boundaryY);
    await page.mouse.down();
    await page.mouse.move(gridBox.x + boundaryX + 80, gridBox.y + boundaryY, { steps: 4 });
    await page.mouse.up();

    await page.waitForFunction(
      (threshold) => {
        const rect = (window as any).__formulaApp.getCellRectA1("B1");
        return rect && rect.x > threshold;
      },
      before.x + 30
    );

    const after = await page.evaluate(() => (window as any).__formulaApp.getCellRectA1("B1"));
    if (!after) throw new Error("Missing B1 rect after resize");
    expect(after.x).toBeGreaterThan(before.x + 30);
  });
});
