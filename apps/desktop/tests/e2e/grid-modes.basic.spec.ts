import { expect, test } from "@playwright/test";

function captureAppErrors(page: import("@playwright/test").Page): string[] {
  const errors: string[] = [];

  page.on("console", (msg) => {
    if (msg.type() !== "error") return;
    errors.push(msg.text());
  });

  page.on("pageerror", (err) => {
    errors.push(err.message ?? String(err));
  });

  return errors;
}

async function waitForIdle(page: import("@playwright/test").Page, capturedErrors: string[]): Promise<void> {
  // Vite may occasionally trigger a one-time full reload after dependency optimization.
  // Retry once if the execution context is destroyed mid-wait.
  for (let attempt = 0; attempt < 2; attempt += 1) {
    try {
      await page.waitForFunction(() => Boolean((window as any).__formulaApp?.whenIdle), null, { timeout: 10_000 });
      await page.evaluate(() => (window as any).__formulaApp.whenIdle());
      return;
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      if (attempt === 0 && message.includes("Execution context was destroyed")) {
        await page.waitForLoadState("domcontentloaded");
        continue;
      }
      const errorText = capturedErrors.length > 0 ? capturedErrors.join("\n") : "(no console errors captured)";
      throw new Error(`Spreadsheet app failed to initialize.\n\nCaptured browser errors:\n${errorText}\n\nOriginal error:\n${String(err)}`);
    }
  }
}

const GRID_MODES = ["legacy", "shared"] as const;

  for (const mode of GRID_MODES) {
    test.describe(`${mode} grid basics`, () => {
      test("keyboard navigation, edit, and scroll", async ({ page }) => {
        const capturedErrors = captureAppErrors(page);
        await page.goto(`/?grid=${mode}`, { waitUntil: "domcontentloaded" });
        await waitForIdle(page, capturedErrors);

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
      await waitForIdle(page, capturedErrors);

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
    const capturedErrors = captureAppErrors(page);
    await page.goto("/?grid=shared", { waitUntil: "domcontentloaded" });
    await waitForIdle(page, capturedErrors);

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
