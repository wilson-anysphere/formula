import { expect, test } from "@playwright/test";

async function waitForIdle(page: import("@playwright/test").Page): Promise<void> {
  await page.evaluate(() => (window as any).__formulaApp.whenIdle());
}

test.describe("clipboard shortcuts (copy/cut/paste)", () => {
  test("Ctrl/Cmd+C copies selection and Ctrl/Cmd+V pastes starting at active cell", async ({ page }) => {
    await page.context().grantPermissions(["clipboard-read", "clipboard-write"]);
    await page.goto("/");

    const modifier = process.platform === "darwin" ? "Meta" : "Control";

    // Seed A1 = Hello, A2 = World.
    await page.click("#grid", { position: { x: 53, y: 29 } });
    await page.keyboard.press("F2");
    const editor = page.locator("textarea.cell-editor");
    await expect(editor).toBeVisible();
    await editor.fill("Hello");
    await page.keyboard.press("Enter");
    await waitForIdle(page);

    await page.keyboard.press("F2");
    await expect(editor).toBeVisible();
    await editor.fill("World");
    await page.keyboard.press("Enter");
    await waitForIdle(page);

    // Select A1:A2 via drag.
    const gridBox = await page.locator("#grid").boundingBox();
    if (!gridBox) throw new Error("Missing grid bounding box");
    await page.mouse.move(gridBox.x + 60, gridBox.y + 40); // A1
    await page.mouse.down();
    await page.mouse.move(gridBox.x + 60, gridBox.y + 64); // A2
    await page.mouse.up();

    await page.keyboard.press(`${modifier}+C`);
    await waitForIdle(page);

    // Paste into C1.
    await page.click("#grid", { position: { x: 260, y: 40 } });
    await page.keyboard.press(`${modifier}+V`);
    await waitForIdle(page);

    // Paste updates the selection to match the pasted dimensions.
    await expect(page.getByTestId("selection-range")).toHaveText("C1:C2");

    const c1Value = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("C1"));
    expect(c1Value).toBe("Hello");
    const c2Value = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("C2"));
    expect(c2Value).toBe("World");

    // Cut A1 and paste to B1.
    await page.click("#grid", { position: { x: 53, y: 29 } });
    await page.keyboard.press(`${modifier}+X`);
    await waitForIdle(page);

    await expect
      .poll(() => page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1")))
      .toBe("");

    await page.click("#grid", { position: { x: 160, y: 40 } });
    await page.keyboard.press(`${modifier}+V`);
    await waitForIdle(page);

    await expect
      .poll(() => page.evaluate(() => (window as any).__formulaApp.getCellValueA1("B1")))
      .toBe("Hello");
  });
});
