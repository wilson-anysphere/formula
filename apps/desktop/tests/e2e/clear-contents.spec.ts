import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function waitForIdle(page: import("@playwright/test").Page): Promise<void> {
  await page.waitForFunction(() => Boolean((window.__formulaApp as any)?.whenIdle), null, { timeout: 10_000 });
  await page.evaluate(() => (window.__formulaApp as any).whenIdle());
}

test.describe("Clear Contents context menu", () => {
  test("clears a single cell via the cell context menu", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);

    const styleIdBefore = await page.evaluate(() => {
      const app = window.__formulaApp as any;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      doc.setCellValue(sheetId, { row: 0, col: 0 }, "Hello");
      doc.setRangeFormat(sheetId, "A1", { font: { bold: true } });
      return (doc.getCell(sheetId, { row: 0, col: 0 }) as any)?.styleId ?? 0;
    });
    await waitForIdle(page);
    expect(styleIdBefore).not.toBe(0);

    const grid = page.locator("#grid");
    const a1Rect = await page.evaluate(() => (window.__formulaApp as any).getCellRectA1("A1"));
    await grid.click({ position: { x: a1Rect.x + a1Rect.width / 2, y: a1Rect.y + a1Rect.height / 2 } });
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    // Avoid flaky right-click handling in the desktop shell; dispatch a deterministic contextmenu event.
    await page.evaluate(
      ({ x, y }) => {
        const grid = document.getElementById("grid");
        if (!grid) throw new Error("Missing #grid container");
        const rect = grid.getBoundingClientRect();
        grid.dispatchEvent(
          new MouseEvent("contextmenu", {
            bubbles: true,
            cancelable: true,
            button: 2,
            clientX: rect.left + x,
            clientY: rect.top + y,
          }),
        );
      },
      { x: a1Rect.x + a1Rect.width / 2, y: a1Rect.y + a1Rect.height / 2 },
    );
    const menu = page.getByTestId("context-menu");
    await expect(menu).toBeVisible();

    const item = menu.getByRole("button", { name: "Clear Contents" });
    await expect(item).toBeEnabled();
    await item.click();

    await waitForIdle(page);
    const { value, styleIdAfter } = await page.evaluate(async () => {
      const app = window.__formulaApp as any;
      const sheetId = app.getCurrentSheetId();
      const doc = app.getDocument();
      return {
        value: await app.getCellValueA1("A1"),
        styleIdAfter: (doc.getCell(sheetId, { row: 0, col: 0 }) as any)?.styleId ?? 0,
      };
    });
    expect(value).toBe("");
    expect(styleIdAfter).toBe(styleIdBefore);
  });

  test("clears a single cell via the command palette", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);

    await page.evaluate(() => {
      const app = window.__formulaApp as any;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      doc.setCellValue(sheetId, { row: 0, col: 0 }, "Hello");
    });
    await waitForIdle(page);

    const grid = page.locator("#grid");
    const a1Rect = await page.evaluate(() => (window.__formulaApp as any).getCellRectA1("A1"));
    await grid.click({ position: { x: a1Rect.x + a1Rect.width / 2, y: a1Rect.y + a1Rect.height / 2 } });
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    const primary = process.platform === "darwin" ? "Meta" : "Control";
    await page.keyboard.press(`${primary}+Shift+P`);

    const input = page.getByTestId("command-palette-input");
    await expect(input).toBeVisible();
    await input.fill("Clear Contents");
    await page.keyboard.press("Enter");

    await waitForIdle(page);
    const value = await page.evaluate(() => (window.__formulaApp as any).getCellValueA1("A1"));
    expect(value).toBe("");
  });

  test("disables Clear Contents when the selection is empty", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);

    // Ensure A1 is empty.
    await page.evaluate(() => {
      const app = window.__formulaApp as any;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      doc.setCellValue(sheetId, { row: 0, col: 0 }, null);
    });
    await waitForIdle(page);

    const grid = page.locator("#grid");
    const a1Rect = await page.evaluate(() => (window.__formulaApp as any).getCellRectA1("A1"));
    await grid.click({ position: { x: a1Rect.x + a1Rect.width / 2, y: a1Rect.y + a1Rect.height / 2 } });
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    // Avoid flaky right-click handling in the desktop shell; dispatch a deterministic contextmenu event.
    await page.evaluate(
      ({ x, y }) => {
        const grid = document.getElementById("grid");
        if (!grid) throw new Error("Missing #grid container");
        const rect = grid.getBoundingClientRect();
        grid.dispatchEvent(
          new MouseEvent("contextmenu", {
            bubbles: true,
            cancelable: true,
            button: 2,
            clientX: rect.left + x,
            clientY: rect.top + y,
          }),
        );
      },
      { x: a1Rect.x + a1Rect.width / 2, y: a1Rect.y + a1Rect.height / 2 },
    );
    const menu = page.getByTestId("context-menu");
    await expect(menu).toBeVisible();

    const item = menu.getByRole("button", { name: "Clear Contents" });
    await expect(item).toBeDisabled();
  });
});
