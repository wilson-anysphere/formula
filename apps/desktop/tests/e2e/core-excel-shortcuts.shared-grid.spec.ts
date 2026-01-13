import { expect, test, type Page } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function waitForIdle(page: Page): Promise<void> {
  await page.waitForFunction(() => Boolean((window as any).__formulaApp?.whenIdle), null, { timeout: 10_000 });
  await page.evaluate(() => (window as any).__formulaApp.whenIdle());
}

async function waitForCellRect(page: Page, a1: string): Promise<{ x: number; y: number; width: number; height: number }> {
  await page.waitForFunction(
    (cell) => {
      const app = (window as any).__formulaApp;
      const rect = app?.getCellRectA1?.(cell);
      return rect && typeof rect.x === "number" && rect.width > 0 && rect.height > 0;
    },
    a1,
    { timeout: 10_000 },
  );
  const rect = (await page.evaluate((cell) => (window as any).__formulaApp.getCellRectA1(cell), a1)) as {
    x: number;
    y: number;
    width: number;
    height: number;
  } | null;
  if (!rect) throw new Error(`Missing cell rect for ${a1}`);
  return rect;
}

async function clickCell(page: Page, a1: string): Promise<void> {
  const rect = await waitForCellRect(page, a1);
  await page.locator("#grid").click({ position: { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 } });
}

async function dispatchCtrlOrCmdKey(
  page: Page,
  key: string,
  opts: { ctrlKey?: boolean; metaKey?: boolean } = {},
): Promise<void> {
  await page.evaluate(
    ({ key, ctrlKey, metaKey }) => {
      const target = (document.activeElement as HTMLElement | null) ?? document.getElementById("grid") ?? window;
      target.dispatchEvent(new KeyboardEvent("keydown", { key, ctrlKey, metaKey, bubbles: true, cancelable: true }));
    },
    { key, ctrlKey: Boolean(opts.ctrlKey), metaKey: Boolean(opts.metaKey) },
  );
}

async function dispatchF6(page: Page, opts: { shiftKey?: boolean } = {}): Promise<void> {
  // Browsers can reserve F6 for built-in chrome focus cycling (address bar/toolbars),
  // which can prevent Playwright's `keyboard.press("F6")` from reaching the app.
  // Dispatching a synthetic `keydown` exercises our in-app focus cycling handler
  // deterministically.
  await page.evaluate(
    ({ shiftKey }) => {
      const target = (document.activeElement as HTMLElement | null) ?? document.getElementById("grid") ?? window;
      target.dispatchEvent(
        new KeyboardEvent("keydown", {
          key: "F6",
          code: "F6",
          shiftKey: Boolean(shiftKey),
          bubbles: true,
          cancelable: true,
        }),
      );
    },
    { shiftKey: opts.shiftKey ?? false },
  );
}

test("shared grid: core Excel shortcuts + interactions smoke", async ({ page }) => {
  await gotoDesktop(page, "/?grid=shared");

  await test.step("Edit A1 via typing + Enter moves to A2", async () => {
    await clickCell(page, "A1");
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    await page.keyboard.type("123");
    await page.keyboard.press("Enter");
    await waitForIdle(page);

    await expect(page.getByTestId("active-cell")).toHaveText("A2");
    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"))).toBe("123");
  });

  await test.step("F2 enters edit mode in the active cell", async () => {
    await expect(page.getByTestId("active-cell")).toHaveText("A2");
    await page.keyboard.press("F2");

    const editor = page.getByTestId("cell-editor");
    await expect(editor).toBeVisible();
    await expect(editor).toBeFocused();

    await editor.fill("456");
    await page.keyboard.press("Enter");
    await waitForIdle(page);

    await expect(page.getByTestId("active-cell")).toHaveText("A3");
    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A2"))).toBe("456");
  });

  await test.step("Ctrl/Cmd+D fills down", async () => {
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const sheetId = app.getCurrentSheetId();
      const doc = app.getDocument();

      doc.setCellValue(sheetId, "D1", 1);
      doc.setCellValue(sheetId, "E1", 2);

      // Select D1:E3.
      app.selectRange({ range: { startRow: 0, endRow: 2, startCol: 3, endCol: 4 } });
      app.focus();
    });
    await waitForIdle(page);

    const modifier = process.platform === "darwin" ? "Meta" : "Control";
    await page.keyboard.press(`${modifier}+D`);
    await waitForIdle(page);

    const [d2, e2, d3, e3] = await Promise.all([
      page.evaluate(() => (window as any).__formulaApp.getCellValueA1("D2")),
      page.evaluate(() => (window as any).__formulaApp.getCellValueA1("E2")),
      page.evaluate(() => (window as any).__formulaApp.getCellValueA1("D3")),
      page.evaluate(() => (window as any).__formulaApp.getCellValueA1("E3")),
    ]);

    expect(d2).toBe("1");
    expect(e2).toBe("2");
    expect(d3).toBe("1");
    expect(e3).toBe("2");
  });

  await test.step("F4 toggles absolute refs while editing in the formula bar", async () => {
    await page.evaluate(() => (window as any).__formulaApp.activateCell({ row: 0, col: 2 })); // C1
    await waitForIdle(page);

    await page.getByTestId("formula-highlight").click();
    const input = page.getByTestId("formula-input");
    await expect(input).toBeVisible();

    await input.fill("=SUM(A1)");

    // Place caret inside the A1 token (just before the closing paren).
    await input.focus();
    await page.keyboard.press("ArrowLeft");

    await page.keyboard.press("F4");
    await expect(input).toHaveValue("=SUM($A$1)");

    // Commit the edit (restores grid focus for subsequent steps).
    await page.keyboard.press("Enter");
    await waitForIdle(page);
  });

  await test.step("Ctrl/Cmd+PgUp/PgDn switches visible sheets", async () => {
    // Lazily create Sheet2 so sheet navigation is observable.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.getDocument().setCellValue("Sheet2", "A1", "Hello from Sheet2");
    });
    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();

    const isMac = process.platform === "darwin";
    await page.evaluate(() => (window as any).__formulaApp.focus());
    await expect(page.locator("#grid")).toBeFocused();

    await dispatchCtrlOrCmdKey(page, "PageDown", { metaKey: isMac, ctrlKey: !isMac });
    await expect(page.getByTestId("sheet-tab-Sheet2")).toHaveAttribute("data-active", "true");

    await dispatchCtrlOrCmdKey(page, "PageUp", { metaKey: isMac, ctrlKey: !isMac });
    await expect(page.getByTestId("sheet-tab-Sheet1")).toHaveAttribute("data-active", "true");
  });

  await test.step("F6 cycles focus to the ribbon and back to the grid", async () => {
    const ribbonRoot = page.getByTestId("ribbon-root");
    await expect(ribbonRoot).toBeVisible();
    const activeRibbonTab = ribbonRoot.locator('[role="tab"][aria-selected="true"]');

    await page.evaluate(() => (window as any).__formulaApp.focus());
    await expect(page.locator("#grid")).toBeFocused();

    // grid -> sheet tabs -> status bar -> ribbon
    await dispatchF6(page);
    await dispatchF6(page);
    await dispatchF6(page);
    await expect(activeRibbonTab).toBeFocused();

    // ribbon -> formula bar -> grid
    await dispatchF6(page);
    await dispatchF6(page);
    await expect(page.locator("#grid")).toBeFocused();
  });
});

