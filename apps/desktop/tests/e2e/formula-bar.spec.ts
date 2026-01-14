import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function waitForIdle(page: import("@playwright/test").Page): Promise<void> {
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
      throw err;
    }
  }
}

async function dragInLocator(
  page: import("@playwright/test").Page,
  locator: import("@playwright/test").Locator,
  from: { x: number; y: number },
  to: { x: number; y: number },
): Promise<void> {
  // Locator-relative coordinates are resilient to layout shifts (e.g. formula bar resizing while
  // entering edit / range selection modes).
  await locator.hover({ position: from });
  await page.mouse.down();
  await locator.hover({ position: to });
  await page.mouse.up();
}

test.describe("formula bar editing + range insertion", () => {
  const modes = ["legacy", "shared"] as const;

  for (const mode of modes) {
    test(`type formula, drag range, commit stores formula in the active cell (${mode})`, async ({ page }) => {
      await gotoDesktop(page, `/?grid=${mode}`);
      await waitForIdle(page);

      // Seed numeric inputs in A1 and A2 (so SUM has a visible result).
      // Click within the first grid cell (accounting for row/column headers).
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

      // Select C1.
      // Account for the row/column headers rendered inside the grid canvas.
      await page.click("#grid", { position: { x: 260, y: 40 } });
      await expect(page.getByTestId("active-cell")).toHaveText("C1");

      // Start editing in the formula bar.
      await page.getByTestId("formula-highlight").click();
      const input = page.getByTestId("formula-input");
      await expect(input).toBeVisible();
      await input.fill("=SUM(");

      // Drag select A1:A2 to insert a range reference.
      if (mode === "shared") {
        await expect
          .poll(() => page.evaluate(() => (window as any).__formulaApp?.sharedGrid?.interactionMode ?? null))
          .toBe("rangeSelection");
      }
      const grid = page.locator("#grid");
      await dragInLocator(page, grid, { x: 60, y: 40 }, { x: 60, y: 64 });

      await expect(input).toHaveValue("=SUM(A1:A2");

      await input.focus();
      await page.keyboard.type(")");
      await page.keyboard.press("Enter");
      await waitForIdle(page);

      const { a1Value, a2Value, c1Formula } = await page.evaluate(() => {
        const app = (window as any).__formulaApp;
        const doc = app.getDocument();
        return {
          a1Value: doc.getCell("Sheet1", "A1").value,
          a2Value: doc.getCell("Sheet1", "A2").value,
          c1Formula: doc.getCell("Sheet1", "C1").formula,
        };
      });
      expect(a1Value).toBe(1);
      expect(a2Value).toBe(2);
      expect(c1Formula).toBe("=SUM(A1:A2)");

      const c1Value = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("C1"));
      expect(c1Value).toBe("3");
    });

    test(`picking a range on another sheet inserts a sheet-qualified reference and commits to the original edit cell (${mode})`, async ({
      page,
    }) => {
      await gotoDesktop(page, `/?grid=${mode}`);
      await waitForIdle(page);

      // Lazily create Sheet2 and seed A1.
      await page.evaluate(() => {
        const app = (window as any).__formulaApp;
        app.getDocument().setCellValue("Sheet2", "A1", 7);
      });
      await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();

      // Start editing on Sheet1!C1.
      await page.click("#grid", { position: { x: 260, y: 40 } });
      await expect(page.getByTestId("active-cell")).toHaveText("C1");

      await page.getByTestId("formula-highlight").click();
      const input = page.getByTestId("formula-input");
      await expect(input).toBeVisible();
      await input.fill("=");

      // Switch to Sheet2 while still editing and pick A1.
      await page.getByTestId("sheet-tab-Sheet2").click();
      await expect(page.getByTestId("sheet-tab-Sheet2")).toHaveAttribute("data-active", "true");
      // Sheet switching during formula editing should not steal focus away from the formula bar.
      await expect(input).toBeFocused();
      await page.click("#grid", { position: { x: 53, y: 29 } });
      await expect(input).toHaveValue("=Sheet2!A1");

      // Commit should apply to the original edit cell (Sheet1!C1) and restore the sheet.
      await input.focus();
      await page.keyboard.press("Enter");
      await waitForIdle(page);

      await expect(page.getByTestId("sheet-tab-Sheet1")).toHaveAttribute("data-active", "true");
      // Formula-bar commit (Enter) advances selection to the next cell (Excel-like).
      await expect(page.getByTestId("active-cell")).toHaveText("C2");

      const { sheet1Formula, sheet2Formula, sheet2Value } = await page.evaluate(() => {
        const app = (window as any).__formulaApp;
        const doc = app.getDocument();
        return {
          sheet1Formula: doc.getCell("Sheet1", "C1").formula,
          sheet2Formula: doc.getCell("Sheet2", "A1").formula,
          sheet2Value: doc.getCell("Sheet2", "A1").value,
        };
      });

      expect(sheet1Formula).toBe("=Sheet2!A1");
      expect(sheet2Formula).toBeNull();
      expect(sheet2Value).toBe(7);

      const c1Value = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("C1"));
      expect(c1Value).toBe("7");
    });

    test(`Ctrl/Cmd+PgDn switches sheets while editing a formula (formula bar focused) (${mode})`, async ({
      page,
    }) => {
      await gotoDesktop(page, `/?grid=${mode}`);
      await waitForIdle(page);

      // Lazily create Sheet2 and seed A1.
      await page.evaluate(() => {
        const app = (window as any).__formulaApp;
        app.getDocument().setCellValue("Sheet2", "A1", 7);
      });
      await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();

      // Start editing on Sheet1!C1.
      await page.click("#grid", { position: { x: 260, y: 40 } });
      await expect(page.getByTestId("active-cell")).toHaveText("C1");

      await page.getByTestId("formula-highlight").click();
      const input = page.getByTestId("formula-input");
      await expect(input).toBeVisible();
      await input.fill("=");

      // Wait for the app to recognize this as a formula-editing session so sheet navigation is allowed
      // (Excel behavior for cross-sheet reference building).
      await expect
        .poll(() =>
          page.evaluate(() => {
            const app = (window as any).__formulaApp;
            return Boolean(app?.isFormulaBarFormulaEditing?.());
          }),
        )
        .toBe(true);

      // Ctrl/Cmd+PgDn should switch sheets while editing a formula in the formula bar.
      // Dispatch synthetically to avoid any browser tab-switching shortcuts.
      const ctrlKey = process.platform !== "darwin";
      const metaKey = process.platform === "darwin";
      await page.evaluate(
        ({ ctrlKey, metaKey }) => {
          const input = document.querySelector('[data-testid="formula-input"]') as HTMLTextAreaElement | null;
          if (!input) throw new Error("Missing formula bar input");
          const evt = new KeyboardEvent("keydown", { key: "PageDown", ctrlKey, metaKey, bubbles: true, cancelable: true });
          input.dispatchEvent(evt);
        },
        { ctrlKey, metaKey },
      );
      await expect(page.getByTestId("sheet-tab-Sheet2")).toHaveAttribute("data-active", "true");
      await expect(input).toBeFocused();

      // Picking a cell should insert a sheet-qualified reference (because the reference is on a different sheet).
      // Click inside A1 (avoid the shared-grid corner header/select-all region).
      await page.click("#grid", { position: { x: 80, y: 40 } });
      await expect(input).toHaveValue("=Sheet2!A1");
    });

    test(`sheet switcher <select> switches sheets while editing a formula and keeps the formula bar focused (${mode})`, async ({
      page,
    }) => {
      await gotoDesktop(page, `/?grid=${mode}`);
      await waitForIdle(page);

      // Lazily create Sheet2 and seed A1.
      await page.evaluate(() => {
        const app = (window as any).__formulaApp;
        app.getDocument().setCellValue("Sheet2", "A1", 7);
      });
      await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();

      // Start editing on Sheet1!C1.
      await page.click("#grid", { position: { x: 260, y: 40 } });
      await expect(page.getByTestId("active-cell")).toHaveText("C1");

      await page.getByTestId("formula-highlight").click();
      const input = page.getByTestId("formula-input");
      await expect(input).toBeVisible();
      await input.fill("=");

      const switcher = page.getByTestId("sheet-switcher");
      await switcher.selectOption("Sheet2", { force: true });

      await expect(page.getByTestId("sheet-tab-Sheet2")).toHaveAttribute("data-active", "true");
      await expect(input).toBeFocused();

      await page.click("#grid", { position: { x: 53, y: 29 } });
      await expect(input).toHaveValue("=Sheet2!A1");
    });

    test(`sheet overflow quick pick switches sheets while editing a formula and keeps the formula bar focused (${mode})`, async ({
      page,
    }) => {
      await gotoDesktop(page, `/?grid=${mode}`);
      await waitForIdle(page);

      // Lazily create Sheet2 and seed A1.
      await page.evaluate(() => {
        const app = (window as any).__formulaApp;
        app.getDocument().setCellValue("Sheet2", "A1", 7);
      });
      await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();

      // Start editing on Sheet1!C1.
      await page.click("#grid", { position: { x: 260, y: 40 } });
      await expect(page.getByTestId("active-cell")).toHaveText("C1");

      await page.getByTestId("formula-highlight").click();
      const input = page.getByTestId("formula-input");
      await expect(input).toBeVisible();
      await input.fill("=");

      await page.getByTestId("sheet-overflow").click();
      const quickPick = page.getByTestId("quick-pick");
      await expect(quickPick).toBeVisible();
      await quickPick.getByRole("button", { name: "Sheet2" }).click();

      await expect(page.getByTestId("sheet-tab-Sheet2")).toHaveAttribute("data-active", "true");
      await expect(input).toBeFocused();

      await page.click("#grid", { position: { x: 53, y: 29 } });
      await expect(input).toHaveValue("=Sheet2!A1");
    });

    test(`canceling after switching sheets restores the original edit location without applying edits (${mode})`, async ({ page }) => {
      await gotoDesktop(page, `/?grid=${mode}`);
      await waitForIdle(page);

      // Lazily create Sheet2 and seed A1.
      await page.evaluate(() => {
        const app = (window as any).__formulaApp;
        app.getDocument().setCellValue("Sheet2", "A1", 7);
      });
      await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();

      // Start editing on Sheet1!C1.
      await page.click("#grid", { position: { x: 260, y: 40 } });
      await expect(page.getByTestId("active-cell")).toHaveText("C1");

      await page.getByTestId("formula-highlight").click();
      const input = page.getByTestId("formula-input");
      await expect(input).toBeVisible();
      await input.fill("=");

      // Switch to Sheet2 and pick A1 to insert a reference.
      await page.getByTestId("sheet-tab-Sheet2").click();
      await expect(page.getByTestId("sheet-tab-Sheet2")).toHaveAttribute("data-active", "true");
      // Sheet switching during formula editing should not steal focus away from the formula bar.
      await expect(input).toBeFocused();
      await page.click("#grid", { position: { x: 53, y: 29 } });
      await expect(input).toHaveValue("=Sheet2!A1");

      // Cancel should restore Sheet1!C1 and leave the cell unchanged.
      await input.focus();
      await page.keyboard.press("Escape");
      await waitForIdle(page);

      await expect(page.getByTestId("sheet-tab-Sheet1")).toHaveAttribute("data-active", "true");
      await expect(page.getByTestId("active-cell")).toHaveText("C1");

      const sheet1Formula = await page.evaluate(() => (window as any).__formulaApp.getDocument().getCell("Sheet1", "C1").formula);
      expect(sheet1Formula).toBeNull();
    });

    test(`shows friendly error explanation for #DIV/0! (${mode})`, async ({ page }) => {
      await gotoDesktop(page, `/?grid=${mode}`);
      await waitForIdle(page);

      // Seed A1 = 0.
      await page.click("#grid", { position: { x: 53, y: 29 } });
      await page.keyboard.press("F2");
      const cellEditor = page.locator("textarea.cell-editor");
      await cellEditor.fill("0");
      await page.keyboard.press("Enter");
      await waitForIdle(page);

      // Select B1.
      await page.click("#grid", { position: { x: 160, y: 40 } });
      await expect(page.getByTestId("active-cell")).toHaveText("B1");

      // Enter a division-by-zero formula.
      await page.getByTestId("formula-highlight").click();
      const input = page.getByTestId("formula-input");
      await expect(input).toBeVisible();
      await input.fill("=1/A1");
      await page.keyboard.press("Enter");
      await waitForIdle(page);

      // Formula-bar commit advances selection, so re-select the formula cell before asserting UI state.
      await page.click("#grid", { position: { x: 160, y: 40 } });
      await expect(page.getByTestId("active-cell")).toHaveText("B1");

      // Error button should appear and panel should explain.
      const errorButton = page.getByTestId("formula-error-button");
      await expect(errorButton).toBeVisible();
      await errorButton.click();
      await expect(page.getByTestId("formula-error-panel")).toBeVisible();
      await expect(page.getByTestId("formula-error-panel")).toContainText("Division by zero");
      await expect(page.getByTestId("formula-error-panel").locator(".formula-bar-error-title")).toContainText("B1");
    });
  }
});
