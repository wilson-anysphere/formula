import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("sheet tabs", () => {
  test("switching sheets updates the visible cell values", async ({ page }) => {
    await gotoDesktop(page);

    // Ensure A1 is active before switching sheets so the status bar reflects A1 values.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.activateCell({ row: 0, col: 0 });
    });
    await expect(page.getByTestId("sheet-tab-Sheet1")).toBeVisible();
    await expect(page.getByTestId("sheet-position")).toHaveText("Sheet 1 of 1");

    // Lazily create Sheet2 by writing a value into it.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.getDocument().setCellValue("Sheet2", "A1", "Hello from Sheet2");
    });

    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();
    await expect(page.getByTestId("sheet-position")).toHaveText("Sheet 1 of 2");

    await page.getByTestId("sheet-tab-Sheet2").click();
    await expect(page.getByTestId("active-cell")).toHaveText("A1");
    await expect(page.getByTestId("active-value")).toHaveText("Hello from Sheet2");
    await expect(page.getByTestId("sheet-position")).toHaveText("Sheet 2 of 2");

    // Switching back restores the original Sheet1 value.
    await page.getByTestId("sheet-tab-Sheet1").click();
    await expect(page.getByTestId("active-value")).toHaveText("Seed");
    await expect(page.getByTestId("sheet-position")).toHaveText("Sheet 1 of 2");
  });

  test("add sheet button creates and activates the next SheetN tab", async ({ page }) => {
    await gotoDesktop(page);

    // Ensure A1 is active so the status bar is deterministic after the sheet switch.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.activateCell({ row: 0, col: 0 });
    });

    const nextSheetId = await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const ids = app.getDocument().getSheetIds();
      const existing = new Set((ids.length > 0 ? ids : ["Sheet1"]) as string[]);
      let n = 1;
      while (existing.has(`Sheet${n}`)) n += 1;
      return `Sheet${n}`;
    });

    await page.getByTestId("sheet-add").click();

    const newTab = page.getByTestId(`sheet-tab-${nextSheetId}`);
    await expect(newTab).toBeVisible();
    await expect(newTab).toHaveAttribute("data-active", "true");

    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCurrentSheetId())).toBe(nextSheetId);

    // Sheet activation should return focus to the grid so keyboard shortcuts keep working.
    await expect
      .poll(() => page.evaluate(() => (document.activeElement as HTMLElement | null)?.id))
      .toBe("grid");

    // Verify the new sheet is functional by writing a value into A1 and observing the status bar update.
    await page.evaluate((sheetId) => {
      const app = (window as any).__formulaApp;
      app.getDocument().setCellValue(sheetId, "A1", `Hello from ${sheetId}`);
    }, nextSheetId);

    await expect(page.getByTestId("active-cell")).toHaveText("A1");
    await expect(page.getByTestId("active-value")).toHaveText(`Hello from ${nextSheetId}`);
  });

  test("add sheet marks the document dirty and undo returns to clean state", async ({ page }) => {
    await gotoDesktop(page);

    const initialSheetIds = await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.getDocument().markSaved();
      return app.getWorkbookSheetStore().listAll().map((s: any) => s.id);
    });
    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getDocument().isDirty)).toBe(false);

    await page.getByTestId("sheet-add").click();

    // Adding a sheet should create an undo step and mark the workbook dirty.
    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getDocument().isDirty)).toBe(true);
    await expect
      .poll(() => page.evaluate(() => (window as any).__formulaApp.getWorkbookSheetStore().listAll().length))
      .toBe(initialSheetIds.length + 1);

    // Undo should remove the sheet and return to the last-saved dirty state.
    //
    // Use `app.undo()` (not `doc.undo()`) so SpreadsheetApp can fall back to a valid active
    // sheet id. Otherwise the renderer can re-materialize the deleted sheet via
    // DocumentController's lazy sheet creation.
    await page.evaluate(async () => {
      const app = (window as any).__formulaApp;
      app.undo();
      await app.whenIdle();
    });

    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getDocument().isDirty)).toBe(false);
    await expect
      .poll(() => page.evaluate(() => (window as any).__formulaApp.getWorkbookSheetStore().listAll().map((s: any) => s.id)))
      .toEqual(initialSheetIds);
  });

  test("sheet overflow menu activates the selected sheet", async ({ page }) => {
    await gotoDesktop(page);

    // Create Sheet2..Sheet5 via the UI.
    for (const sheetId of ["Sheet2", "Sheet3", "Sheet4", "Sheet5"]) {
      await page.getByTestId("sheet-add").click();
      await expect(page.getByTestId(`sheet-tab-${sheetId}`)).toBeVisible();
    }

    await expect(page.getByTestId("sheet-position")).toHaveText("Sheet 5 of 5");

    await page.getByTestId("sheet-overflow").click();

    const quickPick = page.getByTestId("quick-pick");
    await expect(quickPick).toBeVisible();
    await quickPick.getByRole("button", { name: "Sheet4" }).click();

    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCurrentSheetId())).toBe("Sheet4");
    await expect(page.getByTestId("sheet-position")).toHaveText("Sheet 4 of 5");
    await expect(page.getByTestId("sheet-tab-Sheet4")).toHaveAttribute("data-active", "true");

    // Sheet activation via the overflow/quick-pick UI should return focus to the grid so
    // keyboard-driven workflows keep working (Excel-like).
    await expect
      .poll(() => page.evaluate(() => (document.activeElement as HTMLElement | null)?.id))
      .toBe("grid");

    await page.keyboard.press("F2");
    await expect(page.locator("textarea.cell-editor")).toBeVisible();
  });

  test("add sheet inserts after the active sheet tab", async ({ page }) => {
    await gotoDesktop(page);

    // Ensure A1 is active so the status bar is deterministic after sheet switches.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.activateCell({ row: 0, col: 0 });
    });

    await expect(page.getByTestId("sheet-tab-Sheet1")).toBeVisible();

    // Add Sheet2 while Sheet1 is active -> [Sheet1, Sheet2].
    await page.getByTestId("sheet-add").click();

    await expect
      .poll(() => page.evaluate(() => (window as any).__formulaApp.getCurrentSheetId()))
      .not.toBe("Sheet1");

    const sheet2 = await page.evaluate(() => String((window as any).__formulaApp.getCurrentSheetId()));
    await expect(page.getByTestId(`sheet-tab-${sheet2}`)).toBeVisible();

    await expect.poll(() =>
      page.evaluate(() =>
        Array.from(document.querySelectorAll("#sheet-tabs .sheet-tabs [data-sheet-id]"))
          .map((el) => (el as HTMLElement).getAttribute("data-sheet-id"))
          .filter(Boolean),
      ),
    ).toEqual(["Sheet1", sheet2]);

    // Activate Sheet1, then add Sheet3 -> [Sheet1, Sheet3, Sheet2].
    await page.getByTestId("sheet-tab-Sheet1").click();
    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCurrentSheetId())).toBe("Sheet1");

    await page.getByTestId("sheet-add").click();
    await expect
      .poll(() =>
        page.evaluate((sheet2Id) => {
          const id = (window as any).__formulaApp.getCurrentSheetId();
          if (id === "Sheet1" || id === sheet2Id) return "";
          return id;
        }, sheet2),
      )
      .not.toBe("");

    const sheet3 = await page.evaluate(() => String((window as any).__formulaApp.getCurrentSheetId()));
    expect(sheet3).not.toBe(sheet2);
    await expect(page.getByTestId(`sheet-tab-${sheet3}`)).toBeVisible();

    await expect.poll(() =>
      page.evaluate(() =>
        Array.from(document.querySelectorAll("#sheet-tabs .sheet-tabs [data-sheet-id]"))
          .map((el) => (el as HTMLElement).getAttribute("data-sheet-id"))
          .filter(Boolean),
      ),
    ).toEqual(["Sheet1", sheet3, sheet2]);
  });

  test("keyboard navigation activates the focused sheet tab", async ({ page }) => {
    await gotoDesktop(page);

    // Ensure A1 is active before switching sheets so the status bar reflects A1 values.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.activateCell({ row: 0, col: 0 });
    });

    // Lazily create Sheet2 by writing a value into it.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.getDocument().setCellValue("Sheet2", "A1", "Hello from Sheet2");
    });
    await expect(page.getByRole("tab", { name: "Sheet2" })).toBeVisible();

    // Focus the tab strip. Tab is used for in-grid navigation, so focus it directly for this test.
    const sheet1Tab = page.getByRole("tab", { name: "Sheet1" });
    await sheet1Tab.focus();
    await expect(sheet1Tab).toBeFocused();

    // Arrow to Sheet2, then activate it.
    await page.keyboard.press("ArrowRight");
    const sheet2Tab = page.getByRole("tab", { name: "Sheet2" });
    await expect(sheet2Tab).toBeFocused();

    await page.keyboard.press("Enter");
    await expect(page.getByRole("tab", { name: "Sheet2" })).toHaveAttribute("aria-selected", "true");
    await expect(page.getByTestId("active-value")).toHaveText("Hello from Sheet2");

    // Sheet activation should return focus to the grid so keyboard navigation can continue.
    await expect
      .poll(() => page.evaluate(() => (document.activeElement as HTMLElement | null)?.id))
      .toBe("grid");
  });

  test("Shift+F10 opens the context menu for the focused sheet tab", async ({ page }) => {
    await gotoDesktop(page);

    const sheet1Tab = page.getByRole("tab", { name: "Sheet1" });
    // The grid captures Tab for in-grid navigation; focus the tab directly for this test.
    await sheet1Tab.focus();
    await expect(sheet1Tab).toBeFocused();

    await page.keyboard.press("Shift+F10");
    const menu = page.getByTestId("sheet-tab-context-menu");
    await expect(menu).toBeVisible();
    await expect(menu.getByRole("button", { name: "Rename" })).toBeVisible();

    await page.keyboard.press("Escape");
    await expect(menu).toBeHidden();
    await expect(sheet1Tab).toBeFocused();
  });

  test("double-click rename commits on Enter and updates tab + switcher labels", async ({ page }) => {
    await gotoDesktop(page);

    const tab = page.getByTestId("sheet-tab-Sheet1");
    await expect(tab).toBeVisible();

    await tab.dblclick();
    const input = tab.locator("input.sheet-tab__input");
    await expect(input).toBeVisible();
    await expect(input).toBeFocused();

    await input.fill("Budget");
    await input.press("Enter");

    await expect(tab.locator(".sheet-tab__name")).toHaveText("Budget");
    await expect(page.getByTestId("sheet-switcher").locator('option[value="Sheet1"]')).toHaveText("Budget");
  });

  test("sheet switcher select activates sheets and updates sheet position", async ({ page }) => {
    await gotoDesktop(page);

    // Ensure A1 is active so subsequent sheet activation has a deterministic cell focus.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.activateCell({ row: 0, col: 0 });
      app.getDocument().setCellValue("Sheet2", "A1", "Hello from Sheet2");
      app.getDocument().setCellValue("Sheet3", "A1", "Hello from Sheet3");
    });

    await expect(page.getByTestId("sheet-position")).toHaveText("Sheet 1 of 3");

    const switcher = page.getByTestId("sheet-switcher");
    await switcher.selectOption("Sheet3", { force: true });
    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCurrentSheetId())).toBe("Sheet3");
    await expect(page.getByTestId("sheet-position")).toHaveText("Sheet 3 of 3");

    await switcher.selectOption("Sheet2", { force: true });
    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCurrentSheetId())).toBe("Sheet2");
    await expect(page.getByTestId("sheet-position")).toHaveText("Sheet 2 of 3");
  });

  test("renaming a sheet rewrites formulas that reference it", async ({ page }) => {
    await gotoDesktop(page);

    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const store =
        typeof app.getWorkbookSheetStore === "function"
          ? app.getWorkbookSheetStore()
          : (window.__workbookSheetStore as any);

      // Create a sheet whose id is not equal to its user-visible name.
      store.addAfter("Sheet1", { id: "s2", name: "Budget" });
      app.getDocument().setCellValue("s2", "A1", 123);
      app.getDocument().setCellFormula("Sheet1", "B1", "=Budget!A1+1");
    });

    const before = await page.evaluate(async () => {
      const app = (window as any).__formulaApp;
      await app.whenIdle();
      return app.getCellDisplayValueA1("B1");
    });
    expect(before).toBe("124");

    const tab = page.getByTestId("sheet-tab-s2");
    await expect(tab.locator(".sheet-tab__name")).toHaveText("Budget");

    await tab.dblclick();
    const input = tab.locator("input.sheet-tab__input");
    await expect(input).toBeVisible();
    await input.fill("Budget2");
    await input.press("Enter");

    await expect(tab.locator(".sheet-tab__name")).toHaveText("Budget2");

    const after = await page.evaluate(async () => {
      const app = (window as any).__formulaApp;
      // Double-clicking a tab can change the active sheet; ensure we're evaluating the
      // formula cell on Sheet1.
      app.activateSheet("Sheet1");
      await app.whenIdle();
      return app.getCellDisplayValueA1("B1");
    });
    expect(after).toBe("124");

    const updatedFormula = await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      return app.getDocument().getCell("Sheet1", "B1").formula;
    });
    expect(updatedFormula).toBe("=Budget2!A1+1");
  });

  test("renaming a sheet rewrites formulas that reference its display name", async ({ page }) => {
    await gotoDesktop(page);

    // Create Sheet2 via the UI so both the DocumentController and the sheet metadata
    // store agree on the sheet list.
    await page.getByTestId("sheet-add").click();
    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();

    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const doc = app.getDocument();
      doc.setCellValue("Sheet1", "A1", 123);
      doc.setCellFormula("Sheet2", "A1", "=Sheet1!A1");
    });

    // Sanity check: formula should initially evaluate correctly.
    await expect.poll(() =>
      page.evaluate(() => (window as any).__formulaApp.getCellComputedValueForSheet("Sheet2", { row: 0, col: 0 })),
    ).toBe(123);

    // Rename Sheet1 via tab strip inline rename.
    const sheet1Tab = page.getByTestId("sheet-tab-Sheet1");
    await sheet1Tab.dblclick();
    const input = sheet1Tab.locator("input.sheet-tab__input");
    await expect(input).toBeVisible();
    await input.fill("RenamedSheet1");
    await input.press("Enter");

    // Excel-style behavior: formulas should be rewritten from Sheet1 -> RenamedSheet1.
    await expect.poll(() =>
      page.evaluate(() => (window as any).__formulaApp.getDocument().getCell("Sheet2", "A1").formula),
    ).toContain("RenamedSheet1");

    // And the formula should still evaluate (it shouldn't degrade to #REF!).
    await expect.poll(() =>
      page.evaluate(() => (window as any).__formulaApp.getCellComputedValueForSheet("Sheet2", { row: 0, col: 0 })),
    ).toBe(123);
  });

  test("rename cancels on Escape (does not commit via blur)", async ({ page }) => {
    await gotoDesktop(page);

    const tab = page.getByTestId("sheet-tab-Sheet1");
    await expect(tab).toBeVisible();

    await tab.dblclick();
    const input = tab.locator("input.sheet-tab__input");
    await expect(input).toBeVisible();

    await input.fill("ShouldNotCommit");
    await input.press("Escape");

    // Give the UI a moment to apply any accidental blur commit.
    await expect(tab).toContainText("Sheet1");
    await expect(tab).not.toContainText("ShouldNotCommit");
  });

  test("Ctrl+PgUp/PgDn does not switch sheets while a sheet tab rename input is focused", async ({ page }) => {
    await gotoDesktop(page);

    // Create a second sheet so sheet navigation would be observable.
    await page.getByTestId("sheet-add").click();
    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();

    // Ensure Sheet1 is active, then enter rename mode.
    const sheet1Tab = page.getByTestId("sheet-tab-Sheet1");
    await sheet1Tab.click();
    await expect(sheet1Tab).toHaveAttribute("data-active", "true");
    const initialSheetId = await page.evaluate(() => (window as any).__formulaApp.getCurrentSheetId());

    await sheet1Tab.dblclick();
    const input = sheet1Tab.locator("input.sheet-tab__input");
    await expect(input).toBeVisible();
    await expect(input).toBeFocused();

    // Dispatch directly to the input so the global shortcut handler must *not* steal it.
    await page.evaluate(() => {
      const input = document.querySelector("input.sheet-tab__input") as HTMLInputElement | null;
      if (!input) throw new Error("Missing rename input");
      const evt = new KeyboardEvent("keydown", { key: "PageDown", ctrlKey: true, bubbles: true, cancelable: true });
      input.dispatchEvent(evt);
    });

    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCurrentSheetId())).toBe(initialSheetId);
    await expect(input).toBeVisible();
    await expect(input).toBeFocused();
  });

  test("Ctrl+PgUp/PgDn does not switch sheets while the formula bar input is focused", async ({ page }) => {
    await gotoDesktop(page);

    // Create a second sheet so sheet navigation would be observable.
    await page.getByTestId("sheet-add").click();
    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();

    // Focus the formula bar input (it is a textarea).
    //
    // In view mode the formula bar input is hidden and hover events are handled by the
    // highlighted <pre> below, so click the highlight to enter edit mode and focus the textarea.
    await page.getByTestId("formula-highlight").click();
    const formulaInput = page.getByTestId("formula-input");
    await expect(formulaInput).toBeVisible();
    await expect(formulaInput).toBeFocused();
    const initialSheetId = await page.evaluate(() => (window as any).__formulaApp.getCurrentSheetId());

    await page.evaluate(() => {
      const input = document.querySelector('[data-testid="formula-input"]') as HTMLTextAreaElement | null;
      if (!input) throw new Error("Missing formula bar input");
      const evt = new KeyboardEvent("keydown", { key: "PageDown", ctrlKey: true, bubbles: true, cancelable: true });
      input.dispatchEvent(evt);
    });

    // Active sheet should remain unchanged.
    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCurrentSheetId())).toBe(initialSheetId);
  });

  test("Ctrl+PgUp/PgDn switches sheets while the formula bar is actively editing a formula", async ({ page }) => {
    await gotoDesktop(page);

    // Create Sheet2 so sheet navigation is observable.
    await page.getByTestId("sheet-add").click();
    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();

    // Start on Sheet1.
    await page.getByTestId("sheet-tab-Sheet1").click();
    await expect(page.getByTestId("sheet-tab-Sheet1")).toHaveAttribute("data-active", "true");

    // Enter formula bar edit mode and ensure the draft is a *formula* (starts with "="),
    // which should allow sheet navigation (Excel behavior for cross-sheet references).
    await page.getByTestId("formula-highlight").click();
    const formulaInput = page.getByTestId("formula-input");
    await expect(formulaInput).toBeVisible();
    await expect(formulaInput).toBeFocused();
    await formulaInput.fill("=");
    await expect(formulaInput).toHaveValue("=");
    // `fill()` updates the DOM value immediately, but the formula bar model updates via an
    // `input` event handler. Wait until the app reports it is in formula-editing mode so the
    // Ctrl+PgUp/PgDn global shortcut can treat this as an Excel-like range-selection session.
    await expect
      .poll(() =>
        page.evaluate(() => {
          const app = (window as any).__formulaApp;
          return Boolean(app?.isFormulaBarFormulaEditing?.());
        }),
      )
      .toBe(true);

    // Dispatch Ctrl+PgDn directly to the textarea so the global handler must allow it.
    await page.evaluate(() => {
      const input = document.querySelector('[data-testid="formula-input"]') as HTMLTextAreaElement | null;
      if (!input) throw new Error("Missing formula bar input");
      const evt = new KeyboardEvent("keydown", { key: "PageDown", ctrlKey: true, bubbles: true, cancelable: true });
      input.dispatchEvent(evt);
    });

    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCurrentSheetId())).toBe("Sheet2");
    await expect(page.getByTestId("sheet-tab-Sheet2")).toHaveAttribute("data-active", "true");
    await expect(formulaInput).toBeFocused();

    // And Ctrl+PgUp should bring us back while staying in formula edit mode.
    await page.evaluate(() => {
      const input = document.querySelector('[data-testid="formula-input"]') as HTMLTextAreaElement | null;
      if (!input) throw new Error("Missing formula bar input");
      const evt = new KeyboardEvent("keydown", { key: "PageUp", ctrlKey: true, bubbles: true, cancelable: true });
      input.dispatchEvent(evt);
    });

    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCurrentSheetId())).toBe("Sheet1");
    await expect(page.getByTestId("sheet-tab-Sheet1")).toHaveAttribute("data-active", "true");
    await expect(formulaInput).toBeFocused();
    await expect(formulaInput).toHaveValue("=");
  });

  test("Ctrl+PgUp/PgDn does not switch sheets while the cell editor is open", async ({ page }) => {
    await gotoDesktop(page);

    // Create a second sheet so sheet navigation would be observable.
    await page.getByTestId("sheet-add").click();
    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();

    // Return to Sheet1 and open the in-cell editor.
    const sheet1Tab = page.getByTestId("sheet-tab-Sheet1");
    await sheet1Tab.click();
    await expect(sheet1Tab).toHaveAttribute("data-active", "true");

    // Click inside A1 (avoid the shared-grid corner header/select-all region).
    await page.click("#grid", { position: { x: 80, y: 40 } });
    await page.keyboard.press("F2");

    const editor = page.locator("textarea.cell-editor");
    await expect(editor).toBeVisible();
    await expect(editor).toBeFocused();

    // Dispatch directly to the textarea so the global shortcut handler must *not* steal it.
    await page.evaluate(() => {
      const editor = document.querySelector("textarea.cell-editor") as HTMLTextAreaElement | null;
      if (!editor) throw new Error("Missing cell editor textarea");
      const evt = new KeyboardEvent("keydown", { key: "PageDown", ctrlKey: true, bubbles: true, cancelable: true });
      editor.dispatchEvent(evt);
    });

    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCurrentSheetId())).toBe("Sheet1");
    await expect(editor).toBeVisible();
    await expect(editor).toBeFocused();
  });

  test("invalid rename (forbidden characters) shows a toast and does not rename the sheet", async ({ page }) => {
    await gotoDesktop(page);

    // Create a second sheet we can attempt to switch to.
    await page.getByTestId("sheet-add").click();
    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();

    // Switch back to Sheet1 and begin renaming.
    const sheet1Tab = page.getByTestId("sheet-tab-Sheet1");
    await sheet1Tab.click();
    await expect(sheet1Tab).toHaveAttribute("data-active", "true");

    await sheet1Tab.dblclick();
    const input = sheet1Tab.locator("input.sheet-tab__input");
    await expect(input).toBeVisible();

    // Trigger an invalid name (contains a forbidden character) and attempt to commit.
    await input.fill("A/B");
    await input.press("Enter");

    await expect(page.locator('[data-testid="toast"]').filter({ hasText: /invalid character/i })).toBeVisible();

    // Cancel the in-progress rename; the sheet name should remain unchanged.
    await input.press("Escape");
    await expect(sheet1Tab.locator(".sheet-tab__name")).toHaveText("Sheet1");

    // Switching sheets remains possible after the invalid attempt.
    await page.getByTestId("sheet-tab-Sheet2").click();
    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCurrentSheetId())).toBe("Sheet2");
  });

  test("invalid rename does not block opening the sheet overflow menu", async ({ page }) => {
    await gotoDesktop(page);

    // Create a second sheet so the overflow menu would normally show multiple options.
    await page.getByTestId("sheet-add").click();
    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();

    // Switch back to Sheet1 and begin renaming.
    const sheet1Tab = page.getByTestId("sheet-tab-Sheet1");
    await sheet1Tab.click();
    await expect(sheet1Tab).toHaveAttribute("data-active", "true");

    await sheet1Tab.dblclick();
    const input = sheet1Tab.locator("input.sheet-tab__input");
    await expect(input).toBeVisible();

    // Trigger an invalid name and attempt to commit.
    await input.fill("A/B");
    await input.press("Enter");
    await expect(page.locator('[data-testid="toast"]').filter({ hasText: /invalid character/i })).toBeVisible();

    // Cancel rename to exit edit mode, then open the overflow menu. Invalid rename should not wedge the sheet UI.
    await input.press("Escape");
    await page.getByTestId("sheet-overflow").click();
    const quickPick = page.getByTestId("quick-pick");
    await expect(quickPick).toBeVisible();
    await page.keyboard.press("Escape");
    await expect(quickPick).toHaveCount(0);

    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCurrentSheetId())).toBe("Sheet1");
  });

  test("invalid rename does not block switching sheets via the sheet switcher <select>", async ({ page }) => {
    await gotoDesktop(page);

    // Create a second sheet we can attempt to switch to.
    await page.getByTestId("sheet-add").click();
    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();

    // Begin renaming Sheet1 and enter an invalid name.
    const sheet1Tab = page.getByTestId("sheet-tab-Sheet1");
    await sheet1Tab.click();
    await expect(sheet1Tab).toHaveAttribute("data-active", "true");

    await sheet1Tab.dblclick();
    const input = sheet1Tab.locator("input.sheet-tab__input");
    await expect(input).toBeVisible();
    await input.fill("A/B");
    await input.press("Enter");
    await expect(page.locator('[data-testid="toast"]').filter({ hasText: /invalid character/i })).toBeVisible();

    // Cancel rename; sheet name remains unchanged.
    await input.press("Escape");
    await expect(sheet1Tab.locator(".sheet-tab__name")).toHaveText("Sheet1");

    // Attempt to switch via the status-bar sheet switcher. Invalid rename should not wedge the sheet UI.
    const switcher = page.getByTestId("sheet-switcher");
    await switcher.selectOption("Sheet2", { force: true });

    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCurrentSheetId())).toBe("Sheet2");
    await expect(page.getByTestId("sheet-tab-Sheet2")).toHaveAttribute("data-active", "true");
    await expect(switcher).toHaveValue("Sheet2");
  });

  test("drag reordering sheet tabs updates Ctrl+PgUp/PgDn navigation order", async ({ page }) => {
    await gotoDesktop(page);

    // Ensure A1 is active so Ctrl+PgUp/PgDn starts from a deterministic sheet/cell.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.activateCell({ row: 0, col: 0 });
    });

    // Create Sheet2 + Sheet3 via the UI.
    await page.getByTestId("sheet-add").click();
    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();
    await page.getByTestId("sheet-add").click();
    await expect(page.getByTestId("sheet-tab-Sheet3")).toBeVisible();
    await expect(page.getByTestId("sheet-tab-Sheet3")).toHaveAttribute("data-active", "true");
    await expect(page.getByTestId("sheet-position")).toHaveText("Sheet 3 of 3");

    // Move Sheet3 before Sheet1 (new order: Sheet3, Sheet1, Sheet2).
    // Use a synthetic HTML5 drop event for determinism (Playwright drag/drop can be flaky).
    const desiredOrder = ["Sheet3", "Sheet1", "Sheet2"];
    await page.evaluate(() => {
      const fromId = "Sheet3";
      const target = document.querySelector('[data-testid="sheet-tab-Sheet1"]') as HTMLElement | null;
      if (!target) throw new Error("Missing Sheet1 tab");
      const rect = target.getBoundingClientRect();

      const dt = new DataTransfer();
      dt.setData("text/sheet-id", fromId);
      dt.setData("text/plain", fromId);

      const drop = new DragEvent("drop", {
        bubbles: true,
        cancelable: true,
        clientX: rect.left + 1,
        clientY: rect.top + rect.height / 2,
      });
      Object.defineProperty(drop, "dataTransfer", { value: dt });
      target.dispatchEvent(drop);
    });

    await expect.poll(() =>
      page.evaluate(() =>
        Array.from(document.querySelectorAll("#sheet-tabs .sheet-tabs [data-sheet-id]")).map((el) =>
          (el as HTMLElement).getAttribute("data-sheet-id"),
        ),
      ),
    ).toEqual(desiredOrder);
    // Active sheet is still Sheet3, but its position is now 1st.
    await expect(page.getByTestId("sheet-position")).toHaveText("Sheet 1 of 3");

    // Focus the grid: Ctrl/Cmd+PgUp/PgDn must work when the grid is focused (real workflow).
    // Click inside A1 (avoid the shared-grid corner header/select-all region).
    await page.click("#grid", { position: { x: 80, y: 40 } });
    await expect
      .poll(() => page.evaluate(() => (document.activeElement as HTMLElement | null)?.id))
      .toBe("grid");

    // Ctrl+PgDn should follow the new order. Dispatch the key event directly on the grid element
    // (not window) to ensure we exercise the focused-grid shortcut path.
    await page.evaluate(() => {
      const grid = document.getElementById("grid");
      if (!grid) throw new Error("Missing #grid");
      const evt = new KeyboardEvent("keydown", {
        key: "PageDown",
        ctrlKey: true,
        bubbles: true,
        cancelable: true,
      });
      grid.dispatchEvent(evt);
    });
    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCurrentSheetId())).toBe("Sheet1");
    await expect(page.getByTestId("sheet-position")).toHaveText("Sheet 2 of 3");

    await page.evaluate(() => {
      const grid = document.getElementById("grid");
      if (!grid) throw new Error("Missing #grid");
      const evt = new KeyboardEvent("keydown", {
        key: "PageDown",
        ctrlKey: true,
        bubbles: true,
        cancelable: true,
      });
      grid.dispatchEvent(evt);
    });
    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCurrentSheetId())).toBe("Sheet2");
    await expect(page.getByTestId("sheet-position")).toHaveText("Sheet 3 of 3");

    // Wrap around to Sheet3.
    await page.evaluate(() => {
      const grid = document.getElementById("grid");
      if (!grid) throw new Error("Missing #grid");
      const evt = new KeyboardEvent("keydown", {
        key: "PageDown",
        ctrlKey: true,
        bubbles: true,
        cancelable: true,
      });
      grid.dispatchEvent(evt);
    });
    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCurrentSheetId())).toBe("Sheet3");
    await expect(page.getByTestId("sheet-position")).toHaveText("Sheet 1 of 3");
  });

  test("dragging Sheet2 before Sheet1 reorders the tab strip, updates the sheet switcher, and marks the document dirty", async ({ page }) => {
    await gotoDesktop(page);

    // Keep Sheet1 active so we can assert it stays active after reordering.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.activateCell({ row: 0, col: 0 });
    });
    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCurrentSheetId())).toBe("Sheet1");

    // Lazily create Sheet2 (DocumentController creates sheets on demand).
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.getDocument().setCellValue("Sheet2", "A1", "Hello from Sheet2");
    });
    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();

    // Reset the dirty bit so we can attribute it to the tab reorder, not sheet creation.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.getDocument().markSaved();
    });
    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getDocument().isDirty)).toBe(false);

    // Move Sheet2 before Sheet1 (new order: Sheet2, Sheet1).
    const desiredOrder = ["Sheet2", "Sheet1"];
    await page.evaluate(() => {
      const fromId = "Sheet2";
      const target = document.querySelector('[data-testid="sheet-tab-Sheet1"]') as HTMLElement | null;
      if (!target) throw new Error("Missing Sheet1 tab");
      const rect = target.getBoundingClientRect();

      const dt = new DataTransfer();
      dt.setData("text/sheet-id", fromId);
      dt.setData("text/plain", fromId);

      const drop = new DragEvent("drop", {
        bubbles: true,
        cancelable: true,
        clientX: rect.left + 1,
        clientY: rect.top + rect.height / 2,
      });
      Object.defineProperty(drop, "dataTransfer", { value: dt });
      target.dispatchEvent(drop);
    });

    await expect.poll(() =>
      page.evaluate(() =>
        Array.from(document.querySelectorAll("#sheet-tabs .sheet-tabs [data-sheet-id]")).map((el) =>
          (el as HTMLElement).getAttribute("data-sheet-id"),
        ),
      ),
    ).toEqual(desiredOrder);

    // Active sheet remains Sheet1.
    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCurrentSheetId())).toBe("Sheet1");
    await expect(page.getByTestId("sheet-tab-Sheet1")).toHaveAttribute("data-active", "true");

    // The sheet switcher should reflect the tab ordering.
    const options = page.getByTestId("sheet-switcher").locator("option");
    await expect(options.nth(0)).toHaveAttribute("value", "Sheet2");
    await expect(options.nth(1)).toHaveAttribute("value", "Sheet1");
    await expect(page.getByTestId("sheet-switcher")).toHaveValue("Sheet1");

    // Reordering marks the workbook dirty.
    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getDocument().isDirty)).toBe(true);
  });

  test("drag reorder maps visible tab positions onto full sheet order with hidden sheets", async ({ page }) => {
    await gotoDesktop(page);

    // Create Sheet2 + Sheet3 via the UI so both the DocumentController and metadata store
    // include them. (We need a hidden sheet in between two visible sheets.)
    await page.getByTestId("sheet-add").click();
    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();
    await page.getByTestId("sheet-add").click();
    await expect(page.getByTestId("sheet-tab-Sheet3")).toBeVisible();

    // Hide Sheet2 so the visible tabs are [Sheet1, Sheet3] while the full order is
    // [Sheet1, Sheet2(hidden), Sheet3].
    await page.getByTestId("sheet-tab-Sheet2").click({ button: "right", position: { x: 10, y: 10 } });
    {
      const menu = page.getByTestId("sheet-tab-context-menu");
      // Ensure the sheet tab context menu is the one that's open.
      await expect(page.getByTestId("context-menu")).toBeHidden();
      await expect(menu).toBeVisible();
      await menu.getByRole("button", { name: "Hide", exact: true }).click();
      await expect(menu).toBeHidden();
    }
    await expect(page.getByTestId("sheet-tab-Sheet2")).toHaveCount(0);

    // Drag Sheet3 before Sheet1.
    //
    // Note: avoid Playwright's `dragTo` here (can hang in the desktop shell); dispatch a
    // synthetic drop event instead.
    const desiredAll = [
      { id: "Sheet3", visibility: "visible" },
      { id: "Sheet1", visibility: "visible" },
      { id: "Sheet2", visibility: "hidden" },
    ];
    await page.evaluate(() => {
      const fromId = "Sheet3";
      const target = document.querySelector('[data-testid="sheet-tab-Sheet1"]') as HTMLElement | null;
      if (!target) throw new Error("Missing Sheet1 tab");
      const rect = target.getBoundingClientRect();

      const dt = new DataTransfer();
      dt.setData("text/sheet-id", fromId);
      dt.setData("text/plain", fromId);

      const drop = new DragEvent("drop", {
        bubbles: true,
        cancelable: true,
        clientX: rect.left + 1,
        clientY: rect.top + rect.height / 2,
      });
      Object.defineProperty(drop, "dataTransfer", { value: dt });
      target.dispatchEvent(drop);
    });

    await expect
      .poll(() =>
        page.evaluate(() => {
          const app = (window as any).__formulaApp;
          return app.getWorkbookSheetStore().listAll().map((s: any) => ({ id: s.id, visibility: s.visibility }));
        }),
      )
      .toEqual(desiredAll);

    // Unhide Sheet2 and ensure it lands at the end (i.e. hidden sheet moved as expected).
    await page.getByTestId("sheet-tab-Sheet1").click({ button: "right", position: { x: 10, y: 10 } });
    {
      const menu = page.getByTestId("sheet-tab-context-menu");
      await expect(page.getByTestId("context-menu")).toBeHidden();
      await expect(menu).toBeVisible();
      await menu.getByRole("button", { name: "Unhide…" }).click();
      await menu.getByRole("button", { name: "Sheet2" }).click();
    }
    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();

    await expect
      .poll(() =>
        page.evaluate(() => {
          const app = (window as any).__formulaApp;
          return app.getWorkbookSheetStore().listVisible().map((s: any) => s.id);
        }),
      )
      .toEqual(["Sheet3", "Sheet1", "Sheet2"]);
  });

  test("sheet position indicator uses visible sheets (hide/unhide)", async ({ page }) => {
    await gotoDesktop(page);

    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.activateCell({ row: 0, col: 0 });
      app.getDocument().setCellValue("Sheet2", "A1", "Hello from Sheet2");
      app.getDocument().setCellValue("Sheet3", "A1", "Hello from Sheet3");
    });

    await expect(page.getByTestId("sheet-position")).toHaveText("Sheet 1 of 3");

    await page.getByTestId("sheet-tab-Sheet3").click();
    await expect(page.getByTestId("sheet-position")).toHaveText("Sheet 3 of 3");

    // Hide Sheet2 so the position indicator reflects visible sheets (Excel-like behavior).
    await page.getByTestId("sheet-tab-Sheet2").click({ button: "right", position: { x: 10, y: 10 } });
    {
      const menu = page.getByTestId("sheet-tab-context-menu");
      await expect(page.getByTestId("context-menu")).toBeHidden();
      await expect(menu).toBeVisible();
      await menu.getByRole("button", { name: "Hide", exact: true }).click();
    }
    await expect(page.getByTestId("sheet-tab-Sheet2")).toHaveCount(0);

    await expect(page.getByTestId("sheet-position")).toHaveText("Sheet 2 of 2");

    // Unhide Sheet2 and ensure position updates.
    await page.getByTestId("sheet-tab-Sheet1").click({ button: "right", position: { x: 10, y: 10 } });
    {
      const menu = page.getByTestId("sheet-tab-context-menu");
      await expect(page.getByTestId("context-menu")).toBeHidden();
      await expect(menu).toBeVisible();
      await menu.getByRole("button", { name: "Unhide…" }).click();
      await menu.getByRole("button", { name: "Sheet2" }).click();
    }
    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();

    await expect(page.getByTestId("sheet-position")).toHaveText("Sheet 3 of 3");
  });

  test("renaming a sheet marks the document dirty", async ({ page }) => {
    await gotoDesktop(page);

    // Demo workbook is treated as an initial saved baseline.
    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getDocument().isDirty)).toBe(false);

    const tab = page.getByTestId("sheet-tab-Sheet1");
    await tab.dblclick();
    const input = tab.locator("input.sheet-tab__input");
    await expect(input).toBeVisible();
    await input.fill("RenamedSheet1");
    await input.press("Enter");

    await expect(tab.locator(".sheet-tab__name")).toHaveText("RenamedSheet1");

    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getDocument().isDirty)).toBe(true);
  });

  test("command palette go to resolves renamed sheet display names back to stable sheet ids", async ({ page }) => {
    await gotoDesktop(page);

    // Rename Sheet1 -> RenamedSheet1 (display name changes, stable id stays Sheet1).
    const sheet1Tab = page.getByTestId("sheet-tab-Sheet1");
    await sheet1Tab.dblclick();
    const renameInput = sheet1Tab.locator("input.sheet-tab__input");
    await expect(renameInput).toBeVisible();
    await renameInput.fill("RenamedSheet1");
    await renameInput.press("Enter");
    await expect(sheet1Tab.locator(".sheet-tab__name")).toHaveText("RenamedSheet1");

    // Create Sheet2 and switch away so the go-to command must resolve and navigate.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.getDocument().setCellValue("Sheet2", "A1", "Hello from Sheet2");
    });
    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();

    await page.getByTestId("sheet-tab-Sheet2").click();
    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCurrentSheetId())).toBe("Sheet2");

    // Open command palette via the command registry (same path used by extensions).
    await page.evaluate(async () => {
      await (window.__formulaCommandRegistry as any).executeCommand("workbench.showCommandPalette");
    });
    await expect(page.getByTestId("command-palette-input")).toBeVisible();

    await page.getByTestId("command-palette-input").fill("RenamedSheet1!A1");
    await expect(page.locator('[data-testid="command-palette-list"] .command-palette__item').first()).toContainText(
      "RenamedSheet1!A1",
    );
    await page.keyboard.press("Enter");

    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCurrentSheetId())).toBe("Sheet1");
    await expect(page.getByTestId("sheet-tab-Sheet1")).toHaveAttribute("data-active", "true");
  });

  test("switching sheets restores grid focus for immediate keyboard editing", async ({ page }) => {
    await gotoDesktop(page);

    // Start with focus on the grid.
    // Click inside A1 (avoid the shared-grid corner header/select-all region).
    await page.click("#grid", { position: { x: 80, y: 40 } });

    // Lazily create Sheet2 so there's another tab to switch to.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.getDocument().setCellValue("Sheet2", "A1", "Hello from Sheet2");
    });
    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();

    // Switching sheets should behave like navigation and leave the grid ready
    // for keyboard-driven workflows (Excel-like).
    await page.getByTestId("sheet-tab-Sheet2").click();
    await page.keyboard.press("F2");
    await expect(page.locator("textarea.cell-editor")).toBeVisible();
  });

  test("Ctrl+PgDn cycles through visible sheets and wraps (skips hidden)", async ({ page }) => {
    await gotoDesktop(page);

    // Create Sheet2 + Sheet3 using the "+" button.
    await page.getByTestId("sheet-add").click();
    await page.getByTestId("sheet-add").click();

    await expect(page.getByTestId("sheet-tab-Sheet3")).toBeVisible();

    // Return to Sheet1 so Ctrl+PgDn navigation starts from the beginning.
    await page.getByTestId("sheet-tab-Sheet1").click();
    await expect(page.getByTestId("sheet-tab-Sheet1")).toHaveAttribute("data-active", "true");

    // Hide the middle sheet.
    await page.getByTestId("sheet-tab-Sheet2").click({ button: "right", position: { x: 10, y: 10 } });
    {
      const menu = page.getByTestId("sheet-tab-context-menu");
      await expect(page.getByTestId("context-menu")).toBeHidden();
      await expect(menu).toBeVisible();
      await menu.getByRole("button", { name: "Hide", exact: true }).click();
    }
    await expect(page.getByTestId("sheet-tab-Sheet2")).toHaveCount(0);

    // Sheet1 -> Sheet3 (Sheet2 is hidden).
    await page.evaluate(() => {
      const grid = document.getElementById("grid");
      if (!grid) throw new Error("Missing #grid");
      grid.focus();
      const evt = new KeyboardEvent("keydown", { key: "PageDown", ctrlKey: true, bubbles: true, cancelable: true });
      grid.dispatchEvent(evt);
    });
    await expect(page.getByTestId("sheet-tab-Sheet3")).toHaveAttribute("data-active", "true");

    // Wrap Sheet3 -> Sheet1.
    await page.evaluate(() => {
      const grid = document.getElementById("grid");
      if (!grid) throw new Error("Missing #grid");
      const evt = new KeyboardEvent("keydown", { key: "PageDown", ctrlKey: true, bubbles: true, cancelable: true });
      grid.dispatchEvent(evt);
    });
    await expect(page.getByTestId("sheet-tab-Sheet1")).toHaveAttribute("data-active", "true");

    // Ctrl+PgUp should follow visible sheets as well (wrap Sheet1 -> Sheet3).
    await page.evaluate(() => {
      const grid = document.getElementById("grid");
      if (!grid) throw new Error("Missing #grid");
      const evt = new KeyboardEvent("keydown", { key: "PageUp", ctrlKey: true, bubbles: true, cancelable: true });
      grid.dispatchEvent(evt);
    });
    await expect(page.getByTestId("sheet-tab-Sheet3")).toHaveAttribute("data-active", "true");
  });

  test("add sheet inserts immediately after the active sheet", async ({ page }) => {
    await page.addInitScript(() => {
      const listeners: Record<string, any> = {};
      (window as any).__tauriListeners = listeners;

      (window as any).__TAURI__ = {
        core: {
          invoke: async (cmd: string, args: any) => {
            switch (cmd) {
              case "add_sheet":
                return { id: args?.name ?? "SheetX", name: args?.name ?? "SheetX" };
              default:
                return null;
            }
          },
        },
        event: {
          listen: async (name: string, handler: any) => {
            listeners[name] = handler;
            return () => {
              delete listeners[name];
            };
          },
        },
      };
    });

    await gotoDesktop(page);

    // Create Sheet2 + Sheet3 so we have three tabs.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const doc = app.getDocument();
      doc.setCellValue("Sheet2", "A1", "Two");
      doc.setCellValue("Sheet3", "A1", "Three");
    });

    await expect(page.getByTestId("sheet-tab-Sheet1")).toBeVisible();
    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();
    await expect(page.getByTestId("sheet-tab-Sheet3")).toBeVisible();

    // Activate the middle sheet (Sheet2) and add a sheet.
    await page.getByTestId("sheet-tab-Sheet2").click();
    await page.getByTestId("sheet-add").click();

    await expect(page.getByTestId("sheet-tab-Sheet4")).toBeVisible();

    const ids = await page.evaluate(() => {
      const root = document.querySelector('[data-testid="sheet-tabs"]');
      if (!root) return [];
      return Array.from(root.querySelectorAll<HTMLButtonElement>("button[data-sheet-id]")).map(
        (btn) => btn.dataset.sheetId ?? "",
      );
    });

    expect(ids).toEqual(["Sheet1", "Sheet2", "Sheet4", "Sheet3"]);
  });

  test("sheet tab context menus can set tab color, hide the active sheet, and unhide via the tab strip background", async ({ page }) => {
    await gotoDesktop(page);

    // Keep A1 active so the status bar reflects deterministic values after sheet switches.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.activateCell({ row: 0, col: 0 });
    });

    // Lazily create Sheet2 by writing a value into it.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.getDocument().setCellValue("Sheet2", "A1", "Hello from Sheet2");
    });

    const sheet2Tab = page.getByTestId("sheet-tab-Sheet2");
    await expect(sheet2Tab).toBeVisible();
    await expect(page.getByTestId("sheet-position")).toHaveText("Sheet 1 of 2");

    // Tab Color palette sets underline visible.
    const tabMenu = page.getByTestId("sheet-tab-context-menu");
    await sheet2Tab.click();
    await expect(sheet2Tab).toHaveAttribute("data-active", "true");
    await expect(page.getByTestId("sheet-position")).toHaveText("Sheet 2 of 2");

    await sheet2Tab.focus();
    await page.keyboard.press("Shift+F10");
    await expect(tabMenu).toBeVisible();
    await tabMenu.getByRole("button", { name: "Tab Color" }).click();
    await tabMenu.getByRole("button", { name: "Blue" }).click();
    await expect(tabMenu).toBeHidden();
    await expect(sheet2Tab.locator(".sheet-tab__color")).toBeVisible();

    // Hiding the active sheet should activate an adjacent visible sheet (Sheet1).
    await sheet2Tab.focus();
    await page.keyboard.press("Shift+F10");
    await expect(tabMenu).toBeVisible();
    await tabMenu.getByRole("button", { name: "Hide", exact: true }).click();
    await expect(tabMenu).toBeHidden();

    await expect(sheet2Tab).toHaveCount(0);
    await expect(page.getByTestId("sheet-tab-Sheet1")).toHaveAttribute("data-active", "true");
    await expect(page.getByTestId("sheet-position")).toHaveText("Sheet 1 of 1");
    await expect(page.getByTestId("active-value")).toHaveText("Seed");

    // Unhide via background menu restores Sheet2.
    await page.evaluate(() => {
      const strip = document.querySelector<HTMLElement>("#sheet-tabs .sheet-tabs");
      if (!strip) throw new Error("Missing sheet tab strip");
      const rect = strip.getBoundingClientRect();
      strip.dispatchEvent(
        new MouseEvent("contextmenu", {
          bubbles: true,
          cancelable: true,
          clientX: rect.left + rect.width - 4,
          clientY: rect.top + rect.height / 2,
        }),
      );
    });

    await expect(tabMenu).toBeVisible();
    await tabMenu.getByRole("button", { name: "Unhide…" }).click();
    await page.getByTestId("quick-pick").getByRole("button", { name: "Sheet2" }).click();

    await expect(sheet2Tab).toBeVisible();
    await expect(page.getByTestId("sheet-position")).toHaveText("Sheet 1 of 2");

    // Tab color should persist after hide/unhide.
    await expect(sheet2Tab.locator(".sheet-tab__color")).toBeVisible();
  });
});
