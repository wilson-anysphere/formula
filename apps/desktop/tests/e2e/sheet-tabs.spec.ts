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

  test("keyboard navigation activates the focused sheet tab", async ({ page }) => {
    await gotoDesktop(page);

    // Ensure A1 is active before switching sheets so the status bar reflects A1 values.
    // Click inside A1 (avoid the shared-grid corner header/select-all region).
    await page.click("#grid", { position: { x: 80, y: 40 } });

    // Lazily create Sheet2 by writing a value into it.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.getDocument().setCellValue("Sheet2", "A1", "Hello from Sheet2");
    });
    await expect(page.getByRole("tab", { name: "Sheet2" })).toBeVisible();

    // Focus the tab strip via keyboard navigation.
    const sheet1Tab = page.getByRole("tab", { name: "Sheet1" });
    for (let i = 0; i < 20; i += 1) {
      await page.keyboard.press("Tab");
      if (await sheet1Tab.evaluate((el) => el === document.activeElement)) break;
    }
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
    for (let i = 0; i < 20; i += 1) {
      await page.keyboard.press("Tab");
      if (await sheet1Tab.evaluate((el) => el === document.activeElement)) break;
    }
    await expect(sheet1Tab).toBeFocused();

    await page.keyboard.press("Shift+F10");
    const menu = page.getByTestId("sheet-tab-context-menu");
    await expect(menu).toBeVisible();
    await expect(menu.getByRole("button", { name: "Rename" })).toBeVisible();

    await page.keyboard.press("Escape");
    await expect(menu).toBeHidden();
    await expect(sheet1Tab).toBeFocused();
  });

  test("double-click rename commits on Enter and updates the tab label", async ({ page }) => {
    await gotoDesktop(page);

    const tab = page.getByTestId("sheet-tab-Sheet1");
    await expect(tab).toBeVisible();

    await tab.dblclick();
    const input = tab.locator("input.sheet-tab__input");
    await expect(input).toBeVisible();

    await input.fill("Renamed");
    await input.press("Enter");

    await expect(tab).toContainText("Renamed");
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
        typeof app.getWorkbookSheetStore === "function" ? app.getWorkbookSheetStore() : (window as any).__workbookSheetStore;

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

  test("invalid rename keeps editing and blocks switching sheets", async ({ page }) => {
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

    // Trigger an invalid name (empty) and attempt to commit.
    await input.fill("");
    await input.press("Enter");

    // Attempt to switch sheets; invalid rename should block activation.
    await page.getByTestId("sheet-tab-Sheet2").click();

    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCurrentSheetId())).toBe("Sheet1");
    await expect(sheet1Tab).toHaveAttribute("data-active", "true");
    await expect(input).toBeVisible();
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

    // Move Sheet2 after Sheet3 (new order: Sheet1, Sheet3, Sheet2).
    try {
      await page
        .getByTestId("sheet-tab-Sheet2")
        // targetPosition is relative to the Sheet3 tab; use a large X so it lands on the "after" side.
        .dragTo(page.getByTestId("sheet-tab-Sheet3"), { targetPosition: { x: 999, y: 1 } });
    } catch {
      // Ignore; we'll fall back to a synthetic drop below.
    }

    // Playwright drag/drop can be flaky with HTML5 DataTransfer. If the order doesn't match,
    // dispatch a synthetic drop event that exercises the sheet tab DnD plumbing.
    const desiredOrder = ["Sheet1", "Sheet3", "Sheet2"];
    const orderKey = (order: Array<string | null>) => order.filter(Boolean).slice(0, 3).join(",");
    const didReorder = orderKey(
      await page.evaluate(() =>
        Array.from(document.querySelectorAll("#sheet-tabs .sheet-tabs [data-sheet-id]")).map((el) =>
          (el as HTMLElement).getAttribute("data-sheet-id"),
        ),
      ),
    );

    if (didReorder !== desiredOrder.join(",")) {
      await page.evaluate(() => {
        const fromId = "Sheet2";
        const target = document.querySelector('[data-testid="sheet-tab-Sheet3"]') as HTMLElement | null;
        if (!target) throw new Error("Missing Sheet3 tab");
        const rect = target.getBoundingClientRect();

        const dt = new DataTransfer();
        dt.setData("text/sheet-id", fromId);
        dt.setData("text/plain", fromId);

        const drop = new DragEvent("drop", {
          bubbles: true,
          cancelable: true,
          clientX: rect.left + rect.width - 1,
          clientY: rect.top + rect.height / 2,
        });
        Object.defineProperty(drop, "dataTransfer", { value: dt });
        target.dispatchEvent(drop);
      });
    }

    await expect.poll(() =>
      page.evaluate(() =>
        Array.from(document.querySelectorAll("#sheet-tabs .sheet-tabs [data-sheet-id]")).map((el) =>
          (el as HTMLElement).getAttribute("data-sheet-id"),
        ),
      ),
    ).toEqual(desiredOrder);
    // Active sheet is still Sheet3, but its position is now 2nd.
    await expect(page.getByTestId("sheet-position")).toHaveText("Sheet 2 of 3");

    // Ctrl+PgDn should follow the new order. We dispatch the key event directly to avoid
    // platform/browser-specific tab switching behavior.
    await page.evaluate(() => {
      const evt = new KeyboardEvent("keydown", { key: "PageDown", ctrlKey: true, bubbles: true, cancelable: true });
      document.getElementById("grid")?.dispatchEvent(evt);
    });
    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCurrentSheetId())).toBe("Sheet2");
    await expect(page.getByTestId("sheet-position")).toHaveText("Sheet 3 of 3");

    await page.evaluate(() => {
      const evt = new KeyboardEvent("keydown", { key: "PageDown", ctrlKey: true, bubbles: true, cancelable: true });
      document.getElementById("grid")?.dispatchEvent(evt);
    });
    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCurrentSheetId())).toBe("Sheet1");
    await expect(page.getByTestId("sheet-position")).toHaveText("Sheet 1 of 3");

    await page.evaluate(() => {
      const evt = new KeyboardEvent("keydown", { key: "PageDown", ctrlKey: true, bubbles: true, cancelable: true });
      document.getElementById("grid")?.dispatchEvent(evt);
    });
    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCurrentSheetId())).toBe("Sheet3");
    await expect(page.getByTestId("sheet-position")).toHaveText("Sheet 2 of 3");
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
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.getWorkbookSheetStore().hide("Sheet2");
    });
    await expect(page.locator('[data-testid="sheet-tab-Sheet2"]')).toHaveCount(0);

    // Drag Sheet3 before Sheet1.
    try {
      await page
        .getByTestId("sheet-tab-Sheet3")
        .dragTo(page.getByTestId("sheet-tab-Sheet1"), { targetPosition: { x: 1, y: 1 } });
    } catch {
      // Ignore; we'll fall back to a synthetic drop below.
    }

    // As with other tab drag tests, use a synthetic drop if Playwright's dragTo doesn't take.
    const desiredAll = [
      { id: "Sheet3", visibility: "visible" },
      { id: "Sheet1", visibility: "visible" },
      { id: "Sheet2", visibility: "hidden" },
    ];

    const currentAll = await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      return app.getWorkbookSheetStore().listAll().map((s: any) => ({ id: s.id, visibility: s.visibility }));
    });

    if (JSON.stringify(currentAll) !== JSON.stringify(desiredAll)) {
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
    }

    await expect
      .poll(() =>
        page.evaluate(() => {
          const app = (window as any).__formulaApp;
          return app.getWorkbookSheetStore().listAll().map((s: any) => ({ id: s.id, visibility: s.visibility }));
        }),
      )
      .toEqual(desiredAll);

    // Unhide Sheet2 and ensure it lands at the end (i.e. hidden sheet moved as expected).
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.getWorkbookSheetStore().unhide("Sheet2");
    });
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

    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.getWorkbookSheetStore().hide("Sheet2");
    });

    await expect(page.getByTestId("sheet-position")).toHaveText("Sheet 2 of 2");

    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.getWorkbookSheetStore().unhide("Sheet2");
    });

    await expect(page.getByTestId("sheet-position")).toHaveText("Sheet 3 of 3");
  });

  test("renaming a sheet marks the document dirty", async ({ page }) => {
    await gotoDesktop(page);

    // Demo workbook is treated as an initial saved baseline.
    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getDocument().isDirty)).toBe(false);

    const tab = page.getByTestId("sheet-tab-Sheet1");
    await tab.dblclick();
    const input = tab.locator("input");
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
    const renameInput = sheet1Tab.locator("input");
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
      await (window as any).__formulaCommandRegistry.executeCommand("workbench.showCommandPalette");
    });
    await expect(page.getByTestId("command-palette-input")).toBeVisible();

    await page.getByTestId("command-palette-input").fill("RenamedSheet1!A1");
    await expect(page.locator('[data-testid="command-palette-list"] .command-palette__item').first()).toContainText(
      "Go to RenamedSheet1!A1",
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
});
