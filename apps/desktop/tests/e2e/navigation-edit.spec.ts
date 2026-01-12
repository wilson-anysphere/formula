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

test.describe("grid keyboard navigation + in-place editing", () => {
  test("arrow navigation updates active cell", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);

    // Focus + select A1 (top-left).
    await page.click("#grid", { position: { x: 80, y: 40 } });

    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    await page.keyboard.press("ArrowRight");
    await page.keyboard.press("ArrowRight");
    await expect(page.getByTestId("active-cell")).toHaveText("C1");

    await page.keyboard.press("ArrowDown");
    await expect(page.getByTestId("active-cell")).toHaveText("C2");
  });

  test("Tab/Shift+Tab navigate cells when the grid is focused", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);
    await page.click("#grid", { position: { x: 80, y: 40 } });

    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    await page.keyboard.press("Tab");
    await expect(page.getByTestId("active-cell")).toHaveText("B1");

    await page.keyboard.press("Shift+Tab");
    await expect(page.getByTestId("active-cell")).toHaveText("A1");
  });

  test("Tab wraps within a multi-cell selection (Excel-style)", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);

    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const sheetId = app.getCurrentSheetId();
      app.selectRange({ sheetId, range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 } });
    });

    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    await page.keyboard.press("Tab");
    await expect(page.getByTestId("active-cell")).toHaveText("B1");

    await page.keyboard.press("Tab");
    await expect(page.getByTestId("active-cell")).toHaveText("A2");

    await page.keyboard.press("Tab");
    await expect(page.getByTestId("active-cell")).toHaveText("B2");

    // Wrap back to the start of the selection.
    await page.keyboard.press("Tab");
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    // Shift+Tab wraps backwards.
    await page.keyboard.press("Shift+Tab");
    await expect(page.getByTestId("active-cell")).toHaveText("B2");
  });

  test("typing begins in-place edit and commits", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);
    await page.click("#grid", { position: { x: 80, y: 40 } });

    const recalcBefore = await page.evaluate(() => (window as any).__formulaApp.getRecalcCount());

    // Type directly without pressing F2 first.
    await page.keyboard.press("h");
    const editor = page.locator("textarea.cell-editor");
    await expect(editor).toBeVisible();

    await page.keyboard.type("ello");
    await page.keyboard.press("Enter");
    await waitForIdle(page);

    await expect(page.getByTestId("active-cell")).toHaveText("A2");
    const a1Value = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"));
    expect(a1Value).toBe("hello");

    const recalcAfter = await page.evaluate(() => (window as any).__formulaApp.getRecalcCount());
    expect(recalcAfter).toBeGreaterThan(recalcBefore);
  });

  test("Ctrl+End jumps to last used cell", async ({ page, browserName }) => {
    // WebKit on Linux can be flaky with End/Home in headless; keep the coverage
    // focused on chromium/firefox in CI.
    test.skip(browserName === "webkit", "End/Home key handling is unreliable in webkit headless");

    await gotoDesktop(page);
    await waitForIdle(page);
    await page.click("#grid", { position: { x: 80, y: 40 } });

    await page.keyboard.press("Control+End");
    await expect(page.getByTestId("active-cell")).toHaveText("D5");
  });

  test("PageDown scrolls the grid and keeps the active cell visible", async ({ page, browserName }) => {
    test.skip(browserName === "webkit", "PageDown key handling can be unreliable in webkit headless");
    await gotoDesktop(page);
    await waitForIdle(page);
    const grid = page.locator("#grid");
    await grid.click({ position: { x: 60, y: 40 } });

    const scrollBefore = await page.evaluate(() => (window as any).__formulaApp.getScroll().y);
    const activeBefore = await page.evaluate(() => (window as any).__formulaApp.getActiveCell());

    await page.keyboard.press("PageDown");

    const scrollAfter = await page.evaluate(() => (window as any).__formulaApp.getScroll().y);
    const activeAfter = await page.evaluate(() => (window as any).__formulaApp.getActiveCell());

    expect(scrollAfter).toBeGreaterThan(scrollBefore);
    expect(activeAfter.row).toBeGreaterThan(activeBefore.row);
  });

  test("F2 edit commits with Enter and cancels with Escape", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);
    await page.click("#grid", { position: { x: 80, y: 40 } });

    const recalcBefore = await page.evaluate(() => (window as any).__formulaApp.getRecalcCount());

    // Edit A1.
    await page.keyboard.press("F2");
    const editor = page.locator("textarea.cell-editor");
    await expect(editor).toBeVisible();

    await editor.fill("Hello");
    await page.keyboard.press("Enter");
    await waitForIdle(page);

    // Commit moves down.
    await expect(page.getByTestId("active-cell")).toHaveText("A2");

    const recalcAfterCommit = await page.evaluate(() => (window as any).__formulaApp.getRecalcCount());
    expect(recalcAfterCommit).toBeGreaterThan(recalcBefore);

    const a1Value = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"));
    expect(a1Value).toBe("Hello");

    // Start editing A2 but cancel.
    const a2Before = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A2"));
    await page.keyboard.press("F2");
    await expect(editor).toBeVisible();
    await editor.fill("ShouldNotCommit");
    await page.keyboard.press("Escape");

    const recalcAfterCancel = await page.evaluate(() => (window as any).__formulaApp.getRecalcCount());
    expect(recalcAfterCancel).toBe(recalcAfterCommit);

    const a2Value = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A2"));
    expect(a2Value).toBe(a2Before);
    await expect(page.getByTestId("active-cell")).toHaveText("A2");

    // Delete clears cell contents.
    await page.keyboard.press("ArrowUp");
    await expect(page.getByTestId("active-cell")).toHaveText("A1");
    const recalcBeforeDelete = await page.evaluate(() => (window as any).__formulaApp.getRecalcCount());
    await page.keyboard.press("Delete");
    await waitForIdle(page);
    const a1Cleared = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"));
    expect(a1Cleared).toBe("");
    const recalcAfterDelete = await page.evaluate(() => (window as any).__formulaApp.getRecalcCount());
    expect(recalcAfterDelete).toBeGreaterThan(recalcBeforeDelete);
  });

  test("Ctrl/Cmd+Z undo + Ctrl/Cmd+Shift+Z / Ctrl/Cmd+Y redo update cells", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);
    await page.click("#grid", { position: { x: 80, y: 40 } });

    // Clear any seeded value so undo returns to an empty cell.
    await page.keyboard.press("Delete");
    await waitForIdle(page);

    // Create a dependent formula in B1 so we can verify recomputation after undo/redo.
    await page.keyboard.press("ArrowRight");
    await expect(page.getByTestId("active-cell")).toHaveText("B1");
    await page.keyboard.press("F2");
    const editor = page.locator("textarea.cell-editor");
    await expect(editor).toBeVisible();
    await editor.fill("=A1");
    await page.keyboard.press("Enter"); // commit, moves to B2
    await waitForIdle(page);

    // Edit A1.
    await page.keyboard.press("ArrowUp"); // back to B1
    await page.keyboard.press("ArrowLeft"); // A1
    await page.keyboard.press("F2");
    await expect(editor).toBeVisible();
    await editor.fill("Hello");
    await page.keyboard.press("Enter");
    await waitForIdle(page);

    const modifier = process.platform === "darwin" ? "Meta" : "Control";

    await page.keyboard.press(`${modifier}+Z`);
    await waitForIdle(page);
    const a1AfterUndo = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"));
    expect(a1AfterUndo).toBe("");
    const b1AfterUndo = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("B1"));
    expect(b1AfterUndo).toBe("");

    await page.keyboard.press(`${modifier}+Shift+Z`);
    await waitForIdle(page);
    const a1AfterRedo = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"));
    expect(a1AfterRedo).toBe("Hello");
    const b1AfterRedo = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("B1"));
    expect(b1AfterRedo).toBe("Hello");

    // Redo should also work via Ctrl/Cmd+Y.
    await page.keyboard.press(`${modifier}+Z`);
    await waitForIdle(page);
    const a1AfterUndo2 = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"));
    expect(a1AfterUndo2).toBe("");
    const b1AfterUndo2 = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("B1"));
    expect(b1AfterUndo2).toBe("");

    await page.keyboard.press(`${modifier}+Y`);
    await waitForIdle(page);
    const a1AfterRedoY = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"));
    expect(a1AfterRedoY).toBe("Hello");
    const b1AfterRedoY = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("B1"));
    expect(b1AfterRedoY).toBe("Hello");

    // Undo/redo should still work even if the grid is not focused (e.g. after clicking a toolbar button).
    await page.getByTestId("ribbon-root").getByTestId("split-vertical").click();
    await page.keyboard.press(`${modifier}+Z`);
    await waitForIdle(page);
    const a1AfterUndoFromToolbar = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"));
    expect(a1AfterUndoFromToolbar).toBe("");
    const b1AfterUndoFromToolbar = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("B1"));
    expect(b1AfterUndoFromToolbar).toBe("");

    await page.keyboard.press(`${modifier}+Shift+Z`);
    await waitForIdle(page);
    const a1AfterRedoFromToolbar = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"));
    expect(a1AfterRedoFromToolbar).toBe("Hello");
    const b1AfterRedoFromToolbar = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("B1"));
    expect(b1AfterRedoFromToolbar).toBe("Hello");
  });

  test("Ctrl/Cmd+Z does not steal undo while editing a cell", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);
    await page.click("#grid", { position: { x: 80, y: 40 } });

    // Clear any seeded value.
    await page.keyboard.press("Delete");
    await waitForIdle(page);

    // Commit A1 = Hello.
    await page.keyboard.press("F2");
    const editor = page.locator("textarea.cell-editor");
    await expect(editor).toBeVisible();
    await page.keyboard.type("Hello");
    await page.keyboard.press("Enter");
    await waitForIdle(page);

    await page.keyboard.press("ArrowUp"); // back to A1

    // Start editing and append a character.
    await page.keyboard.press("F2");
    await expect(editor).toBeVisible();
    await page.keyboard.type("X");
    await expect(editor).toHaveValue("HelloX");

    const modifier = process.platform === "darwin" ? "Meta" : "Control";
    await page.keyboard.press(`${modifier}+Z`);

    // Should undo inside the textarea (native undo), not the spreadsheet history.
    await expect(editor).toHaveValue("Hello");
    const a1AfterEditorUndo = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"));
    expect(a1AfterEditorUndo).toBe("Hello");

    await page.keyboard.press("Escape");
    await waitForIdle(page);
    const a1Final = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"));
    expect(a1Final).toBe("Hello");
  });

  test("Ctrl/Cmd+Z does not steal undo while editing the formula bar", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);
    await page.click("#grid", { position: { x: 80, y: 40 } });

    // Clear any seeded value.
    await page.keyboard.press("Delete");
    await waitForIdle(page);

    // Commit A1 = Hello.
    await page.keyboard.press("F2");
    const cellEditor = page.locator("textarea.cell-editor");
    await expect(cellEditor).toBeVisible();
    await cellEditor.fill("Hello");
    await page.keyboard.press("Enter");
    await waitForIdle(page);

    // Move back to A1 and start editing in the formula bar.
    await page.keyboard.press("ArrowUp");
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    await page.getByTestId("formula-highlight").click();
    const input = page.getByTestId("formula-input");
    await expect(input).toBeVisible();
    await expect(input).toHaveValue("Hello");

    // Create a text edit inside the formula bar.
    await page.keyboard.type("X");
    await expect(input).toHaveValue("HelloX");

    const modifier = process.platform === "darwin" ? "Meta" : "Control";
    await page.keyboard.press(`${modifier}+Z`);

    // Should undo inside the textarea (native undo), not the spreadsheet history.
    await expect(input).toHaveValue("Hello");
    const a1AfterFormulaUndo = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"));
    expect(a1AfterFormulaUndo).toBe("Hello");
  });
});
