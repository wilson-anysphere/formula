import { expect, test } from "@playwright/test";

import { expectSheetPosition, gotoDesktop } from "./helpers";

test.describe("sheet navigation shortcuts", () => {
  test("Ctrl+PageDown / Ctrl+PageUp switches the active sheet (wraps)", async ({ page }) => {
    await gotoDesktop(page);
    await expect(page.getByTestId("sheet-tab-Sheet1")).toBeVisible();
    await expectSheetPosition(page, { position: 1, total: 1 });

    // Ensure the grid has focus by clicking the center of A1 once the layout is ready.
    await page.waitForFunction(() => {
      const app = (window as any).__formulaApp;
      const rect = app?.getCellRectA1?.("A1");
      return rect && rect.width > 0 && rect.height > 0;
    });
    const a1 = (await page.evaluate(() => (window as any).__formulaApp.getCellRectA1("A1"))) as {
      x: number;
      y: number;
      width: number;
      height: number;
    };
    await page
      .locator("#grid")
      .click({ position: { x: a1.x + a1.width / 2, y: a1.y + a1.height / 2 } });
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    // Lazily create Sheet2 by writing a value into it.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.getDocument().setCellValue("Sheet2", "A1", "Hello from Sheet2");
    });
    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();
    await expectSheetPosition(page, { position: 1, total: 2 });

    const ctrlKey = process.platform !== "darwin";
    const metaKey = process.platform === "darwin";

    const dispatch = async (key: "PageUp" | "PageDown") => {
      await page.evaluate(
        ({ key, ctrlKey, metaKey }) => {
          const grid = document.getElementById("grid");
          if (!grid) throw new Error("Missing #grid");
          grid.focus();
          const evt = new KeyboardEvent("keydown", { key, ctrlKey, metaKey, bubbles: true, cancelable: true });
          grid.dispatchEvent(evt);
        },
        { key, ctrlKey, metaKey },
      );
    };

    // Next sheet.
    await dispatch("PageDown");
    await expect(page.getByTestId("sheet-tab-Sheet2")).toHaveAttribute("data-active", "true");
    await expect(page.getByTestId("active-cell")).toHaveText("A1");
    await expect(page.getByTestId("active-value")).toHaveText("Hello from Sheet2");
    await expectSheetPosition(page, { position: 2, total: 2 });

    // Previous sheet.
    await dispatch("PageUp");
    await expect(page.getByTestId("sheet-tab-Sheet1")).toHaveAttribute("data-active", "true");
    await expect(page.getByTestId("active-cell")).toHaveText("A1");
    await expect(page.getByTestId("active-value")).toHaveText("Seed");
    await expectSheetPosition(page, { position: 1, total: 2 });

    // Wrap-around at the start.
    await dispatch("PageUp");
    await expect(page.getByTestId("sheet-tab-Sheet2")).toHaveAttribute("data-active", "true");
    await expectSheetPosition(page, { position: 2, total: 2 });

    // Wrap-around at the end.
    await dispatch("PageDown");
    await expect(page.getByTestId("sheet-tab-Sheet1")).toHaveAttribute("data-active", "true");
    await expectSheetPosition(page, { position: 1, total: 2 });
  });

  test("Ctrl/Cmd+PageDown keeps focus in the sheet tab strip when invoked from a focused tab", async ({ page }) => {
    await gotoDesktop(page);
    await expect(page.getByTestId("sheet-tab-Sheet1")).toBeVisible();

    // Create Sheet2 so sheet navigation is observable.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.getDocument().setCellValue("Sheet2", "A1", "Hello from Sheet2");
    });
    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();
    await expect(page.getByTestId("sheet-position")).toHaveText("Sheet 1 of 2");

    const sheet1Tab = page.getByTestId("sheet-tab-Sheet1");
    await sheet1Tab.focus();
    await expect(sheet1Tab).toBeFocused();

    // Dispatch directly so we don't trigger browser tab switching.
    await page.evaluate((isMac) => {
      const target = document.activeElement ?? window;
      target.dispatchEvent(
        new KeyboardEvent("keydown", {
          key: "PageDown",
          metaKey: isMac,
          ctrlKey: !isMac,
          bubbles: true,
          cancelable: true,
        }),
      );
    }, process.platform === "darwin");

    const sheet2Tab = page.getByTestId("sheet-tab-Sheet2");
    await expect(sheet2Tab).toHaveAttribute("data-active", "true");
    await expect(sheet2Tab).toBeFocused();
    await expect
      .poll(() => page.evaluate(() => Boolean((document.activeElement as HTMLElement | null)?.closest("#sheet-tabs"))))
      .toBe(true);
  });

  test("Ctrl/Cmd+PageUp/PageDown works while the formula bar is editing a formula (keeps focus)", async ({ page }) => {
    await gotoDesktop(page);
    await expect(page.getByTestId("sheet-tab-Sheet1")).toBeVisible();

    // Lazily create Sheet2 so sheet navigation is observable.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.getDocument().setCellValue("Sheet2", "A1", "Hello from Sheet2");
    });
    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();
    await expectSheetPosition(page, { position: 1, total: 2 });

    // In view mode the formula bar input is hidden and click events are handled by the
    // highlighted <pre> below, so click the highlight to enter edit mode and focus the textarea.
    await page.getByTestId("formula-highlight").click();
    const formulaInput = page.getByTestId("formula-input");
    await expect(formulaInput).toBeVisible();
    await expect(formulaInput).toBeFocused();
    await formulaInput.fill("=");
    await expect(formulaInput).toHaveValue("=");

    // `fill()` updates the DOM value immediately, but the formula bar model updates via an `input`
    // event handler. Wait until the app reports it is in formula-editing mode so the global
    // Ctrl/Cmd+PgUp/PgDn handler can treat this as an Excel-like cross-sheet reference session.
    await expect
      .poll(() =>
        page.evaluate(() => {
          const app = (window as any).__formulaApp;
          return Boolean(app?.isFormulaBarFormulaEditing?.());
        }),
      )
      .toBe(true);

    // While editing a formula, Ctrl/Cmd+PgDn should still switch sheets (Excel-like cross-sheet reference building).
    await page.evaluate((isMac) => {
      const input = document.querySelector('[data-testid="formula-input"]') as HTMLTextAreaElement | null;
      if (!input) throw new Error("Missing formula input");
      input.dispatchEvent(
        new KeyboardEvent("keydown", {
          key: "PageDown",
          metaKey: isMac,
          ctrlKey: !isMac,
          bubbles: true,
          cancelable: true,
        }),
      );
    }, process.platform === "darwin");

    await expect(page.getByTestId("sheet-tab-Sheet2")).toHaveAttribute("data-active", "true");
    await expectSheetPosition(page, { position: 2, total: 2 });

    // Sheet switching should not interrupt formula editing; focus remains in the formula bar.
    await expect(formulaInput).toBeFocused();
    await expect(formulaInput).toHaveValue("=");
  });

  test("Ctrl/Cmd+PageUp/PageDown is global (works from ribbon focus) and restores grid focus", async ({ page }) => {
    await gotoDesktop(page);
    await expect(page.getByTestId("sheet-tab-Sheet1")).toBeVisible();
    await expectSheetPosition(page, { position: 1, total: 1 });

    // Lazily create Sheet2 so sheet navigation is observable.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.getDocument().setCellValue("Sheet2", "A1", "Hello from Sheet2");
    });
    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();
    await expectSheetPosition(page, { position: 1, total: 2 });

    // Ensure we start on Sheet1.
    await page.getByTestId("sheet-tab-Sheet1").click();
    await expect(page.getByTestId("sheet-tab-Sheet1")).toHaveAttribute("data-active", "true");

    // Focus a non-input UI surface (the ribbon) so Ctrl/Cmd+PgDn must be handled globally.
    const ribbon = page.getByTestId("ribbon-root");
    await expect(ribbon).toBeVisible();

    const homeTab = ribbon.getByRole("tab", { name: "Home" });
    await homeTab.click();
    await expect(homeTab).toHaveAttribute("aria-selected", "true");

    const bold = ribbon.locator('button[data-command-id="home.font.bold"]');
    await expect(bold).toBeVisible();
    await bold.focus();
    await expect(bold).toBeFocused();

    // Avoid Playwright's keyboard.press to sidestep browser tab-switch shortcuts; dispatching still
    // exercises our global handler (capture phase) and focus scoping rules.
    await page.evaluate((isMac) => {
      const target = document.activeElement ?? window;
      target.dispatchEvent(
        new KeyboardEvent("keydown", {
          key: "PageDown",
          metaKey: isMac,
          ctrlKey: !isMac,
          bubbles: true,
          cancelable: true,
        }),
      );
    }, process.platform === "darwin");

    await expect(page.getByTestId("sheet-tab-Sheet2")).toHaveAttribute("data-active", "true");
    await expectSheetPosition(page, { position: 2, total: 2 });

    // After switching sheets, focus should return to the grid so keyboard workflows keep working.
    await expect
      .poll(() => page.evaluate(() => (document.activeElement as HTMLElement | null)?.id))
      .toBe("grid");

    // Repeat in the opposite direction (Sheet2 -> Sheet1) while focus is outside the grid.
    await bold.focus();
    await expect(bold).toBeFocused();

    await page.evaluate((isMac) => {
      const target = document.activeElement ?? window;
      target.dispatchEvent(
        new KeyboardEvent("keydown", {
          key: "PageUp",
          metaKey: isMac,
          ctrlKey: !isMac,
          bubbles: true,
          cancelable: true,
        }),
      );
    }, process.platform === "darwin");

    await expect(page.getByTestId("sheet-tab-Sheet1")).toHaveAttribute("data-active", "true");
    await expectSheetPosition(page, { position: 1, total: 2 });
    await expect
      .poll(() => page.evaluate(() => (document.activeElement as HTMLElement | null)?.id))
      .toBe("grid");
  });
});
