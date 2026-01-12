import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("comments", () => {
  test("clicking the comments panel input commits an in-progress in-cell edit without stealing focus", async ({ page }) => {
    await gotoDesktop(page);

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

    const grid = page.locator("#grid");
    await grid.click({ position: { x: a1.x + a1.width / 2, y: a1.y + a1.height / 2 } });
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    // Open the comments panel (this focuses the comment input by design).
    await page.getByTestId("ribbon-root").getByTestId("open-comments-panel").click();
    const panel = page.getByTestId("comments-panel");
    const newCommentInput = panel.getByTestId("new-comment-input");
    await expect(panel).toBeVisible();
    await expect(newCommentInput).toBeFocused();

    // Return focus to the grid and start an in-cell edit, but do not press Enter.
    await grid.click({ position: { x: a1.x + a1.width / 2, y: a1.y + a1.height / 2 } });
    await page.keyboard.press("h");
    const editor = page.locator("textarea.cell-editor");
    await expect(editor).toBeVisible();
    await page.keyboard.type("ello");
    await expect(editor).toHaveValue("hello");

    // Clicking the comments input should commit the edit but leave focus on the input
    // (no focus ping-pong back to the grid).
    await newCommentInput.click();
    await expect(newCommentInput).toBeFocused();
    await expect(editor).toBeHidden();

    await page.evaluate(() => (window as any).__formulaApp.whenIdle());
    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"))).toBe("hello");
    await expect(newCommentInput).toBeFocused();
  });

  test("add comment, reply, resolve", async ({ page }) => {
    await gotoDesktop(page);

    // Focus + select A1. Avoid clicking the header corner (which can select the
    // whole sheet) by using the runtime's computed cell rect.
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

    const grid = page.locator("#grid");
    await grid.click({ position: { x: a1.x + a1.width / 2, y: a1.y + a1.height / 2 } });
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    // The "open comments panel" toggle is provided by the ribbon UI. Scope this locator to the
    // ribbon so it stays stable even if other parts of the UI introduce additional toggles.
    await page.getByTestId("ribbon-root").getByTestId("open-comments-panel").click();
    const panel = page.getByTestId("comments-panel");
    await expect(panel).toBeVisible();

    await panel.getByTestId("new-comment-input").fill("Hello from e2e");
    await panel.getByTestId("submit-comment").click();

    const thread = panel.getByTestId("comment-thread").first();
    await expect(thread).toContainText("Hello from e2e");
    await expect(thread).toHaveAttribute("data-resolved", "false");

    await thread.getByTestId("reply-input").fill("Reply from e2e");
    await thread.getByTestId("submit-reply").click();
    await expect(thread).toContainText("Reply from e2e");

    await thread.getByTestId("resolve-comment").click();
    await expect(thread).toHaveAttribute("data-resolved", "true");
  });

  test("cell context menu > Add Comment opens the panel and focuses the input", async ({ page }) => {
    await gotoDesktop(page);

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

    const grid = page.locator("#grid");
    await grid.click({ position: { x: a1.x + a1.width / 2, y: a1.y + a1.height / 2 } });
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
      { x: a1.x + a1.width / 2, y: a1.y + a1.height / 2 },
    );
    const menu = page.getByTestId("context-menu");
    await expect(menu).toBeVisible();

    const expectedShortcut = process.platform === "darwin" ? "⇧F2" : "Shift+F2";
    const addComment = menu.getByRole("button", { name: "Add Comment" });
    await expect(addComment.locator('span[aria-hidden="true"]')).toHaveText(expectedShortcut);

    await addComment.click();

    await expect(page.getByTestId("comments-panel")).toBeVisible();
    await expect(page.getByTestId("new-comment-input")).toBeFocused();
  });

  test("opening the grid context menu commits an in-progress in-cell edit", async ({ page }) => {
    await gotoDesktop(page);

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

    // Ensure the cell starts empty so commit assertions are deterministic.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      doc.setCellValue(sheetId, "A1", null, { label: "Clear A1" });
    });

    const grid = page.locator("#grid");
    await grid.click({ position: { x: a1.x + a1.width / 2, y: a1.y + a1.height / 2 } });
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    await page.keyboard.press("F2");
    const editor = page.getByTestId("cell-editor");
    await expect(editor).toBeVisible();
    await page.keyboard.type("hello");
    await expect(editor).toHaveValue("hello");

    // Open the context menu while editing. This should commit the edit via the editor blur handler.
    // (Dispatch the event directly to avoid flaky right-click handling.)
    await page.evaluate(
      ({ x, y }) => {
        const gridEl = document.getElementById("grid");
        if (!gridEl) throw new Error("Missing #grid");
        const rect = gridEl.getBoundingClientRect();
        gridEl.dispatchEvent(
          new MouseEvent("contextmenu", {
            bubbles: true,
            cancelable: true,
            button: 2,
            clientX: rect.left + x,
            clientY: rect.top + y,
          }),
        );
      },
      { x: a1.x + a1.width / 2, y: a1.y + a1.height / 2 },
    );

    const menu = page.getByTestId("context-menu");
    await expect(menu).toBeVisible();
    await expect(editor).toBeHidden();

    await page.evaluate(() => (window as any).__formulaApp.whenIdle());
    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"))).toBe("hello");

    // Add Comment should remain available after the edit is committed.
    const addComment = menu.getByRole("button", { name: "Add Comment" });
    await expect(addComment).toBeEnabled();
  });

  test("comments shortcuts do not interrupt in-cell editing", async ({ page }) => {
    await gotoDesktop(page);

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

    const grid = page.locator("#grid");
    await grid.click({ position: { x: a1.x + a1.width / 2, y: a1.y + a1.height / 2 } });
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    await page.keyboard.press("F2");
    const editor = page.getByTestId("cell-editor");
    await expect(editor).toBeVisible();
    await expect(editor).toBeFocused();

    const panel = page.getByTestId("comments-panel");
    await expect(panel).not.toBeVisible();

    // Shift+F2 should not open comments while the inline editor is active.
    await page.keyboard.press("Shift+F2");
    await expect(panel).not.toBeVisible();
    await expect(editor).toBeVisible();

    // Ctrl/Cmd+Shift+M should also not toggle comments while editing.
    const modifier = process.platform === "darwin" ? "Meta" : "Control";
    await page.keyboard.press(`${modifier}+Shift+M`);
    await expect(panel).not.toBeVisible();
    await expect(editor).toBeVisible();
  });

  test("Shift+F2 opens the comments panel and focuses the input", async ({ page }) => {
    await gotoDesktop(page);

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

    const grid = page.locator("#grid");
    await grid.click({ position: { x: a1.x + a1.width / 2, y: a1.y + a1.height / 2 } });
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    await expect(page.getByTestId("comments-panel")).not.toBeVisible();

    await page.keyboard.press("Shift+F2");

    await expect(page.getByTestId("comments-panel")).toBeVisible();
    await expect(page.getByTestId("new-comment-input")).toBeFocused();
  });

  test("Ctrl/Cmd+Shift+M toggles the comments panel and returns focus to the grid on close", async ({ page }) => {
    await gotoDesktop(page);

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

    const grid = page.locator("#grid");
    await grid.click({ position: { x: a1.x + a1.width / 2, y: a1.y + a1.height / 2 } });
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    const modifier = process.platform === "darwin" ? "Meta" : "Control";

    await page.keyboard.press(`${modifier}+Shift+M`);
    await expect(page.getByTestId("comments-panel")).toBeVisible();

    // The comments panel input intentionally stops keydown propagation (and the global
    // keybinding handler ignores INPUT/TEXTAREA targets). To exercise the "toggle off"
    // behavior, return focus to the grid first.
    await grid.click({ position: { x: a1.x + a1.width / 2, y: a1.y + a1.height / 2 } });
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    await page.keyboard.press(`${modifier}+Shift+M`);
    await expect(page.getByTestId("comments-panel")).not.toBeVisible();

    // Ensure focus has returned to the grid (arrow keys should navigate selection).
    await page.keyboard.press("ArrowRight");
    await expect(page.getByTestId("active-cell")).toHaveText("B1");
  });

  test("Ctrl+Cmd+Shift+M toggles the comments panel (fallback chord)", async ({ page }) => {
    await gotoDesktop(page);

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

    const grid = page.locator("#grid");
    await grid.click({ position: { x: a1.x + a1.width / 2, y: a1.y + a1.height / 2 } });
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    const dispatch = async () => {
      await page.evaluate(() => {
        const el = document.getElementById("grid");
        if (!el) throw new Error("Missing #grid");
        el.dispatchEvent(
          new KeyboardEvent("keydown", {
            key: "M",
            code: "KeyM",
            ctrlKey: true,
            metaKey: true,
            shiftKey: true,
            bubbles: true,
            cancelable: true,
          }),
        );
      });
    };

    await dispatch();
    await expect(page.getByTestId("comments-panel")).toBeVisible();
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    await dispatch();
    await expect(page.getByTestId("comments-panel")).not.toBeVisible();
    await expect(page.getByTestId("active-cell")).toHaveText("A1");
  });
});
