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

    await menu.getByRole("button", { name: "Add Comment" }).click();

    await expect(page.getByTestId("comments-panel")).toBeVisible();
    await expect(page.getByTestId("new-comment-input")).toBeFocused();
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
});
