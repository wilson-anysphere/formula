import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("comments", () => {
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

    await grid.click({ button: "right", position: { x: a1.x + a1.width / 2, y: a1.y + a1.height / 2 } });
    const menu = page.getByTestId("context-menu");
    await expect(menu).toBeVisible();

    await menu.getByRole("button", { name: "Add Comment" }).click();

    await expect(page.getByTestId("comments-panel")).toBeVisible();
    await expect(page.getByTestId("new-comment-input")).toBeFocused();
  });
});
