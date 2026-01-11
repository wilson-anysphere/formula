import { expect, test } from "@playwright/test";

test.describe("comments", () => {
  test("add comment, reply, resolve", async ({ page }) => {
    await page.goto("/");
    await page.waitForFunction(() => (window as any).__formulaApp != null);

    // Focus + select A1 (top-left).
    await page.click("#grid", { position: { x: 5, y: 5 } });
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    await page.getByTestId("open-comments-panel").click();
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
});
