import { expect, test } from "@playwright/test";

test.describe("comments", () => {
  test("add comment, reply, resolve", async ({ page }) => {
    await page.goto("/");

    // Select cell A1.
    await page.click('[data-cell=\"A1\"]');

    // Open comments side panel.
    await page.click('[data-testid=\"open-comments-panel\"]');

    // Add a new comment.
    await page.fill('[data-testid=\"new-comment-input\"]', "Hello from e2e");
    await page.click('[data-testid=\"submit-comment\"]');

    // Reply.
    await page.fill('[data-testid=\"reply-input\"]', "Reply from e2e");
    await page.click('[data-testid=\"submit-reply\"]');

    // Resolve.
    await page.click('[data-testid=\"resolve-comment\"]');

    // Validate UI reflects resolved status.
    await expect(page.locator('[data-testid=\"comment-thread\"]')).toHaveAttribute("data-resolved", "true");
  });
});

