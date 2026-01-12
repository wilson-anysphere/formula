import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("collab status indicator", () => {
  test("shows Local when collaboration is not enabled", async ({ page }) => {
    await gotoDesktop(page);
    await expect(page.getByTestId("collab-status")).toBeVisible();
    await expect(page.getByTestId("collab-status")).toHaveAttribute("data-collab-mode", "local");
    await expect(page.getByTestId("collab-status")).not.toHaveAttribute("data-collab-doc-id", /.+/);
    await expect(page.getByTestId("collab-status")).toHaveText("Local");
  });
});
