import { expect, test } from "@playwright/test";

test.describe("outline grouping", () => {
  test("collapse and expand a row group", async ({ page }) => {
    await page.goto("/");

    // The demo sheet seeds an outline group for rows 2-4 with a summary row at 5.
    const toggle = page.getByTestId("outline-toggle-row-5");
    await expect(toggle).toHaveText("-");

    // Focus + select A1.
    await page.click("#grid", { position: { x: 60, y: 40 } });
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    // Collapse: rows 2-4 become hidden, so ArrowDown from row 1 lands on row 5.
    await toggle.click();
    await expect(toggle).toHaveText("+");

    await page.keyboard.press("ArrowDown");
    await expect(page.getByTestId("active-cell")).toHaveText("A5");

    // Expand again; ArrowUp should now land on the last detail row (row 4).
    await toggle.click();
    await expect(toggle).toHaveText("-");

    await page.keyboard.press("ArrowUp");
    await expect(page.getByTestId("active-cell")).toHaveText("A4");

    // Column group: columns 2-4 with a summary col at 5 (B-D grouped under E).
    const colToggle = page.getByTestId("outline-toggle-col-5");
    await expect(colToggle).toHaveText("-");

    // Reset focus to A1 for a deterministic jump.
    await page.click("#grid", { position: { x: 60, y: 40 } });
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    await colToggle.click();
    await expect(colToggle).toHaveText("+");

    await page.keyboard.press("ArrowRight");
    await expect(page.getByTestId("active-cell")).toHaveText("E1");

    await colToggle.click();
    await expect(colToggle).toHaveText("-");

    await page.keyboard.press("ArrowLeft");
    await expect(page.getByTestId("active-cell")).toHaveText("D1");
  });
});
