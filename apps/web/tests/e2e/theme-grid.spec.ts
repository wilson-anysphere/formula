import { expect, test } from "@playwright/test";

test("CanvasGrid reacts to prefers-color-scheme + prefers-contrast theme changes", async ({ page }) => {
  await page.emulateMedia({ colorScheme: "light", contrast: "no-preference" });
  await page.goto("/");

  await expect(page.getByTestId("engine-status")).toContainText("ready", { timeout: 30_000 });

  const gridContainer = page.getByTestId("canvas-grid");

  const vTrack = gridContainer.locator('div[aria-hidden="true"]').first();
  const vThumb = vTrack.locator("div").first();

  const thumbBg = () => vThumb.evaluate((el) => getComputedStyle(el).backgroundColor);

  const lightBg = await thumbBg();
  // Default light theme uses `rgba(0, 0, 0, 0.25)` for the thumb.
  expect(lightBg).toContain("0.25");

  // Emulate high-contrast and ensure the thumb theme token updates.
  await page.emulateMedia({ colorScheme: "light", contrast: "more" });
  await expect.poll(thumbBg).not.toBe(lightBg);

  const contrastBg = await thumbBg();
  expect(contrastBg).not.toContain("0.25");

  // Emulate dark + high-contrast and ensure the thumb theme token updates again.
  await page.emulateMedia({ colorScheme: "dark", contrast: "more" });
  await expect.poll(thumbBg).not.toBe(contrastBg);
});

