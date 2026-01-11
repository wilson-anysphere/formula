import { expect, test } from "@playwright/test";

test("CanvasGrid reacts to prefers-color-scheme + prefers-contrast theme changes", async ({ page }) => {
  await page.emulateMedia({ colorScheme: "light", contrast: "no-preference" });
  await page.goto("/");

  await expect(page.getByTestId("engine-status")).toContainText("ready", { timeout: 30_000 });

  const gridContainer = page.getByTestId("canvas-grid");

  const vTrack = gridContainer.locator('div[aria-hidden="true"]').first();
  const vThumb = vTrack.locator("div").first();

  const thumbBg = () => vThumb.evaluate((el) => getComputedStyle(el).backgroundColor);
  const scrollbarThumbVar = () =>
    page.evaluate(() => getComputedStyle(document.documentElement).getPropertyValue("--formula-grid-scrollbar-thumb").trim());

  const lightBg = await thumbBg();
  const lightThumbVar = await scrollbarThumbVar();
  // Default light theme uses `rgba(0, 0, 0, 0.25)` for the thumb.
  expect(lightBg).toContain("0.25");

  // Emulate dark mode and ensure the thumb theme token updates.
  await page.emulateMedia({ colorScheme: "dark", contrast: "no-preference" });
  await expect.poll(thumbBg).not.toBe(lightBg);

  const darkBg = await thumbBg();
  expect(darkBg).not.toContain("0.25");

  // Emulate high-contrast and ensure the thumb theme token updates.
  // Firefox currently does not reliably emulate `prefers-contrast` via Playwright,
  // so gate the assertions on whether the CSS variables actually change.
  await page.emulateMedia({ colorScheme: "light", contrast: "more" });
  const contrastThumbVar = await scrollbarThumbVar();
  if (!contrastThumbVar || contrastThumbVar === lightThumbVar) return;

  await expect.poll(thumbBg).not.toBe(lightBg);

  const contrastBg = await thumbBg();
  expect(contrastBg).not.toContain("0.25");

  // Emulate dark + high-contrast and ensure the thumb theme token updates again.
  await page.emulateMedia({ colorScheme: "dark", contrast: "more" });
  await expect.poll(thumbBg).not.toBe(contrastBg);
});
