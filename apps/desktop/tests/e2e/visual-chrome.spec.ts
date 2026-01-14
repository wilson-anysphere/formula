import { expect, test } from "@playwright/test";

import { gotoDesktop, waitForDesktopReady } from "./helpers";

type Theme = "light" | "dark";
type RibbonDensity = "full" | "compact" | "hidden";

type Scenario = {
  label: string;
  viewport: { width: number; height: number };
  ribbonDensity: RibbonDensity;
};

const scenarios: Scenario[] = [
  { label: "full", viewport: { width: 1400, height: 900 }, ribbonDensity: "full" },
  { label: "compact", viewport: { width: 1000, height: 900 }, ribbonDensity: "compact" },
  { label: "hidden", viewport: { width: 700, height: 900 }, ribbonDensity: "hidden" },
];

const themes: Theme[] = ["light", "dark"];

const screenshotOptions = {
  animations: "disabled" as const,
  caret: "hide" as const,
  // Element screenshots occasionally take longer to stabilize on CI (fonts/layout); keep a
  // generous timeout to reduce flakes.
  timeout: 15_000,
};

async function applyVisualStabilizers(page: import("@playwright/test").Page, theme: Theme): Promise<void> {
  await page.evaluate((theme) => {
    document.documentElement.setAttribute("data-theme", theme);
    // Prefer a deterministic reduced motion flag even when the environment does not
    // advertise `prefers-reduced-motion: reduce`.
    document.documentElement.setAttribute("data-reduced-motion", "true");
  }, theme);

  // Ensure the DOM has committed style/layout updates before capturing screenshots.
  await page.evaluate(() => new Promise<void>((resolve) => requestAnimationFrame(() => requestAnimationFrame(() => resolve()))));
}

async function seedWorkbookForStableChrome(page: import("@playwright/test").Page): Promise<void> {
  await page.evaluate(() => {
    const app = (window as any).__formulaApp;
    if (!app) throw new Error("Missing window.__formulaApp");
    const doc = app.getDocument?.();
    if (!doc) throw new Error("Missing DocumentController");

    // Create a few sheets so the tab strip renders both active + inactive tabs.
    doc.setCellValue("Sheet2", "A1", "Two");
    doc.setCellValue("Sheet3", "A1", "Three");

    // Keep selection / status bar deterministic.
    app.activateCell?.({ row: 0, col: 0 }); // A1
  });

  await expect(page.getByTestId("sheet-tab-Sheet1")).toBeVisible();
  await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();
  await expect(page.getByTestId("sheet-tab-Sheet3")).toBeVisible();
}

test.describe("visual chrome", () => {
  for (const theme of themes) {
    for (const scenario of scenarios) {
      test(`visual chrome (${theme}, ${scenario.label}, ${scenario.viewport.width}x${scenario.viewport.height})`, async ({ page }) => {
        // Ensure deterministic ribbon/theming state before the app bootstraps.
        await page.addInitScript((theme: Theme) => {
          try {
            localStorage.setItem("formula.ui.ribbonCollapsed", "false");
            localStorage.setItem("formula.settings.appearance.v1", JSON.stringify({ themePreference: theme }));
          } catch {
            // ignore storage errors (disabled/quota/etc.)
          }
        }, theme);

        await page.setViewportSize(scenario.viewport);

        await gotoDesktop(page);
        await waitForDesktopReady(page);

        await applyVisualStabilizers(page, theme);
        await seedWorkbookForStableChrome(page);
        await waitForDesktopReady(page);

        const titlebar = page.getByTestId("titlebar").locator(".formula-titlebar");
        const ribbon = page.getByTestId("ribbon-root");
        const formulaBar = page.locator("#formula-bar");
        const sheetTabs = page.getByTestId("sheet-tabs");
        const statusbar = page.locator(".statusbar__main");

        await expect(ribbon).toHaveAttribute("data-responsive-density", scenario.ribbonDensity);

        const suffix = `${theme}-${scenario.label}-${scenario.viewport.width}`;
        await expect(titlebar).toHaveScreenshot(`titlebar-${suffix}.png`, screenshotOptions);
        await expect(ribbon).toHaveScreenshot(`ribbon-${suffix}.png`, screenshotOptions);
        await expect(formulaBar).toHaveScreenshot(`formula-bar-${suffix}.png`, screenshotOptions);
        await expect(sheetTabs).toHaveScreenshot(`sheet-tabs-${suffix}.png`, screenshotOptions);
        await expect(statusbar).toHaveScreenshot(`statusbar-${suffix}.png`, screenshotOptions);
      });
    }
  }
});
