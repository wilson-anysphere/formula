import { expect, test, type Locator, type Page } from "@playwright/test";

import { gotoDesktop } from "./helpers";

type Theme = "light" | "dark";

const screenshotOptions = {
  animations: "disabled" as const,
  caret: "hide" as const,
  // Element screenshots occasionally take longer to stabilize on CI (fonts/layout); keep a
  // generous timeout to reduce flakes.
  timeout: 15_000,
};

function commandModifier(): string {
  return process.platform === "darwin" ? "Meta" : "Control";
}

async function setTheme(page: Page, theme: Theme): Promise<void> {
  const ribbonRoot = page.getByTestId("ribbon-root");
  await ribbonRoot.getByRole("tab", { name: "View", exact: true }).click();

  const themeDropdown = ribbonRoot.getByTestId("theme-selector");
  await expect(themeDropdown).toBeVisible();
  await themeDropdown.click();

  await page.locator(`[role="menuitem"][data-command-id="view.appearance.theme.${theme}"]`).click();
  await expect(page.locator("html")).toHaveAttribute("data-theme", theme);
}

async function openCommandPalette(page: Page): Promise<Locator> {
  await page.keyboard.press(`${commandModifier()}+Shift+P`);
  const palette = page.getByTestId("command-palette");
  await expect(palette).toBeVisible();
  // Ensure the listbox has populated at least one item before capturing.
  await expect(page.getByTestId("command-palette-list").locator("li.command-palette__item").first()).toBeVisible();
  return palette;
}

async function openContextMenu(page: Page): Promise<Locator> {
  // Ensure the grid is focused and has an active cell.
  await page.click("#grid", { position: { x: 80, y: 40 } });
  await expect(page.getByTestId("active-cell")).toHaveText("A1");

  await page.keyboard.press("Shift+F10");
  const menu = page.getByTestId("context-menu");
  await expect(menu).toBeVisible();

  // Stabilize the default state by explicitly focusing the same menu item each run.
  await menu.getByRole("button", { name: "Copy" }).focus();
  return menu;
}

async function openFormatCellsDialog(page: Page): Promise<Locator> {
  const ribbonRoot = page.getByTestId("ribbon-root");
  await ribbonRoot.getByRole("tab", { name: "Home", exact: true }).click();

  // Prefer the dedicated button when available (stable, avoids dropdown timing).
  const directButton = ribbonRoot.getByTestId("ribbon-format-cells");
  if (await directButton.isVisible()) {
    await directButton.click();
  } else {
    await ribbonRoot.locator('[data-command-id="home.number.moreFormats"]').click();
    // The dropdown menu uses the canonical command id.
    await page.locator('[role="menuitem"][data-command-id="format.openFormatCells"]').click();
  }

  const dialog = page.getByTestId("format-cells-dialog");
  await expect(dialog).toBeVisible();
  // Avoid default focus rings / caret capture in snapshots.
  await page.evaluate(() => (document.activeElement instanceof HTMLElement ? document.activeElement.blur() : undefined));
  return dialog;
}

async function openCommentsPanel(page: Page): Promise<Locator> {
  const ribbonRoot = page.getByTestId("ribbon-root");
  await ribbonRoot.getByRole("tab", { name: "Home", exact: true }).click();
  await ribbonRoot.getByTestId("open-comments-panel").click();

  const panel = page.getByTestId("comments-panel");
  await expect(panel).toBeVisible();
  return panel;
}

test.describe("visual overlays", () => {
  test.use({
    viewport: { width: 1280, height: 720 },
    reducedMotion: "reduce",
  });

  for (const theme of ["light", "dark"] as const) {
    test(`${theme} theme`, async ({ page }) => {
      await gotoDesktop(page);
      await setTheme(page, theme);

      const palette = await openCommandPalette(page);
      await expect(palette).toHaveScreenshot(`command-palette-${theme}.png`, screenshotOptions);
      await page.keyboard.press("Escape");
      await expect(palette).toBeHidden();

      const contextMenu = await openContextMenu(page);
      await expect(contextMenu).toHaveScreenshot(`context-menu-${theme}.png`, screenshotOptions);
      await page.keyboard.press("Escape");
      await expect(contextMenu).toBeHidden();

      const formatCellsDialog = await openFormatCellsDialog(page);
      await expect(formatCellsDialog).toHaveScreenshot(`format-cells-dialog-${theme}.png`, screenshotOptions);
      await page.keyboard.press("Escape");
      await expect(formatCellsDialog).toBeHidden();

      const commentsPanel = await openCommentsPanel(page);
      await expect(commentsPanel).toHaveScreenshot(`comments-panel-${theme}.png`, screenshotOptions);
      // Toggle back off to avoid leaving it open if we add more screenshots later.
      await page.getByTestId("ribbon-root").getByTestId("open-comments-panel").click();
      await expect(commentsPanel).toBeHidden();
    });
  }
});
