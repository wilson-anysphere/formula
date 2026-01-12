import { expect, test } from "@playwright/test";

import { gotoDesktop, waitForDesktopReady } from "./helpers";

test.describe("marketplace panel", () => {
  test.beforeEach(async ({ page }) => {
    await page.addInitScript(() => {
      try {
        localStorage.setItem(
          "formula.extensionHost.permissions",
          JSON.stringify({
            "formula.e2e-events": { storage: true },
          }),
        );
      } catch {
        // ignore
      }
    });
  });

  test("opens from the ribbon (View → Panels)", async ({ page }) => {
    await gotoDesktop(page);
    await waitForDesktopReady(page);

    const viewTab = page.getByRole("tab", { name: "View", exact: true });
    await expect(viewTab).toBeVisible();
    await viewTab.click();

    const openMarketplacePanel = page.getByTestId("ribbon-root").getByTestId("open-marketplace-panel");
    await expect(openMarketplacePanel).toBeVisible();
    await openMarketplacePanel.click();

    const panel = page.getByTestId("panel-marketplace");
    await expect(panel).toBeVisible();

    // The command is wired via `toggleDockPanel`, so clicking again should close it.
    await openMarketplacePanel.click();
    await expect(panel).toHaveCount(0);
  });

  test("can render search results (stubbed /api/search)", async ({ page }) => {
    await page.route("**/api/search**", async (route) => {
      await route.fulfill({
        status: 200,
        contentType: "application/json",
        body: JSON.stringify({
          total: 1,
          results: [
            {
              id: "acme.hello",
              name: "hello",
              displayName: "Hello Extension",
              publisher: "acme",
              description: "A test extension returned by Playwright stubs.",
              latestVersion: "1.0.0",
              verified: false,
              featured: false,
              categories: [],
              tags: [],
              screenshots: [],
              downloadCount: 0,
              updatedAt: new Date().toISOString(),
            },
          ],
          nextCursor: null,
        }),
      });
    });

    // The panel does a best-effort per-item details fetch (`/api/extensions/:id`). In Vite dev
    // servers this could return HTML; stub it to a clean 404 so the panel consistently falls
    // back to the summary response.
    await page.route("**/api/extensions/**", async (route) => {
      await route.fulfill({ status: 404, body: "" });
    });

    await gotoDesktop(page);
    await waitForDesktopReady(page);

    // The Marketplace panel should be usable without eagerly loading the extension host.
    // (Loading extensions spins up Workers; keep that work lazy until the user opens the
    // Extensions panel or runs a command.)
    await page.waitForFunction(() => Boolean((window as any).__formulaExtensionHostManager), undefined, {
      timeout: 30_000,
    });
    expect(await page.evaluate(() => (window as any).__formulaExtensionHostManager?.ready)).toBe(false);

    await page.getByRole("tab", { name: "View", exact: true }).click();
    await page.getByTestId("ribbon-root").getByTestId("open-marketplace-panel").click();

    const panel = page.getByTestId("panel-marketplace");
    await expect(panel).toBeVisible();

    await panel.getByPlaceholder("Search extensions…").fill("hello");
    await panel.getByRole("button", { name: "Search", exact: true }).click();

    await expect(panel).toContainText("Hello Extension (acme.hello)");

    expect(await page.evaluate(() => (window as any).__formulaExtensionHostManager?.ready)).toBe(false);
  });
});
