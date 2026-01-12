import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function waitForIdle(page: import("@playwright/test").Page): Promise<void> {
  // Vite may occasionally trigger a one-time full reload after dependency optimization.
  // Retry once if the execution context is destroyed mid-wait.
  for (let attempt = 0; attempt < 2; attempt += 1) {
    try {
      await page.waitForFunction(() => Boolean((window as any).__formulaApp?.whenIdle), null, { timeout: 60_000 });
      await page.evaluate(() => (window as any).__formulaApp.whenIdle());
      return;
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      if (attempt === 0 && message.includes("Execution context was destroyed")) {
        await page.waitForLoadState("domcontentloaded");
        continue;
      }
      throw err;
    }
  }
}

async function waitForGridCanvasesToBeSized(
  page: import("@playwright/test").Page,
  rootSelector: string,
): Promise<void> {
  // Canvas sizing happens asynchronously (ResizeObserver + rAF). Ensure the renderer
  // has produced non-zero backing buffers before attempting hit-testing.
  await page.waitForFunction(
    (selector) => {
      const root = document.querySelector(selector);
      if (!root) return false;
      const canvases = root.querySelectorAll("canvas");
      if (canvases.length === 0) return false;
      return Array.from(canvases).every((c) => (c as HTMLCanvasElement).width > 0 && (c as HTMLCanvasElement).height > 0);
    },
    rootSelector,
    { timeout: 10_000 },
  );
}

test.describe("Grid context menus", () => {
  test("right-clicking a row header opens a menu with Row Height…", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);

    await expect(page.locator("#grid")).toBeVisible();
    // Avoid flaky right-click handling in the desktop shell; dispatch a deterministic contextmenu event.
    await page.evaluate(() => {
      const grid = document.getElementById("grid");
      if (!grid) throw new Error("Missing #grid container");
      const rect = grid.getBoundingClientRect();
      grid.dispatchEvent(
        new MouseEvent("contextmenu", {
          bubbles: true,
          cancelable: true,
          button: 2,
          clientX: rect.left + 10,
          clientY: rect.top + 40,
        }),
      );
    });

    const menu = page.getByTestId("context-menu");
    await expect(menu).toBeVisible();
    await expect(menu.getByRole("button", { name: "Row Height…" })).toBeVisible();
  });

  test("right-clicking a column header opens a menu with Column Width…", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);

    await expect(page.locator("#grid")).toBeVisible();
    // Avoid flaky right-click handling in the desktop shell; dispatch a deterministic contextmenu event.
    await page.evaluate(() => {
      const grid = document.getElementById("grid");
      if (!grid) throw new Error("Missing #grid container");
      const rect = grid.getBoundingClientRect();
      grid.dispatchEvent(
        new MouseEvent("contextmenu", {
          bubbles: true,
          cancelable: true,
          button: 2,
          clientX: rect.left + 100,
          clientY: rect.top + 10,
        }),
      );
    });

    const menu = page.getByTestId("context-menu");
    await expect(menu).toBeVisible();
    await expect(menu.getByRole("button", { name: "Column Width…" })).toBeVisible();
  });

  test("right-clicking a row header in split-view secondary pane opens a menu with Row Height…", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");
    await waitForIdle(page);

    await page.getByTestId("ribbon-root").getByTestId("split-vertical").click();
    const secondary = page.locator("#grid-secondary");
    await expect(secondary).toBeVisible();
    await expect(secondary.locator("canvas")).toHaveCount(3);
    await waitForGridCanvasesToBeSized(page, "#grid-secondary");

    // Avoid flaky right-click handling in the desktop shell; dispatch a deterministic contextmenu event.
    await page.evaluate(() => {
      const grid = document.getElementById("grid-secondary");
      if (!grid) throw new Error("Missing #grid-secondary container");
      const rect = grid.getBoundingClientRect();
      grid.dispatchEvent(
        new MouseEvent("contextmenu", {
          bubbles: true,
          cancelable: true,
          button: 2,
          clientX: rect.left + 10,
          clientY: rect.top + 40,
        }),
      );
    });
    const menu = page.getByTestId("context-menu");
    await expect(menu).toBeVisible();
    await expect(menu.getByRole("button", { name: "Row Height…" })).toBeVisible();
  });

  test("right-clicking a column header in split-view secondary pane opens a menu with Column Width…", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");
    await waitForIdle(page);

    await page.getByTestId("ribbon-root").getByTestId("split-vertical").click();
    const secondary = page.locator("#grid-secondary");
    await expect(secondary).toBeVisible();
    await expect(secondary.locator("canvas")).toHaveCount(3);
    await waitForGridCanvasesToBeSized(page, "#grid-secondary");

    // Avoid flaky right-click handling in the desktop shell; dispatch a deterministic contextmenu event.
    await page.evaluate(() => {
      const grid = document.getElementById("grid-secondary");
      if (!grid) throw new Error("Missing #grid-secondary container");
      const rect = grid.getBoundingClientRect();
      grid.dispatchEvent(
        new MouseEvent("contextmenu", {
          bubbles: true,
          cancelable: true,
          button: 2,
          clientX: rect.left + 100,
          clientY: rect.top + 10,
        }),
      );
    });
    const menu = page.getByTestId("context-menu");
    await expect(menu).toBeVisible();
    await expect(menu.getByRole("button", { name: "Column Width…" })).toBeVisible();
  });
});
