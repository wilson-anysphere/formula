import { expect, type Page, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function closeIntentionalOverlays(page: Page): Promise<void> {
  // Defensive: close any menus/modals/toasts that might intentionally extend past the viewport
  // and would otherwise turn this into a flaky "did some overlay happen to be open?" test.
  //
  // Most chrome overflow regressions we've hit are persistent (ribbon/statusbar sizing),
  // so it is safe to close transient UI before asserting.
  for (let i = 0; i < 3; i += 1) {
    await page.keyboard.press("Escape").catch(() => {});
  }
}

async function assertNoHorizontalOverflow(page: Page, { label }: { label: string }): Promise<void> {
  // Wait for a couple animation frames so any resize/layout effects flush before measuring.
  await page.evaluate(
    () =>
      new Promise<void>((resolve) => {
        requestAnimationFrame(() => requestAnimationFrame(() => resolve()));
      }),
  );

  // Ensure font loading doesn't shift sizes after we've checked.
  await page.evaluate(async () => {
    await (document as any).fonts?.ready;
  });

  try {
    await page.waitForFunction(
      () => {
        const doc = document.documentElement;
        const app = document.querySelector<HTMLElement>("#app");
        if (!app) return false;
        return doc.scrollWidth <= doc.clientWidth + 1 && app.scrollWidth <= app.clientWidth + 1;
      },
      undefined,
      { timeout: 5_000 },
    );
  } catch (err) {
    const metrics = await page.evaluate(() => {
      const doc = document.documentElement;
      const app = document.querySelector<HTMLElement>("#app");
      return {
        location: window.location.href,
        viewport: {
          innerWidth: window.innerWidth,
          innerHeight: window.innerHeight,
          devicePixelRatio: window.devicePixelRatio,
        },
        documentElement: {
          scrollWidth: doc.scrollWidth,
          clientWidth: doc.clientWidth,
          overflowPx: doc.scrollWidth - doc.clientWidth,
        },
        app: app
          ? {
              scrollWidth: app.scrollWidth,
              clientWidth: app.clientWidth,
              overflowPx: app.scrollWidth - app.clientWidth,
            }
          : null,
      };
    });

    const message = err instanceof Error ? err.message : String(err);
    throw new Error(`Horizontal overflow detected (${label}): ${message}\n\n${JSON.stringify(metrics, null, 2)}`);
  }

  const { documentElement, app } = await page.evaluate(() => {
    const doc = document.documentElement;
    const app = document.querySelector<HTMLElement>("#app");
    if (!app) throw new Error('Missing root element: "#app"');
    return {
      documentElement: { scrollWidth: doc.scrollWidth, clientWidth: doc.clientWidth },
      app: { scrollWidth: app.scrollWidth, clientWidth: app.clientWidth },
    };
  });

  expect(documentElement.scrollWidth).toBeLessThanOrEqual(documentElement.clientWidth + 1);
  expect(app.scrollWidth).toBeLessThanOrEqual(app.clientWidth + 1);
}

test.describe("horizontal overflow", () => {
  const viewports = [700, 900, 1200] as const;

  for (const width of viewports) {
    test(`desktop chrome does not introduce horizontal overflow at ${width}px`, async ({ page }) => {
      await page.setViewportSize({ width, height: 900 });
      await gotoDesktop(page);

      // Ensure the chrome is rendered before measuring.
      await expect(page.getByTestId("ribbon-root")).toBeVisible();
      await expect(page.locator(".statusbar__main")).toBeVisible();

      await closeIntentionalOverlays(page);

      await assertNoHorizontalOverflow(page, { label: `${width}px viewport` });
    });
  }
});
