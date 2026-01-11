import type { Page } from "@playwright/test";

/**
 * Navigate to the desktop shell and wait for the e2e harness to be ready.
 *
 * The app boot sequence can involve dynamic imports (WASM engine, scripting runtimes),
 * so make tests wait for `window.__formulaApp` before interacting with the grid.
 */
export async function gotoDesktop(page: Page, path: string = "/"): Promise<void> {
  // Vite may trigger a one-time full reload after dependency optimization. If that
  // happens mid-wait, retry once after the navigation completes.
  for (let attempt = 0; attempt < 2; attempt += 1) {
    try {
      await page.goto(path);
      await page.waitForFunction(() => Boolean((window as any).__formulaApp), undefined, { timeout: 60_000 });
      return;
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      if (attempt === 0 && message.includes("Execution context was destroyed")) {
        await page.waitForLoadState("load");
        continue;
      }
      throw err;
    }
  }
}

export async function waitForDesktopReady(page: Page): Promise<void> {
  for (let attempt = 0; attempt < 2; attempt += 1) {
    try {
      await page.waitForFunction(() => Boolean((window as any).__formulaApp), undefined, { timeout: 60_000 });
      return;
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      if (attempt === 0 && message.includes("Execution context was destroyed")) {
        await page.waitForLoadState("load");
        continue;
      }
      throw err;
    }
  }
}
