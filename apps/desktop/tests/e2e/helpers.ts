import type { Page } from "@playwright/test";

type DesktopReadyOptions = {
  /**
   * Whether to call `__formulaApp.whenIdle()` after `__formulaApp` is available.
   *
   * Defaults to true because most e2e tests want the app (engine, grid, workers)
   * to finish bootstrapping before interacting.
   */
  waitForIdle?: boolean;
  /**
   * Optional cap on how long we'll wait for `whenIdle()` before proceeding.
   * This is useful for tests that intentionally enable background activity (e.g. collaboration),
   * where the app may never become fully idle.
   */
  idleTimeoutMs?: number;
  /**
   * Optional cap on how long we'll wait for `window.__formulaApp` after navigation.
   *
   * Collaboration tests can be slower on first-run (WASM, python runtime, Vite optimize),
   * and should pass a larger value.
   */
  appReadyTimeoutMs?: number;
};

/**
 * Navigate to the desktop shell and wait for the e2e harness to be ready.
 *
 * The app boot sequence can involve dynamic imports (WASM engine, scripting runtimes),
 * so make tests wait for `window.__formulaApp` before interacting with the grid.
 */
export async function gotoDesktop(page: Page, path: string = "/", options: DesktopReadyOptions = {}): Promise<void> {
  const { waitForIdle = true, idleTimeoutMs, appReadyTimeoutMs } = options;
  const consoleErrors: string[] = [];
  const pageErrors: string[] = [];
  const requestFailures: string[] = [];

  const onConsole = (msg: any): void => {
    try {
      if (msg?.type?.() !== "error") return;
      consoleErrors.push(msg.text());
    } catch {
      // ignore listener failures
    }
  };

  const onPageError = (err: unknown): void => {
    const message =
      err instanceof Error
        ? `${err.name}: ${err.message}\n${err.stack ?? ""}`.trim()
        : String(err);
    pageErrors.push(message);
  };

  const onRequestFailed = (req: any): void => {
    try {
      const method = typeof req?.method === "function" ? req.method() : "REQUEST";
      const url = typeof req?.url === "function" ? req.url() : String(req?.url ?? "");
      const failure = typeof req?.failure === "function" ? req.failure() : null;
      const suffix = failure?.errorText ? ` (${failure.errorText})` : "";
      requestFailures.push(`${method} ${url}${suffix}`.trim());
    } catch {
      // ignore listener failures
    }
  };

  page.on("console", onConsole);
  page.on("pageerror", onPageError);
  page.on("requestfailed", onRequestFailed);

  const formatStartupDiagnostics = async (): Promise<string> => {
    const uniqueConsole = [...new Set(consoleErrors)];
    const uniquePage = [...new Set(pageErrors)];
    const uniqueRequests = [...new Set(requestFailures)];

    let appProbe: { present: boolean; truthy: boolean; type: string; nullish: boolean; ctor: string | null } | null = null;
    try {
      appProbe = await page.evaluate(() => {
        const present = "__formulaApp" in window;
        const value = window.__formulaApp;
        return {
          present,
          truthy: Boolean(value),
          type: typeof value,
          nullish: value == null,
          ctor: value && typeof value === "object" ? (value as any).constructor?.name ?? null : null,
        };
      });
    } catch {
      // ignore
    }

    let viteOverlayText = "";
    try {
      viteOverlayText = await page.evaluate(() => {
        const overlay = document.querySelector("vite-error-overlay") as any;
        if (!overlay) return "";
        return (overlay.textContent ?? "").trim();
      });
    } catch {
      // ignore
    }

    const parts: string[] = [];
    if (uniqueConsole.length > 0) parts.push(`Console errors:\n${uniqueConsole.join("\n")}`);
    if (uniquePage.length > 0) parts.push(`Page errors:\n${uniquePage.join("\n")}`);
    if (uniqueRequests.length > 0) parts.push(`Request failures:\n${uniqueRequests.join("\n")}`);
    if (appProbe) parts.push(`window.__formulaApp probe:\n${JSON.stringify(appProbe, null, 2)}`);
    if (viteOverlayText) parts.push(`Vite error overlay:\n${viteOverlayText}`);
    return parts.length > 0 ? `\n\n${parts.join("\n\n")}` : "";
  };

  // Vite may trigger a one-time full reload after dependency optimization. If that
  // happens mid-wait, retry once after the navigation completes.
  for (let attempt = 0; attempt < 2; attempt += 1) {
    try {
      // Desktop e2e relies on waiting for `__formulaApp` (and `whenIdle`) rather than the
      // window `load` event. Under heavy load, waiting for `load` can occasionally hang
      // (e.g. if a long-lived request prevents the event from firing).
      await page.goto(path, { waitUntil: "domcontentloaded" });
      await page.waitForFunction(() => Boolean(window.__formulaApp), undefined, {
        timeout: typeof appReadyTimeoutMs === "number" && appReadyTimeoutMs > 0 ? appReadyTimeoutMs : 60_000,
      });
      // `__formulaApp` is assigned early in `main.ts` so tests can still introspect failures,
      // but that means we need to explicitly wait for the app to settle before interacting.
      await page.evaluate(async ({ waitForIdle, idleTimeoutMs }) => {
        const app = window.__formulaApp as any;
        if (!waitForIdle) return;
        if (app && typeof app.whenIdle === "function") {
          if (typeof idleTimeoutMs === "number" && idleTimeoutMs > 0) {
            await Promise.race([app.whenIdle(), new Promise<void>((r) => setTimeout(r, idleTimeoutMs))]);
          } else {
            await app.whenIdle();
          }
        }
      }, { waitForIdle, idleTimeoutMs });

      // The desktop build includes a built-in "formula.e2e-events" extension used by other
      // Playwright specs (e.g. `extension-events.spec.ts`). That extension writes to
      // `formula.storage`, which is gated by the `"storage"` permission. If we don't pre-seed a
      // grant here, unrelated tests can flake when the permission prompt appears and blocks
      // pointer events.
      //
      // Merge the grant instead of overwriting because many specs pre-configure additional
      // extension permissions via `page.addInitScript(...)`.
      await page.evaluate(() => {
        const key = "formula.extensionHost.permissions";
        const extensionId = "formula.e2e-events";
        let existing: any = {};
        try {
          const raw = localStorage.getItem(key);
          existing = raw ? JSON.parse(raw) : {};
        } catch {
          existing = {};
        }
        if (!existing || typeof existing !== "object" || Array.isArray(existing)) {
          existing = {};
        }
        const current = existing[extensionId];
        const merged =
          current && typeof current === "object" && !Array.isArray(current)
            ? { ...current, storage: true }
            : { storage: true };
        existing[extensionId] = merged;
        try {
          localStorage.setItem(key, JSON.stringify(existing));
        } catch {
          // ignore storage errors (disabled/quota/etc.)
        }
      });

      page.off("console", onConsole);
      page.off("pageerror", onPageError);
      page.off("requestfailed", onRequestFailed);
      return;
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      if (
        attempt === 0 &&
        (message.includes("Execution context was destroyed") ||
          message.includes("net::ERR_ABORTED") ||
          message.includes("frame was detached"))
      ) {
        await page.waitForLoadState("domcontentloaded");
        continue;
      }
      const diag = await formatStartupDiagnostics();
      page.off("console", onConsole);
      page.off("pageerror", onPageError);
      page.off("requestfailed", onRequestFailed);
      throw new Error(`${message}${diag}`);
    }
  }

  page.off("console", onConsole);
  page.off("pageerror", onPageError);
  page.off("requestfailed", onRequestFailed);
}

export async function waitForDesktopReady(page: Page): Promise<void> {
  for (let attempt = 0; attempt < 2; attempt += 1) {
    try {
      await page.waitForFunction(() => Boolean(window.__formulaApp), undefined, { timeout: 60_000 });
      await page.evaluate(async () => {
        const app = window.__formulaApp as any;
        if (app && typeof app.whenIdle === "function") {
          await app.whenIdle();
        }
      });
      return;
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      if (
        attempt === 0 &&
        (message.includes("Execution context was destroyed") ||
          message.includes("net::ERR_ABORTED") ||
          message.includes("frame was detached"))
      ) {
        await page.waitForLoadState("domcontentloaded");
        continue;
      }
      throw err;
    }
  }
}

/**
 * Click the ribbon button that toggles the Extensions panel.
 *
 * The only stable `data-testid="open-extensions-panel"` control lives in the ribbon (React).
 * The static `apps/desktop/index.html` shell must not include legacy debug buttons with the same
 * test id (to avoid Playwright strict-mode collisions), so always scope this locator to the ribbon.
 */
export async function openExtensionsPanel(page: Page): Promise<void> {
  const panel = page.getByTestId("panel-extensions");
  const panelVisible = await panel.isVisible().catch(() => false);

  if (!panelVisible) {
    const ribbonRoot = page.getByTestId("ribbon-root");
    const ribbonButton = ribbonRoot.getByTestId("open-extensions-panel");
    const ribbonButtonVisible = await ribbonButton.isVisible().catch(() => false);
    if (!ribbonButtonVisible) {
      // The Extensions toggle lives in the Home ribbon tab. If another tab is active, the button
      // may not be rendered/visible, so select Home before clicking.
      await ribbonRoot.getByRole("tab", { name: "Home", exact: true }).click();
    }
    await ribbonButton.click({ timeout: 30_000 });

    await panel.waitFor({ state: "visible", timeout: 30_000 });
  }

  // If the panel was already open (e.g. persisted layout after reload), it can render the
  // "Loading extensions…" placeholder until the lazy extension host boot completes. Wait for
  // the panel body to settle (either loaded or errored) before returning.
  await page.waitForFunction(() => {
    const panel = document.querySelector('[data-testid="panel-extensions"]');
    if (!panel) return false;
    const text = panel.textContent ?? "";
    return !text.includes("Loading extensions");
  }, undefined, { timeout: 30_000 });
}
