import { expect, type Locator, type Page } from "@playwright/test";

// Playwright's desktop e2e suite uses a 60s per-test timeout (see `apps/desktop/playwright.config.ts`).
// Keep the default `__formulaApp` readiness wait slightly below that so failures surface the rich
// diagnostics from these helpers instead of tripping the global timeout first.
const DEFAULT_APP_READY_TIMEOUT_MS = 55_000;

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
  /**
   * Whether to wait for the desktop context menu overlay to be mounted.
   *
   * Many e2e specs open context menus early; if the dock-layout bootstrap is still
   * initializing (or a delayed Vite reload is in progress), the `ContextMenu` instance
   * might not be attached yet even though `window.__formulaApp` exists.
   *
   * Defaults to true.
   */
  waitForContextMenu?: boolean;
  /**
   * Optional cap on how long we'll wait for the context menu overlay to mount.
   */
  contextMenuTimeoutMs?: number;
};

/**
 * Navigate to the desktop shell and wait for the e2e harness to be ready.
 *
 * The app boot sequence can involve dynamic imports (WASM engine, scripting runtimes),
 * so make tests wait for `window.__formulaApp` before interacting with the grid.
 */
export async function gotoDesktop(page: Page, path: string = "/", options: DesktopReadyOptions = {}): Promise<void> {
  const {
    waitForIdle = true,
    idleTimeoutMs,
    appReadyTimeoutMs,
    waitForContextMenu = true,
    contextMenuTimeoutMs,
  } = options;
  const consoleErrors: string[] = [];
  const pageErrors: string[] = [];
  const requestFailures: string[] = [];
  let signalNetworkChanged: (() => void) | null = null;

  const onConsole = (msg: any): void => {
    try {
      if (msg?.type?.() !== "error") return;
      const text = msg.text();
      consoleErrors.push(text);
      if (text.includes("net::ERR_NETWORK_CHANGED")) {
        signalNetworkChanged?.();
      }
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
      if (typeof failure?.errorText === "string" && failure.errorText.includes("net::ERR_NETWORK_CHANGED")) {
        signalNetworkChanged?.();
      }
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

  const maxAttempts = 5;

  // Vite may trigger a full reload after dependency optimization (or when it discovers a new
  // dynamic import). When that happens mid-navigation, Chromium surfaces `net::ERR_NETWORK_CHANGED`
  // for module requests and the app never boots. Retry a few times so desktop e2e is resilient on
  // first-run optimize behavior.
  for (let attempt = 0; attempt < maxAttempts; attempt += 1) {
    const requestStart = requestFailures.length;
    let networkChanged = false;
    let resolveNetworkChanged: (() => void) | null = null;
    const networkChangedPromise = new Promise<void>((resolve) => {
      resolveNetworkChanged = resolve;
    });
    signalNetworkChanged = () => {
      if (networkChanged) return;
      networkChanged = true;
      resolveNetworkChanged?.();
    };

    try {
      // Desktop e2e relies on waiting for `__formulaApp` (and `whenIdle`) rather than the
      // window `load` event. Under heavy load, waiting for `load` can occasionally hang
      // (e.g. if a long-lived request prevents the event from firing).
      await page.goto(path, { waitUntil: "domcontentloaded" });
      const appReadyTimeout =
        typeof appReadyTimeoutMs === "number" && appReadyTimeoutMs > 0 ? appReadyTimeoutMs : DEFAULT_APP_READY_TIMEOUT_MS;
      const appReadyPromise = page.waitForFunction(() => Boolean(window.__formulaApp), undefined, {
        timeout: appReadyTimeout,
      });
      // If we abandon this attempt due to a Vite reload, `appReadyPromise` can reject later (after
      // its own timeout) and trigger unhandled rejection noise. Attach a handler now so retries
      // remain clean.
      void appReadyPromise.catch(() => {});

      // Vite can surface transient `net::ERR_NETWORK_CHANGED` errors during dependency
      // optimization. If the page recovers quickly (e.g. via an automatic reload), allow the boot
      // to complete; otherwise treat the network change as a signal to retry navigation.
      await Promise.race([
        appReadyPromise,
        networkChangedPromise.then(async () => {
          await Promise.race([appReadyPromise, page.waitForTimeout(Math.min(10_000, appReadyTimeout))]);
          const ready = await page.evaluate(() => Boolean(window.__formulaApp)).catch(() => false);
          if (ready) return;
          throw new Error("net::ERR_NETWORK_CHANGED");
        }),
      ]);

      // `__formulaApp` is assigned early in `main.ts` so tests can still introspect failures,
      // but that means we need to explicitly wait for the app to settle before interacting.
      await page.evaluate(
        async ({ waitForIdle, idleTimeoutMs }) => {
          const app = window.__formulaApp as any;
          if (!waitForIdle) return;
          if (app && typeof app.whenIdle === "function") {
            if (typeof idleTimeoutMs === "number" && idleTimeoutMs > 0) {
              await Promise.race([app.whenIdle(), new Promise<void>((r) => setTimeout(r, idleTimeoutMs))]);
            } else {
              await app.whenIdle();
            }
          }
        },
        { waitForIdle, idleTimeoutMs },
      );

      if (waitForContextMenu) {
        const menuWait = page.waitForFunction(
          () => Boolean(document.querySelector('[data-testid="context-menu"]')),
          undefined,
          { timeout: typeof contextMenuTimeoutMs === "number" && contextMenuTimeoutMs > 0 ? contextMenuTimeoutMs : 30_000 },
        );
        await menuWait;
      }

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

      // If the desktop shell threw during startup (uncaught exception), the app can partially
      // mount (e.g. titlebar/grid) while leaving other regions (ribbon/sheet tabs) uninitialized.
      // Surface those failures early with the collected diagnostics so flaky "element not found"
      // assertions are easier to debug.
      if (pageErrors.length > 0) {
        const diag = await formatStartupDiagnostics();
        page.off("console", onConsole);
        page.off("pageerror", onPageError);
        page.off("requestfailed", onRequestFailed);
        throw new Error(`Desktop startup failed with page errors.${diag}`);
      }

      page.off("console", onConsole);
      page.off("pageerror", onPageError);
      page.off("requestfailed", onRequestFailed);
      return;
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      const attemptFailures = requestFailures.slice(requestStart);
      const sawNetworkChanged =
        networkChanged || attemptFailures.some((entry) => entry.includes("net::ERR_NETWORK_CHANGED"));
      if (
        attempt < maxAttempts - 1 &&
        (message.includes("Execution context was destroyed") ||
          message.includes("net::ERR_ABORTED") ||
          message.includes("net::ERR_NETWORK_CHANGED") ||
          message.includes("interrupted by another navigation") ||
          message.includes("frame was detached") ||
          sawNetworkChanged)
      ) {
        // Clear captured diagnostics for the retry so a successful second attempt doesn't inherit
        // errors from the first attempt.
        consoleErrors.length = 0;
        pageErrors.length = 0;
        requestFailures.length = 0;
        try {
          await page.waitForLoadState("domcontentloaded");
        } catch {
          // ignore
        }
        // Give the Vite dev server a moment to stabilize (it may restart once after optimizing deps).
        await page.waitForTimeout(250 * (attempt + 1));
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
  const consoleErrors: string[] = [];
  const pageErrors: string[] = [];
  const requestFailures: string[] = [];
  let signalNetworkChanged: (() => void) | null = null;

  const onConsole = (msg: any): void => {
    try {
      if (msg?.type?.() !== "error") return;
      const text = msg.text();
      consoleErrors.push(text);
      if (text.includes("net::ERR_NETWORK_CHANGED")) {
        signalNetworkChanged?.();
      }
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
      if (typeof failure?.errorText === "string" && failure.errorText.includes("net::ERR_NETWORK_CHANGED")) {
        signalNetworkChanged?.();
      }
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

  const appReadyTimeout = DEFAULT_APP_READY_TIMEOUT_MS;
  const maxAttempts = 3;
  for (let attempt = 0; attempt < maxAttempts; attempt += 1) {
    const requestStart = requestFailures.length;
    let networkChanged = false;
    let resolveNetworkChanged: (() => void) | null = null;
    const networkChangedPromise = new Promise<void>((resolve) => {
      resolveNetworkChanged = resolve;
    });
    signalNetworkChanged = () => {
      if (networkChanged) return;
      networkChanged = true;
      resolveNetworkChanged?.();
    };

    try {
      const appWait = page.waitForFunction(() => Boolean(window.__formulaApp), undefined, { timeout: appReadyTimeout });
      // If we abandon this attempt due to a Vite reload, `appWait` can reject later (after its
      // own timeout) and trigger unhandled rejection noise. Attach a handler now so retries remain
      // clean.
      void appWait.catch(() => {});
      await Promise.race([
        appWait,
        networkChangedPromise.then(async () => {
          await Promise.race([appWait, page.waitForTimeout(10_000)]);
          const ready = await page.evaluate(() => Boolean(window.__formulaApp)).catch(() => false);
          if (ready) return;
          throw new Error("net::ERR_NETWORK_CHANGED");
        }),
      ]);

      await page.evaluate(async () => {
        const app = window.__formulaApp as any;
        if (app && typeof app.whenIdle === "function") {
          await app.whenIdle();
        }
      });

      page.off("console", onConsole);
      page.off("pageerror", onPageError);
      page.off("requestfailed", onRequestFailed);
      return;
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      const attemptFailures = requestFailures.slice(requestStart);
      const sawNetworkChanged =
        networkChanged ||
        consoleErrors.some((text) => text.includes("net::ERR_NETWORK_CHANGED")) ||
        attemptFailures.some((entry) => entry.includes("net::ERR_NETWORK_CHANGED"));
      if (
        attempt < maxAttempts - 1 &&
        (message.includes("Execution context was destroyed") ||
          message.includes("net::ERR_ABORTED") ||
          message.includes("net::ERR_NETWORK_CHANGED") ||
          message.includes("interrupted by another navigation") ||
          message.includes("frame was detached") ||
          sawNetworkChanged)
      ) {
        // Clear captured diagnostics for the retry so a successful second attempt doesn't inherit
        // errors from the first attempt.
        consoleErrors.length = 0;
        pageErrors.length = 0;
        requestFailures.length = 0;

        try {
          // If Vite restarted mid-load we can end up in a state where module requests fail and the
          // app never boots. Force a reload so the next attempt has a clean slate.
          await page.reload({ waitUntil: "domcontentloaded" });
        } catch {
          // Fall back to waiting for the current navigation to settle.
          try {
            await page.waitForLoadState("domcontentloaded");
          } catch {
            // ignore
          }
        }
        // Give the dev server a moment to stabilize (it may restart once after optimizing deps).
        await page.waitForTimeout(250 * (attempt + 1));
        continue;
      }

      const diag = await formatStartupDiagnostics();
      signalNetworkChanged = null;
      page.off("console", onConsole);
      page.off("pageerror", onPageError);
      page.off("requestfailed", onRequestFailed);
      throw new Error(`${message}${diag}`);
    }
  }

  signalNetworkChanged = null;
  page.off("console", onConsole);
  page.off("pageerror", onPageError);
  page.off("requestfailed", onRequestFailed);
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
    // Prefer the ribbon UI when available (so we still cover the click path), but fall back to
    // command execution when the ribbon isn't mounted (some harnesses may omit it, and it can be
    // slow to initialize under heavy load).
    const ribbonRoot = page.getByTestId("ribbon-root");
    const ribbonAvailable = await ribbonRoot.isVisible().catch(() => false);
    const ribbonButton = ribbonRoot.getByTestId("open-extensions-panel");
    const ribbonButtonVisible = ribbonAvailable ? await ribbonButton.isVisible().catch(() => false) : false;

    if (ribbonAvailable && ribbonButtonVisible) {
      await ribbonButton.click({ timeout: 30_000 });
    } else {
      // Execute the canonical command directly via the exposed CommandRegistry.
      await page.evaluate(async () => {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const registry: any = (window as any).__formulaCommandRegistry;
        if (!registry) throw new Error("Missing window.__formulaCommandRegistry (desktop e2e harness)");
        await registry.executeCommand("view.togglePanel.extensions");
      });
    }

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

  // Extension host globals are installed lazily. Many suites access `window.__formulaExtensionHost`
  // directly, so ensure it's present before returning.
  await page.waitForFunction(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const w: any = window as any;
    return Boolean(w.__formulaExtensionHost) || Boolean(w.__formulaExtensionHostManager);
  }, undefined, { timeout: 30_000 });
}

/**
 * Open the sheet tab context menu for a given sheet id.
 *
 * We dispatch a deterministic `contextmenu` event instead of relying on Playwright's
 * right-click helpers, which can be flaky in the desktop shell (and in some WebView
 * environments where native context menu handling differs).
 */
export async function openSheetTabContextMenu(page: Page, sheetId: string): Promise<Locator> {
  await page.evaluate((id) => {
    // Avoid interpolating the sheet id into a CSS selector (ids can contain characters
    // that require escaping). Instead, query all sheet tab buttons and match by attribute.
    //
    // Prefer `data-sheet-id` over `data-testid`: `data-testid^="sheet-tab-"` would also
    // match the context menu overlay (`data-testid="sheet-tab-context-menu"`), which could
    // collide if a sheet id ever equals `"context-menu"`.
    const tab =
      Array.from(document.querySelectorAll<HTMLButtonElement>('button[role="tab"][data-sheet-id]')).find(
        (el) => el.dataset.sheetId === id,
      ) ?? null;
    if (!tab) throw new Error(`Missing sheet tab button for sheet id: ${id}`);
    const rect = tab.getBoundingClientRect();
    const x = rect.left + rect.width / 2;
    const y = rect.top + rect.height / 2;
    // Clamp to the viewport so the context menu anchor is never off-screen, even if the
    // tab is scrolled out of view.
    const clampedX = Math.max(0, Math.min(x, window.innerWidth - 1));
    const clampedY = Math.max(0, Math.min(y, window.innerHeight - 1));
    tab.dispatchEvent(
      new MouseEvent("contextmenu", {
        bubbles: true,
        cancelable: true,
        // Explicitly mark this as a right-click.
        button: 2,
        // Use the center of the tab so narrow tabs still receive the event.
        clientX: clampedX,
        clientY: clampedY,
      }),
    );
  }, sheetId);

  const menu = page.getByTestId("sheet-tab-context-menu");
  await menu.waitFor({ state: "visible", timeout: 10_000 });
  return menu;
}

export async function openSheetTabStripContextMenu(page: Page): Promise<Locator> {
  await page.evaluate(() => {
    const strip = document.querySelector<HTMLElement>("#sheet-tabs .sheet-tabs");
    if (!strip) throw new Error("Missing sheet tab strip");
    const rect = strip.getBoundingClientRect();
    const x = rect.left + rect.width - 4;
    const y = rect.top + rect.height / 2;
    const clampedX = Math.max(0, Math.min(x, window.innerWidth - 1));
    const clampedY = Math.max(0, Math.min(y, window.innerHeight - 1));
    strip.dispatchEvent(
      new MouseEvent("contextmenu", {
        bubbles: true,
        cancelable: true,
        button: 2,
        // Aim near the end of the strip so the menu isn't rendered on top of the first tab.
        clientX: clampedX,
        clientY: clampedY,
      }),
    );
  });

  const menu = page.getByTestId("sheet-tab-context-menu");
  await menu.waitFor({ state: "visible", timeout: 10_000 });
  return menu;
}

export async function expectSheetPosition(
  page: Page,
  { position, total }: { position: number; total: number },
  { timeoutMs = 10_000 }: { timeoutMs?: number } = {},
): Promise<void> {
  await expect
    .poll(
      async () => {
        try {
          return await page.evaluate(() => {
            const el = document.querySelector('[data-testid="sheet-position"]');
            if (!el) return null;
            const positionAttr = el.getAttribute("data-sheet-position");
            const totalAttr = el.getAttribute("data-sheet-total");
            if (positionAttr != null && totalAttr != null) {
              const position = Number(positionAttr);
              const total = Number(totalAttr);
              if (
                Number.isFinite(position) &&
                Number.isFinite(total) &&
                Number.isInteger(position) &&
                Number.isInteger(total) &&
                position >= 0 &&
                total >= 0
              ) {
                return { position, total };
              }
            }
            const raw = (el?.textContent ?? "").trim();
            const nums = raw.match(/\d+/g) ?? [];
            if (nums.length < 2) return null;
            const position = Number(nums[0]);
            const total = Number(nums[1]);
            if (
              Number.isFinite(position) &&
              Number.isFinite(total) &&
              Number.isInteger(position) &&
              Number.isInteger(total) &&
              position >= 0 &&
              total >= 0
            ) {
              return { position, total };
            }
            return null;
          });
        } catch {
          return null;
        }
      },
      { timeout: timeoutMs },
    )
    .toEqual({ position, total });
}
