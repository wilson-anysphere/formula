import { expect, test } from "@playwright/test";
import http from "node:http";

import { gotoDesktop, openExtensionsPanel, waitForDesktopReady } from "./helpers";

test.describe("Extensions permissions UI", () => {
  test("can view and revoke extension network permission from the Extensions panel", async ({ page }) => {
    const server = http.createServer((_req, res) => {
      res.writeHead(200, {
        "Content-Type": "text/plain",
        "Access-Control-Allow-Origin": "*",
      });
      res.end("hello");
    });

    await new Promise<void>((resolve) => server.listen(0, "127.0.0.1", resolve));
    const address = server.address();
    const port = typeof address === "object" && address ? address.port : null;
    if (!port) throw new Error("Failed to allocate test port");

    const url = `http://127.0.0.1:${port}/`;
    const extensionId = "formula.sample-hello";

    try {
      await page.addInitScript(() => {
        // Start with a clean permission store so this test exercises the allow/deny UI prompt flow.
        try {
          localStorage.removeItem("formula.extensionHost.permissions");
        } catch {
          // ignore
        }
      });

      await gotoDesktop(page);

      await openExtensionsPanel(page);

      await expect(page.getByTestId(`extension-card-${extensionId}`)).toBeVisible();
      await expect(page.getByTestId(`permission-row-${extensionId}-network`)).toContainText("not granted");

      await page.getByTestId("run-command-with-args-sampleHello.fetchText").click();
      await expect(page.getByTestId("input-box")).toBeVisible();
      await page.getByTestId("input-box-field").fill(JSON.stringify([url]));
      await page.getByTestId("input-box-ok").click();

      await expect(page.getByTestId("extension-permission-prompt")).toBeVisible();
      await expect(page.getByTestId("extension-permission-ui.commands")).toBeVisible();
      await page.getByTestId("extension-permission-allow").click();
      await expect(page.getByTestId("extension-permission-ui.commands")).toHaveCount(0);

      await expect(page.getByTestId("extension-permission-prompt")).toBeVisible();
      await expect(page.getByTestId("extension-permission-network")).toBeVisible();
      await page.getByTestId("extension-permission-allow").click();
      await expect(page.getByTestId("extension-permission-network")).toHaveCount(0);

      await expect(page.getByTestId("toast-root")).toContainText("Fetched: hello");

      await expect(page.getByTestId(`permission-row-${extensionId}-ui.commands`)).toContainText("granted");
      await expect(page.getByTestId(`permission-row-${extensionId}-network`)).toContainText("mode: allowlist");
      await expect(page.getByTestId(`permission-row-${extensionId}-network`)).toContainText("127.0.0.1");

      await page.getByTestId(`revoke-permission-${extensionId}-network`).click();
      await expect(page.getByTestId(`permission-row-${extensionId}-ui.commands`)).toContainText("granted");
      await expect(page.getByTestId(`permission-row-${extensionId}-network`)).toContainText("not granted");

      await page.getByTestId("run-command-with-args-sampleHello.fetchText").click();
      await expect(page.getByTestId("input-box")).toBeVisible();
      await page.getByTestId("input-box-field").fill(JSON.stringify([url]));
      await page.getByTestId("input-box-ok").click();
      await expect(page.getByTestId("extension-permission-prompt")).toBeVisible();
      await expect(page.getByTestId("extension-permission-network")).toBeVisible();
      await page.getByTestId("extension-permission-deny").click();

      await expect(page.getByTestId("toast-root")).toContainText("Permission denied");
      await expect(page.getByTestId(`permission-row-${extensionId}-network`)).toContainText("not granted");
    } finally {
      await new Promise<void>((resolve) => server.close(() => resolve()));
    }
  });

  test("global reset clears extension permissions", async ({ page }) => {
    const server = http.createServer((_req, res) => {
      res.writeHead(200, {
        "Content-Type": "text/plain",
        "Access-Control-Allow-Origin": "*",
      });
      res.end("hello");
    });

    await new Promise<void>((resolve) => server.listen(0, "127.0.0.1", resolve));
    const address = server.address();
    const port = typeof address === "object" && address ? address.port : null;
    if (!port) throw new Error("Failed to allocate test port");

    const url = `http://127.0.0.1:${port}/`;
    const extensionId = "formula.sample-hello";

    try {
      await page.addInitScript(() => {
        try {
          localStorage.removeItem("formula.extensionHost.permissions");
        } catch {
          // ignore
        }
      });

      await gotoDesktop(page);

      await openExtensionsPanel(page);

      await expect(page.getByTestId(`extension-card-${extensionId}`)).toBeVisible();
      await expect(page.getByTestId(`permission-row-${extensionId}-network`)).toContainText("not granted");

      // First run: allow ui.commands + network so we have something to reset.
      await page.getByTestId("run-command-with-args-sampleHello.fetchText").click();
      await expect(page.getByTestId("input-box")).toBeVisible();
      await page.getByTestId("input-box-field").fill(JSON.stringify([url]));
      await page.getByTestId("input-box-ok").click();

      await expect(page.getByTestId("extension-permission-prompt")).toBeVisible();
      await page.getByTestId("extension-permission-allow").click();
      await expect(page.getByTestId("extension-permission-prompt")).toBeVisible();
      await page.getByTestId("extension-permission-allow").click();

      await expect(page.getByTestId("toast-root")).toContainText("Fetched: hello");
      await expect(page.getByTestId(`permission-row-${extensionId}-network`)).toContainText("mode: allowlist");

      // Reset all extension permissions globally.
      await page.getByTestId("reset-all-extension-permissions").click();
      await expect(page.getByTestId(`permission-row-${extensionId}-network`)).toContainText("not granted");

      // Next run: deny network permission.
      await page.getByTestId("run-command-with-args-sampleHello.fetchText").click();
      await expect(page.getByTestId("input-box")).toBeVisible();
      await page.getByTestId("input-box-field").fill(JSON.stringify([url]));
      await page.getByTestId("input-box-ok").click();

      await expect(page.getByTestId("extension-permission-prompt")).toBeVisible();
      if (await page.getByTestId("extension-permission-ui.commands").isVisible()) {
        await page.getByTestId("extension-permission-allow").click();
        await expect(page.getByTestId("extension-permission-prompt")).toBeVisible();
      }
      await expect(page.getByTestId("extension-permission-network")).toBeVisible();
      await page.getByTestId("extension-permission-deny").click();
      await expect(page.getByTestId("toast-root")).toContainText("Permission denied");
      await expect(page.getByTestId(`permission-row-${extensionId}-network`)).toContainText("not granted");
    } finally {
      await new Promise<void>((resolve) => server.close(() => resolve()));
    }
  });

  test("persists granted permissions across reload", async ({ page }) => {
    const server = http.createServer((_req, res) => {
      res.writeHead(200, {
        "Content-Type": "text/plain",
        "Access-Control-Allow-Origin": "*",
      });
      res.end("hello");
    });

    await new Promise<void>((resolve) => server.listen(0, "127.0.0.1", resolve));
    const address = server.address();
    const port = typeof address === "object" && address ? address.port : null;
    if (!port) throw new Error("Failed to allocate test port");

    const url = `http://127.0.0.1:${port}/`;
    const extensionId = "formula.sample-hello";

    try {
      await gotoDesktop(page);

      // Clear any prior permission grants in this browser context, but do it after boot
      // so we don't clear again on reload.
      await page.evaluate(() => {
        const key = "formula.extensionHost.permissions";
        try {
          // Preserve the built-in e2e-events storage grant so we don't trigger
          // an unrelated permission prompt that blocks UI interactions.
          localStorage.setItem(key, JSON.stringify({ "formula.e2e-events": { storage: true } }));
        } catch {
          // ignore
        }
      });

      await openExtensionsPanel(page);

      await expect(page.getByTestId(`permission-row-${extensionId}-network`)).toContainText("not granted");

      await page.getByTestId("run-command-with-args-sampleHello.fetchText").click();
      await expect(page.getByTestId("input-box")).toBeVisible();
      await page.getByTestId("input-box-field").fill(JSON.stringify([url]));
      await page.getByTestId("input-box-ok").click();

      await expect(page.getByTestId("extension-permission-prompt")).toBeVisible();
      await page.getByTestId("extension-permission-allow").click();
      await expect(page.getByTestId("extension-permission-prompt")).toBeVisible();
      await page.getByTestId("extension-permission-allow").click();

      await expect(page.getByTestId("toast-root")).toContainText("Fetched: hello");
      await expect(page.getByTestId(`permission-row-${extensionId}-network`)).toContainText("127.0.0.1");

      await page.reload({ waitUntil: "domcontentloaded" });
      await waitForDesktopReady(page);

      await openExtensionsPanel(page);

      await expect(page.getByTestId(`extension-card-${extensionId}`)).toBeVisible({ timeout: 30_000 });
      await expect(page.getByTestId(`permission-row-${extensionId}-network`)).toContainText("127.0.0.1");

      // Running again should succeed without prompting (permissions were persisted).
      await page.getByTestId("run-command-with-args-sampleHello.fetchText").click();
      await expect(page.getByTestId("input-box")).toBeVisible();
      await page.getByTestId("input-box-field").fill(JSON.stringify([url]));
      await page.getByTestId("input-box-ok").click();
      await expect(page.getByTestId("toast-root")).toContainText("Fetched: hello");
      await expect(page.getByTestId("extension-permission-prompt")).toHaveCount(0);
    } finally {
      await new Promise<void>((resolve) => server.close(() => resolve()));
    }
  });

  test("revoking cells.read blocks sumSelection", async ({ page }) => {
    const extensionId = "formula.sample-hello";

    await page.addInitScript(() => {
      try {
        localStorage.removeItem("formula.extensionHost.permissions");
      } catch {
        // ignore
      }
    });

      await gotoDesktop(page);

      await page.evaluate(() => {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const app: any = window.__formulaApp as any;
        if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
        const doc = app.getDocument();
        const sheetId = app.getCurrentSheetId();

      doc.setCellValue(sheetId, { row: 0, col: 0 }, 1);
      doc.setCellValue(sheetId, { row: 0, col: 1 }, 2);
      doc.setCellValue(sheetId, { row: 1, col: 0 }, 3);
      doc.setCellValue(sheetId, { row: 1, col: 1 }, 4);

      app.selectRange({
        sheetId,
        range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
      });
    });

    await openExtensionsPanel(page);

    await page.getByTestId("run-command-sampleHello.sumSelection").click();

    await expect(page.getByTestId("extension-permission-prompt")).toBeVisible();
    await expect(page.getByTestId("extension-permission-ui.commands")).toBeVisible();
    await page.getByTestId("extension-permission-allow").click();
    await expect(page.getByTestId("extension-permission-ui.commands")).toHaveCount(0);

    await expect(page.getByTestId("extension-permission-prompt")).toBeVisible();
    await expect(page.getByTestId("extension-permission-cells.read")).toBeVisible();
    await page.getByTestId("extension-permission-allow").click();
    await expect(page.getByTestId("extension-permission-cells.read")).toHaveCount(0);

    await expect(page.getByTestId("extension-permission-prompt")).toBeVisible();
    await expect(page.getByTestId("extension-permission-cells.write")).toBeVisible();
    await page.getByTestId("extension-permission-allow").click();
    await expect(page.getByTestId("extension-permission-cells.write")).toHaveCount(0);

    await expect(page.getByTestId("toast-root")).toContainText("Sum: 10");
    await expect(page.getByTestId(`permission-row-${extensionId}-cells.read`)).toContainText("granted");
    await expect(page.getByTestId(`permission-row-${extensionId}-cells.write`)).toContainText("granted");

    // Revoke only cells.read permission; ensure other grants remain.
    await page.getByTestId(`revoke-permission-${extensionId}-cells.read`).click();
    await expect(page.getByTestId(`permission-row-${extensionId}-cells.read`)).toContainText("not granted");
    await expect(page.getByTestId(`permission-row-${extensionId}-cells.write`)).toContainText("granted");

    await page.getByTestId("run-command-sampleHello.sumSelection").click();

    await expect(page.getByTestId("extension-permission-prompt")).toBeVisible();
    await expect(page.getByTestId("extension-permission-cells.read")).toBeVisible();
    await page.getByTestId("extension-permission-deny").click();
    await expect(page.getByTestId("extension-permission-cells.read")).toHaveCount(0);

    await expect(page.getByTestId("toast-root")).toContainText("Permission denied");
    await expect(page.getByTestId(`permission-row-${extensionId}-cells.read`)).toContainText("not granted");
    await expect(page.getByTestId(`permission-row-${extensionId}-cells.write`)).toContainText("granted");
  });

  test("network allowlist prompts again for a new host", async ({ page }) => {
    // Use Playwright routing to avoid relying on loopback alias configuration (127.0.0.2).
    // We only care about hostname-based allowlisting, not real DNS.
    const urlA = "http://allowed.example/";
    const urlB = "http://blocked.example/";

    await page.route("http://allowed.example/**", async (route) => {
      await route.fulfill({
        status: 200,
        contentType: "text/plain",
        headers: { "Access-Control-Allow-Origin": "*" },
        body: "hello",
      });
    });
    const extensionId = "formula.sample-hello";

    try {
      await page.addInitScript(() => {
        try {
          localStorage.removeItem("formula.extensionHost.permissions");
        } catch {
          // ignore
        }
      });

      await gotoDesktop(page);

      await openExtensionsPanel(page);

      // First run: allow permissions so network allowlist gets seeded for 127.0.0.1.
      await page.getByTestId("run-command-with-args-sampleHello.fetchText").click();
      await expect(page.getByTestId("input-box")).toBeVisible();
      await page.getByTestId("input-box-field").fill(JSON.stringify([urlA]));
      await page.getByTestId("input-box-ok").click();

      await expect(page.getByTestId("extension-permission-prompt")).toBeVisible();
      await expect(page.getByTestId("extension-permission-ui.commands")).toBeVisible();
      await page.getByTestId("extension-permission-allow").click();
      await expect(page.getByTestId("extension-permission-ui.commands")).toHaveCount(0);

      await expect(page.getByTestId("extension-permission-prompt")).toBeVisible();
      await expect(page.getByTestId("extension-permission-network")).toBeVisible();
      await page.getByTestId("extension-permission-allow").click();
      await expect(page.getByTestId("extension-permission-network")).toHaveCount(0);

      await expect(page.getByTestId("toast-root")).toContainText("Fetched: hello");
      await expect(page.getByTestId(`permission-row-${extensionId}-network`)).toContainText("allowed.example");

      // Second run: different host should prompt again because network is allowlisted by hostname.
      await page.getByTestId("run-command-with-args-sampleHello.fetchText").click();
      await expect(page.getByTestId("input-box")).toBeVisible();
      await page.getByTestId("input-box-field").fill(JSON.stringify([urlB]));
      await page.getByTestId("input-box-ok").click();

      await expect(page.getByTestId("extension-permission-prompt")).toBeVisible();
      const networkEntry = page.getByTestId("extension-permission-network");
      await expect(networkEntry).toBeVisible();
      await expect(networkEntry).toContainText("blocked.example");
      await page.getByTestId("extension-permission-deny").click();
      await expect(page.getByTestId("extension-permission-network")).toHaveCount(0);
      await expect(page.getByTestId("toast-root")).toContainText("Permission denied");

      // Permissions UI should still only include the original allowlisted host.
      const networkRow = page.getByTestId(`permission-row-${extensionId}-network`);
      await expect(networkRow).toContainText("allowed.example");
      await expect(networkRow).not.toContainText("blocked.example");
    } finally {
      await page.unroute("http://allowed.example/**").catch(() => {});
    }
  });

  test("revoked permissions persist across reload", async ({ page }) => {
    const url = "http://allowed.example/";
    const extensionId = "formula.sample-hello";

    await page.route("http://allowed.example/**", async (route) => {
      await route.fulfill({
        status: 200,
        contentType: "text/plain",
        headers: { "Access-Control-Allow-Origin": "*" },
        body: "hello",
      });
    });

    try {
      await gotoDesktop(page);

      // Clear any prior grants for this browser context, but do it after boot so reload doesn't re-clear.
      await page.evaluate(() => {
        const key = "formula.extensionHost.permissions";
        try {
          localStorage.setItem(key, JSON.stringify({ "formula.e2e-events": { storage: true } }));
        } catch {
          // ignore
        }
      });

      await openExtensionsPanel(page);

      await expect(page.getByTestId(`extension-card-${extensionId}`)).toBeVisible();
      // When no permissions have been granted yet, the UI should show declared permissions as
      // "not granted" with disabled revoke controls.
      await expect(page.getByTestId(`permission-row-${extensionId}-network`)).toContainText("not granted");
      await expect(page.getByTestId(`revoke-permission-${extensionId}-network`)).toBeDisabled();

      // Grant permissions (ui.commands + network) by running fetchText once.
      await page.getByTestId("run-command-with-args-sampleHello.fetchText").click();
      await expect(page.getByTestId("input-box")).toBeVisible();
      await page.getByTestId("input-box-field").fill(JSON.stringify([url]));
      await page.getByTestId("input-box-ok").click();

      await expect(page.getByTestId("extension-permission-prompt")).toBeVisible();
      await page.getByTestId("extension-permission-allow").click();
      await expect(page.getByTestId("extension-permission-prompt")).toBeVisible();
      await page.getByTestId("extension-permission-allow").click();
      await expect(page.getByTestId("toast-root")).toContainText("Fetched: hello");

      // Revoke only network and ensure it stays revoked after reload.
      await page.getByTestId(`revoke-permission-${extensionId}-network`).click();
      await expect(page.getByTestId(`permission-row-${extensionId}-network`)).toContainText("not granted");

      await page.reload({ waitUntil: "domcontentloaded" });
      await waitForDesktopReady(page);

      await openExtensionsPanel(page);

      await expect(page.getByTestId(`extension-card-${extensionId}`)).toBeVisible({ timeout: 30_000 });
      await expect(page.getByTestId(`permission-row-${extensionId}-network`)).toContainText("not granted");

      // Network should prompt again (ui.commands should remain granted and not be re-requested).
      await page.getByTestId("run-command-with-args-sampleHello.fetchText").click();
      await expect(page.getByTestId("input-box")).toBeVisible();
      await page.getByTestId("input-box-field").fill(JSON.stringify([url]));
      await page.getByTestId("input-box-ok").click();

      await expect(page.getByTestId("extension-permission-prompt")).toBeVisible();
      await expect(page.getByTestId("extension-permission-ui.commands")).toHaveCount(0);
      await expect(page.getByTestId("extension-permission-network")).toBeVisible();
      await page.getByTestId("extension-permission-deny").click();
      await expect(page.getByTestId("toast-root")).toContainText("Permission denied");
      await expect(page.getByTestId(`permission-row-${extensionId}-network`)).toContainText("not granted");
    } finally {
      await page.unroute("http://allowed.example/**").catch(() => {});
    }
  });

  test("network allowlist adds a host when permission is granted", async ({ page }) => {
    const urlA = "http://one.example/";
    const urlB = "http://two.example/";
    const extensionId = "formula.sample-hello";

    await page.route("http://one.example/**", async (route) => {
      await route.fulfill({
        status: 200,
        contentType: "text/plain",
        headers: { "Access-Control-Allow-Origin": "*" },
        body: "one",
      });
    });
    await page.route("http://two.example/**", async (route) => {
      await route.fulfill({
        status: 200,
        contentType: "text/plain",
        headers: { "Access-Control-Allow-Origin": "*" },
        body: "two",
      });
    });

    try {
      await page.addInitScript(() => {
        try {
          localStorage.removeItem("formula.extensionHost.permissions");
        } catch {
          // ignore
        }
      });

      await gotoDesktop(page);

      await openExtensionsPanel(page);

      // First host.
      await page.getByTestId("run-command-with-args-sampleHello.fetchText").click();
      await expect(page.getByTestId("input-box")).toBeVisible();
      await page.getByTestId("input-box-field").fill(JSON.stringify([urlA]));
      await page.getByTestId("input-box-ok").click();

      await expect(page.getByTestId("extension-permission-prompt")).toBeVisible();
      await page.getByTestId("extension-permission-allow").click();
      await expect(page.getByTestId("extension-permission-prompt")).toBeVisible();
      await page.getByTestId("extension-permission-allow").click();

      await expect(page.getByTestId("toast-root")).toContainText("Fetched: one");
      await expect(page.getByTestId(`permission-row-${extensionId}-network`)).toContainText("one.example");

      // Second host should prompt again and, when allowed, expand the allowlist.
      await page.getByTestId("run-command-with-args-sampleHello.fetchText").click();
      await expect(page.getByTestId("input-box")).toBeVisible();
      await page.getByTestId("input-box-field").fill(JSON.stringify([urlB]));
      await page.getByTestId("input-box-ok").click();

      await expect(page.getByTestId("extension-permission-prompt")).toBeVisible();
      await expect(page.getByTestId("extension-permission-ui.commands")).toHaveCount(0);
      const networkEntry = page.getByTestId("extension-permission-network");
      await expect(networkEntry).toBeVisible();
      await expect(networkEntry).toContainText("two.example");
      await page.getByTestId("extension-permission-allow").click();
      await expect(page.getByTestId("extension-permission-network")).toHaveCount(0);

      await expect(page.getByTestId("toast-root")).toContainText("Fetched: two");
      await expect(page.getByTestId(`permission-row-${extensionId}-network`)).toContainText("one.example");
      await expect(page.getByTestId(`permission-row-${extensionId}-network`)).toContainText("two.example");
    } finally {
      await page.unroute("http://one.example/**").catch(() => {});
      await page.unroute("http://two.example/**").catch(() => {});
    }
  });

  test("reset permissions for this extension clears granted permissions", async ({ page }) => {
    const url = "http://allowed.example/";
    const extensionId = "formula.sample-hello";

    await page.route("http://allowed.example/**", async (route) => {
      await route.fulfill({
        status: 200,
        contentType: "text/plain",
        headers: { "Access-Control-Allow-Origin": "*" },
        body: "hello",
      });
    });

    try {
      await page.addInitScript(() => {
        try {
          localStorage.removeItem("formula.extensionHost.permissions");
        } catch {
          // ignore
        }
      });

      await gotoDesktop(page);

      await openExtensionsPanel(page);

      // Grant permissions (ui.commands + network) by running fetchText once.
      await page.getByTestId("run-command-with-args-sampleHello.fetchText").click();
      await expect(page.getByTestId("input-box")).toBeVisible();
      await page.getByTestId("input-box-field").fill(JSON.stringify([url]));
      await page.getByTestId("input-box-ok").click();

      await expect(page.getByTestId("extension-permission-prompt")).toBeVisible();
      await page.getByTestId("extension-permission-allow").click();
      await expect(page.getByTestId("extension-permission-prompt")).toBeVisible();
      await page.getByTestId("extension-permission-allow").click();

      await expect(page.getByTestId("toast-root")).toContainText("Fetched: hello");
      await expect(page.getByTestId(`permission-row-${extensionId}-ui.commands`)).toContainText("granted");
      await expect(page.getByTestId(`permission-row-${extensionId}-network`)).toContainText("allowed.example");

      // Reset permissions for this extension.
      await page.getByTestId(`reset-extension-permissions-${extensionId}`).click();

      await expect(page.getByTestId(`permission-row-${extensionId}-ui.commands`)).toContainText("not granted");
      await expect(page.getByTestId(`permission-row-${extensionId}-network`)).toContainText("not granted");

      // Running again should prompt again.
      await page.getByTestId("run-command-with-args-sampleHello.fetchText").click();
      await expect(page.getByTestId("input-box")).toBeVisible();
      await page.getByTestId("input-box-field").fill(JSON.stringify([url]));
      await page.getByTestId("input-box-ok").click();

      await expect(page.getByTestId("extension-permission-prompt")).toBeVisible();
      if (await page.getByTestId("extension-permission-ui.commands").isVisible()) {
        await page.getByTestId("extension-permission-allow").click();
        await expect(page.getByTestId("extension-permission-prompt")).toBeVisible();
      }
      await expect(page.getByTestId("extension-permission-network")).toBeVisible();
      await page.getByTestId("extension-permission-deny").click();
      await expect(page.getByTestId("toast-root")).toContainText("Permission denied");
      await expect(page.getByTestId(`permission-row-${extensionId}-network`)).toContainText("not granted");
    } finally {
      await page.unroute("http://allowed.example/**").catch(() => {});
    }
  });

  test("revoking clipboard permission blocks copySumToClipboard", async ({ page }) => {
    const extensionId = "formula.sample-hello";

    await page.addInitScript(() => {
      try {
        localStorage.removeItem("formula.extensionHost.permissions");
      } catch {
        // ignore
      }
    });

    await gotoDesktop(page);

    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = window.__formulaApp as any;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();

      doc.setCellValue(sheetId, { row: 0, col: 0 }, 1);
      doc.setCellValue(sheetId, { row: 0, col: 1 }, 2);
      doc.setCellValue(sheetId, { row: 1, col: 0 }, 3);
      doc.setCellValue(sheetId, { row: 1, col: 1 }, 4);

      app.selectRange({
        sheetId,
        range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
      });
    });

    await openExtensionsPanel(page);
    await expect(page.getByTestId("panel-extensions")).toBeVisible();

    await expect(page.getByTestId(`permission-row-${extensionId}-clipboard`)).toContainText("not granted");

    await page.getByTestId("run-command-sampleHello.copySumToClipboard").click();

    await expect(page.getByTestId("extension-permission-prompt")).toBeVisible();
    await expect(page.getByTestId("extension-permission-ui.commands")).toBeVisible();
    await page.getByTestId("extension-permission-allow").click();
    await expect(page.getByTestId("extension-permission-ui.commands")).toHaveCount(0);

    await expect(page.getByTestId("extension-permission-prompt")).toBeVisible();
    await expect(page.getByTestId("extension-permission-cells.read")).toBeVisible();
    await page.getByTestId("extension-permission-allow").click();
    await expect(page.getByTestId("extension-permission-cells.read")).toHaveCount(0);

    await expect(page.getByTestId("extension-permission-prompt")).toBeVisible();
    await expect(page.getByTestId("extension-permission-clipboard")).toBeVisible();
    await page.getByTestId("extension-permission-allow").click();
    await expect(page.getByTestId("extension-permission-clipboard")).toHaveCount(0);

    await expect(page.getByTestId(`permission-row-${extensionId}-clipboard`)).toContainText("granted");

    await page.getByTestId(`revoke-permission-${extensionId}-clipboard`).click();
    await expect(page.getByTestId(`permission-row-${extensionId}-clipboard`)).toContainText("not granted");

    await page.getByTestId("run-command-sampleHello.copySumToClipboard").click();
    await expect(page.getByTestId("extension-permission-prompt")).toBeVisible();
    await expect(page.getByTestId("extension-permission-clipboard")).toBeVisible();
    await page.getByTestId("extension-permission-deny").click();

    await expect(page.getByTestId("toast-root")).toContainText("Permission denied");
    await expect(page.getByTestId(`permission-row-${extensionId}-clipboard`)).toContainText("not granted");
  });
});
