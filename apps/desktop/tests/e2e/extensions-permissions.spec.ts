import { expect, test } from "@playwright/test";
import http from "node:http";

import { gotoDesktop, waitForDesktopReady } from "./helpers";

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

      await page.getByTestId("open-extensions-panel").click();
      await expect(page.getByTestId("panel-extensions")).toBeVisible();

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

      await page.getByTestId("open-extensions-panel").click();
      await expect(page.getByTestId("panel-extensions")).toBeVisible();

      await expect(page.getByTestId(`extension-card-${extensionId}`)).toBeVisible();
      await expect(page.getByTestId(`permissions-empty-${extensionId}`)).toBeVisible();

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
      await expect(page.getByTestId(`permission-${extensionId}-network`)).toBeVisible();

      // Reset all extension permissions globally.
      await page.getByTestId("reset-all-extension-permissions").click();
      await expect(page.getByTestId(`permissions-empty-${extensionId}`)).toBeVisible();

      // Next run: deny network permission.
      await page.getByTestId("run-command-with-args-sampleHello.fetchText").click();
      await expect(page.getByTestId("input-box")).toBeVisible();
      await page.getByTestId("input-box-field").fill(JSON.stringify([url]));
      await page.getByTestId("input-box-ok").click();

      await expect(page.getByTestId("extension-permission-prompt")).toBeVisible();
      await expect(page.getByTestId("extension-permission-network")).toBeVisible();
      await page.getByTestId("extension-permission-deny").click();
      await expect(page.getByTestId("toast-root")).toContainText("Permission denied");
      await expect(page.getByTestId(`permissions-empty-${extensionId}`)).toBeVisible();
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
        try {
          localStorage.removeItem("formula.extensionHost.permissions");
        } catch {
          // ignore
        }
      });

      await page.getByTestId("open-extensions-panel").click();
      await expect(page.getByTestId("panel-extensions")).toBeVisible();

      await expect(page.getByTestId(`permissions-empty-${extensionId}`)).toBeVisible();

      await page.getByTestId("run-command-with-args-sampleHello.fetchText").click();
      await expect(page.getByTestId("input-box")).toBeVisible();
      await page.getByTestId("input-box-field").fill(JSON.stringify([url]));
      await page.getByTestId("input-box-ok").click();

      await expect(page.getByTestId("extension-permission-prompt")).toBeVisible();
      await page.getByTestId("extension-permission-allow").click();
      await expect(page.getByTestId("extension-permission-prompt")).toBeVisible();
      await page.getByTestId("extension-permission-allow").click();

      await expect(page.getByTestId("toast-root")).toContainText("Fetched: hello");
      await expect(page.getByTestId(`permission-${extensionId}-network`)).toContainText("127.0.0.1");

      await page.reload();
      await waitForDesktopReady(page);

      await page.getByTestId("open-extensions-panel").click();
      await expect(page.getByTestId("panel-extensions")).toBeVisible();

      await expect(page.getByTestId(`permission-${extensionId}-network`)).toContainText("127.0.0.1");

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
      const app: any = (window as any).__formulaApp;
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

    await page.getByTestId("open-extensions-panel").click();
    await expect(page.getByTestId("panel-extensions")).toBeVisible();

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
    await expect(page.getByTestId(`permission-${extensionId}-cells.read`)).toBeVisible();
    await expect(page.getByTestId(`permission-${extensionId}-cells.write`)).toBeVisible();

    // Revoke only cells.read permission; ensure other grants remain.
    await page.getByTestId(`revoke-permission-${extensionId}-cells.read`).click();
    await expect(page.getByTestId(`permission-${extensionId}-cells.read`)).toHaveCount(0);
    await expect(page.getByTestId(`permission-${extensionId}-cells.write`)).toBeVisible();

    await page.getByTestId("run-command-sampleHello.sumSelection").click();

    await expect(page.getByTestId("extension-permission-prompt")).toBeVisible();
    await expect(page.getByTestId("extension-permission-cells.read")).toBeVisible();
    await page.getByTestId("extension-permission-deny").click();
    await expect(page.getByTestId("extension-permission-cells.read")).toHaveCount(0);

    await expect(page.getByTestId("toast-root")).toContainText("Permission denied");
    await expect(page.getByTestId(`permission-${extensionId}-cells.read`)).toHaveCount(0);
    await expect(page.getByTestId(`permission-${extensionId}-cells.write`)).toBeVisible();
  });

  test("network allowlist prompts again for a new host", async ({ page }) => {
    const serverA = http.createServer((_req, res) => {
      res.writeHead(200, {
        "Content-Type": "text/plain",
        "Access-Control-Allow-Origin": "*",
      });
      res.end("hello-a");
    });

    const serverB = http.createServer((_req, res) => {
      res.writeHead(200, {
        "Content-Type": "text/plain",
        "Access-Control-Allow-Origin": "*",
      });
      res.end("hello-b");
    });

    await new Promise<void>((resolve) => serverA.listen(0, "127.0.0.1", resolve));
    await new Promise<void>((resolve) => serverB.listen(0, "127.0.0.2", resolve));

    const addressA = serverA.address();
    const portA = typeof addressA === "object" && addressA ? addressA.port : null;
    if (!portA) throw new Error("Failed to allocate test port");

    const addressB = serverB.address();
    const portB = typeof addressB === "object" && addressB ? addressB.port : null;
    if (!portB) throw new Error("Failed to allocate test port");

    const urlA = `http://127.0.0.1:${portA}/`;
    const urlB = `http://127.0.0.2:${portB}/`;
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

      await page.getByTestId("open-extensions-panel").click();
      await expect(page.getByTestId("panel-extensions")).toBeVisible();

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

      await expect(page.getByTestId("toast-root")).toContainText("Fetched: hello-a");
      await expect(page.getByTestId(`permission-${extensionId}-network`)).toContainText("127.0.0.1");

      // Second run: different host (127.0.0.2) should prompt again because network is allowlisted by hostname.
      await page.getByTestId("run-command-with-args-sampleHello.fetchText").click();
      await expect(page.getByTestId("input-box")).toBeVisible();
      await page.getByTestId("input-box-field").fill(JSON.stringify([urlB]));
      await page.getByTestId("input-box-ok").click();

      await expect(page.getByTestId("extension-permission-prompt")).toBeVisible();
      const networkEntry = page.getByTestId("extension-permission-network");
      await expect(networkEntry).toBeVisible();
      await expect(networkEntry).toContainText("127.0.0.2");
      await page.getByTestId("extension-permission-deny").click();
      await expect(page.getByTestId("extension-permission-network")).toHaveCount(0);
      await expect(page.getByTestId("toast-root")).toContainText("Permission denied");

      // Permissions UI should still only include the original allowlisted host.
      await expect(page.getByTestId(`permission-${extensionId}-network`)).toContainText("127.0.0.1");
    } finally {
      await new Promise<void>((resolve) => serverA.close(() => resolve()));
      await new Promise<void>((resolve) => serverB.close(() => resolve()));
    }
  });
});
