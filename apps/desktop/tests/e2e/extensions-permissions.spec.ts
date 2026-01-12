import { expect, test } from "@playwright/test";
import http from "node:http";

import { gotoDesktop } from "./helpers";

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
      await page.addInitScript(({ extensionId }) => {
        // Seed a granted network permission so the first fetch succeeds without prompting.
        const key = "formula.extensionHost.permissions";
        const store = { [extensionId]: { network: { mode: "full" } } };
        localStorage.setItem(key, JSON.stringify(store));

        // Deny future network permission prompts (used after revocation below).
        (window as any).__formulaPermissionPrompt = ({ permissions }: { permissions?: string[] }) => {
          return !Array.isArray(permissions) || !permissions.includes("network");
        };
      }, { extensionId });

      await gotoDesktop(page);

      await page.getByTestId("open-extensions-panel").click();
      await expect(page.getByTestId("panel-extensions")).toBeVisible();

      await expect(page.getByTestId(`extension-card-${extensionId}`)).toBeVisible();
      await expect(page.getByTestId(`permission-${extensionId}-network`)).toContainText("mode: full");

      const firstFetch = await page.evaluate(
        async ({ url }) => {
          const mgr = (window as any).__formulaExtensionHostManager;
          if (!mgr) throw new Error("Missing window.__formulaExtensionHostManager");
          return await mgr.executeCommand("sampleHello.fetchText", url);
        },
        { url },
      );
      expect(firstFetch).toBe("hello");

      await page.getByTestId(`revoke-all-permissions-${extensionId}`).click();
      await expect(page.getByTestId(`permissions-empty-${extensionId}`)).toBeVisible();

      const secondFetch = await page.evaluate(
        async ({ url }) => {
          const mgr = (window as any).__formulaExtensionHostManager;
          if (!mgr) throw new Error("Missing window.__formulaExtensionHostManager");
          try {
            await mgr.executeCommand("sampleHello.fetchText", url);
            return { ok: true, errorMessage: "" };
          } catch (err: any) {
            return { ok: false, errorMessage: String(err?.message ?? err) };
          }
        },
        { url },
      );

      expect(secondFetch.ok).toBe(false);
      expect(secondFetch.errorMessage).toContain("Permission denied");
    } finally {
      await new Promise<void>((resolve) => server.close(() => resolve()));
    }
  });
});

