import { expect, test } from "@playwright/test";
import { spawn, type ChildProcess } from "node:child_process";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { gotoDesktop } from "./helpers";

function stablePortFromString(input: string, { base = 23_000, range = 2000 } = {}): number {
  // Deterministic port selection avoids collisions when multiple agent checkouts run
  // Playwright tests on the same host.
  let hash = 0;
  for (let i = 0; i < input.length; i += 1) {
    hash = (hash * 31 + input.charCodeAt(i)) >>> 0;
  }
  return base + (hash % range);
}

async function waitForHealthz(url: string, timeoutMs = 30_000): Promise<void> {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    try {
      const res = await fetch(url);
      if (res.ok) return;
    } catch {
      // ignore
    }
    await new Promise((resolve) => setTimeout(resolve, 200));
  }
  throw new Error(`Timed out waiting for sync server health check at ${url}`);
}

test.describe("collab status indicator (collab mode)", () => {
  test.describe.configure({ mode: "serial" });

  let syncServer: ChildProcess | null = null;
  let wsUrl: string;

  test.beforeAll(async () => {
    const here = path.dirname(fileURLToPath(import.meta.url));
    const desktopRoot = path.resolve(here, "../..");
    const repoRoot = path.resolve(desktopRoot, "../..");
    const syncServerRoot = path.resolve(repoRoot, "services/sync-server");

    const port = stablePortFromString(repoRoot);
    wsUrl = `ws://127.0.0.1:${port}`;

    const dataDir = path.join(os.tmpdir(), `formula-sync-server-e2e-${port}`);

    syncServer = spawn(process.execPath, ["scripts/node-with-tsx.mjs", "src/index.ts"], {
      cwd: syncServerRoot,
      env: {
        ...process.env,
        NODE_ENV: "test",
        SYNC_SERVER_HOST: "127.0.0.1",
        SYNC_SERVER_PORT: String(port),
        // Avoid optional leveldb deps in e2e environments.
        SYNC_SERVER_PERSISTENCE_BACKEND: "file",
        SYNC_SERVER_DATA_DIR: dataDir,
      },
      stdio: "inherit",
    });

    await waitForHealthz(`http://127.0.0.1:${port}/healthz`);
  });

  test.afterAll(async () => {
    if (syncServer) {
      syncServer.kill("SIGTERM");
      syncServer = null;
    }
  });

  test("shows Synced after connecting to the sync server", async ({ page }) => {
    test.setTimeout(120_000);

    const page2 = await page.context().newPage();
    const docId = `e2e-doc-${Date.now()}`;
    const token = "dev-token";

    const urlForUser = (userId: string): string => {
      const params = new URLSearchParams({
        collab: "1",
        docId,
        wsUrl,
        token,
        userId,
        userName: userId,
      });
      return `/?${params.toString()}`;
    };

    await Promise.all([gotoDesktop(page, urlForUser("user1")), gotoDesktop(page2, urlForUser("user2"))]);

    await expect(page.getByTestId("collab-status")).toContainText(docId);
    await expect(page2.getByTestId("collab-status")).toContainText(docId);

    await expect(page.getByTestId("collab-status")).toContainText("Synced", { timeout: 30_000 });
    await expect(page2.getByTestId("collab-status")).toContainText("Synced", { timeout: 30_000 });
  });
});
