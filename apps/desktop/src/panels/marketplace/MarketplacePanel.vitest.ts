// @vitest-environment jsdom

import { afterEach, describe, expect, it, vi } from "vitest";

import { createMarketplacePanel } from "./MarketplacePanel.js";

function flushPromises() {
  return new Promise<void>((resolve) => setTimeout(resolve, 0));
}

async function waitFor(condition: () => boolean, timeoutMs = 2_000) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    if (condition()) return;
    // eslint-disable-next-line no-await-in-loop
    await flushPromises();
  }
  throw new Error("Timed out waiting for condition");
}

afterEach(() => {
  document.body.innerHTML = "";
  vi.restoreAllMocks();
});

describe("MarketplacePanel", () => {
  it("shows install warnings via toast after a successful install", async () => {
    const toastRoot = document.createElement("div");
    toastRoot.id = "toast-root";
    document.body.appendChild(toastRoot);

    const container = document.createElement("div");
    document.body.appendChild(container);

    const marketplaceClient = {
      search: vi.fn(async () => ({
        total: 1,
        results: [
          {
            id: "formula.sample-hello",
            name: "sample-hello",
            displayName: "Sample Hello",
            publisher: "formula",
            description: "hello",
            latestVersion: "1.0.0",
            verified: true,
            featured: false,
          },
        ],
        nextCursor: null,
      })),
      getExtension: vi.fn(async (id: string) => ({
        id,
        latestVersion: "1.0.0",
        verified: true,
        featured: false,
        deprecated: false,
        blocked: false,
        malicious: false,
        versions: [{ version: "1.0.0", scanStatus: "passed" }],
      })),
    };

    let installedRecord: any = null;
    const extensionManager = {
      getInstalled: vi.fn(async (id: string) => (installedRecord?.id === id ? installedRecord : null)),
      install: vi.fn(async (id: string) => {
        installedRecord = {
          id,
          version: "1.0.0",
          installedAt: new Date().toISOString(),
          warnings: [{ kind: "deprecated", message: "Deprecated extension", scanStatus: null }],
        };
        return installedRecord;
      }),
      uninstall: vi.fn(async (_id: string) => {
        installedRecord = null;
      }),
      checkForUpdates: vi.fn(async () => []),
      update: vi.fn(async (id: string) => installedRecord ?? { id, version: "1.0.0", installedAt: new Date().toISOString() }),
    };

    createMarketplacePanel({ container, marketplaceClient: marketplaceClient as any, extensionManager: extensionManager as any });

    const searchInput = container.querySelector<HTMLInputElement>('input[type="search"]');
    expect(searchInput).toBeInstanceOf(HTMLInputElement);
    searchInput!.value = "sample";

    const searchButton = Array.from(container.querySelectorAll("button")).find((b) => b.textContent === "Search");
    expect(searchButton).toBeInstanceOf(HTMLButtonElement);
    searchButton!.click();

    await waitFor(() => Boolean(container.querySelector(".marketplace-result")));
    expect(Array.from(container.querySelectorAll(".marketplace-badge")).map((el) => el.textContent)).toContain("verified");

    const installButton = Array.from(container.querySelectorAll("button")).find((b) => b.textContent === "Install");
    expect(installButton).toBeInstanceOf(HTMLButtonElement);
    installButton!.click();

    await waitFor(() => container.textContent?.includes("Installed") ?? false);
    await waitFor(() => Boolean(document.querySelector('[data-testid="toast"][data-type="warning"]')));

    const toast = document.querySelector<HTMLElement>('[data-testid="toast"][data-type="warning"]');
    expect(toast?.textContent).toContain("Deprecated");
  });

  it("surfaces install cancellation errors via toast when confirm() rejects", async () => {
    vi.spyOn(console, "error").mockImplementation(() => {});

    const toastRoot = document.createElement("div");
    toastRoot.id = "toast-root";
    document.body.appendChild(toastRoot);

    // jsdom's default confirm() may be unimplemented; override it.
    // eslint-disable-next-line no-global-assign
    (window as any).confirm = vi.fn(() => false);

    const container = document.createElement("div");
    document.body.appendChild(container);

    const marketplaceClient = {
      search: vi.fn(async () => ({
        total: 1,
        results: [
          {
            id: "formula.sample-hello",
            name: "sample-hello",
            displayName: "Sample Hello",
            publisher: "formula",
            description: "hello",
            latestVersion: "1.0.0",
            verified: true,
            featured: false,
          },
        ],
        nextCursor: null,
      })),
      getExtension: vi.fn(async (id: string) => ({
        id,
        latestVersion: "1.0.0",
        verified: true,
        featured: false,
        deprecated: true,
        blocked: false,
        malicious: false,
        versions: [{ version: "1.0.0", scanStatus: "passed" }],
      })),
    };

    const extensionManager = {
      getInstalled: vi.fn(async (_id: string) => null),
      install: vi.fn(async (id: string, _version?: any, options?: any) => {
        const warning = { kind: "deprecated", message: "Deprecated extension", scanStatus: null };
        if (options?.confirm) {
          const ok = await options.confirm(warning);
          if (!ok) throw new Error("Extension install cancelled");
        }
        return { id, version: "1.0.0", installedAt: new Date().toISOString(), warnings: [warning] };
      }),
      uninstall: vi.fn(async () => {}),
      checkForUpdates: vi.fn(async () => []),
      update: vi.fn(async () => {
        throw new Error("not implemented");
      }),
    };

    createMarketplacePanel({ container, marketplaceClient: marketplaceClient as any, extensionManager: extensionManager as any });

    const searchInput = container.querySelector<HTMLInputElement>('input[type="search"]');
    expect(searchInput).toBeInstanceOf(HTMLInputElement);
    searchInput!.value = "sample";

    const searchButton = Array.from(container.querySelectorAll("button")).find((b) => b.textContent === "Search");
    expect(searchButton).toBeInstanceOf(HTMLButtonElement);
    searchButton!.click();

    await waitFor(() => Boolean(container.querySelector(".marketplace-result")));

    const installButton = Array.from(container.querySelectorAll("button")).find((b) => b.textContent === "Install");
    expect(installButton).toBeInstanceOf(HTMLButtonElement);
    installButton!.click();

    await waitFor(() => Boolean(document.querySelector('[data-testid="toast"][data-type="error"]')));
    const toast = document.querySelector<HTMLElement>('[data-testid="toast"][data-type="error"]');
    expect(toast?.textContent).toContain("cancelled");
    await waitFor(() => container.textContent?.toLowerCase().includes("error") ?? false);
  });
});
