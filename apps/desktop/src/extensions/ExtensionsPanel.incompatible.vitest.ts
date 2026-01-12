/**
 * @vitest-environment jsdom
 */

import React, { act } from "react";
import { createRoot } from "react-dom/client";
import { describe, expect, it, vi } from "vitest";

import { ExtensionsPanel } from "./ExtensionsPanel";

// Suppress React 18 act() warnings in Vitest/jsdom.
// https://react.dev/reference/react/act#act-in-tests
(globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

async function waitFor(condition: () => boolean, { timeoutMs = 2000, intervalMs = 10 } = {}) {
  const start = Date.now();
  for (;;) {
    if (condition()) return;
    if (Date.now() - start > timeoutMs) {
      throw new Error("Timed out waiting for condition");
    }
    // eslint-disable-next-line no-await-in-loop
    await new Promise((resolve) => setTimeout(resolve, intervalMs));
  }
}

describe("ExtensionsPanel (IndexedDB installs)", () => {
  it("renders incompatible installs + uses update() action instead of repair()", async () => {
    const rootEl = document.createElement("div");
    document.body.appendChild(rootEl);
    const root = createRoot(rootEl);

    const listeners = new Set<() => void>();

    const manager = {
      ready: true,
      error: null,
      subscribe: (listener: () => void) => {
        listeners.add(listener);
        return () => listeners.delete(listener);
      },
      loadBuiltInExtensions: vi.fn(async () => {}),
      host: { listExtensions: () => [] },
      getContributedCommands: () => [],
      getContributedPanels: () => [],
      getContributedKeybindings: () => [],
      getGrantedPermissions: vi.fn(async () => ({})),
      revokePermission: vi.fn(async () => {}),
      resetPermissionsForExtension: vi.fn(async () => {}),
      resetAllPermissions: vi.fn(async () => {}),
    } as any;

    let installedList: any[] = [
      {
        id: "test.incompatible-ext",
        version: "1.0.0",
        incompatible: true,
        incompatibleReason: "engine mismatch: need ^2.0.0",
      },
    ];

    const webExtensionManager = {
      verifyAllInstalled: vi.fn(async () => ({})),
      listInstalled: vi.fn(async () => installedList),
      repair: vi.fn(async () => {}),
      update: vi.fn(async (id: string) => {
        installedList = [{ id, version: "1.0.1" }];
        return installedList[0];
      }),
      loadInstalled: vi.fn(async () => {}),
    } as any;

    await act(async () => {
      root.render(
        React.createElement(ExtensionsPanel, {
          manager,
          webExtensionManager,
          onSyncExtensions: vi.fn(),
          onExecuteCommand: vi.fn(),
          onOpenPanel: vi.fn(),
        }),
      );
    });

    // Allow effects (refreshInstalled) to run.
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 0));
    });

    const status = rootEl.querySelector(
      '[data-testid="installed-extension-status-test.incompatible-ext"]',
    ) as HTMLDivElement | null;
    expect(status).toBeTruthy();
    expect(status?.textContent).toContain("Incompatible: engine mismatch");

    const actionButton = rootEl.querySelector(
      '[data-testid="repair-extension-test.incompatible-ext"]',
    ) as HTMLButtonElement | null;
    expect(actionButton).toBeTruthy();
    expect(actionButton?.textContent).toContain("Update");

    await act(async () => {
      actionButton?.click();
    });
    await waitFor(() => webExtensionManager.update.mock.calls.length > 0);
    // Flush state updates triggered by the async click handler.
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 0));
    });
    expect(webExtensionManager.update).toHaveBeenCalledWith("test.incompatible-ext");
    expect(webExtensionManager.repair).not.toHaveBeenCalled();
    expect(webExtensionManager.loadInstalled).toHaveBeenCalledWith("test.incompatible-ext");

    const statusAfter = rootEl.querySelector(
      '[data-testid="installed-extension-status-test.incompatible-ext"]',
    ) as HTMLDivElement | null;
    expect(statusAfter).toBeTruthy();
    expect(statusAfter?.textContent).toBe("OK");

    await act(async () => {
      root.unmount();
    });
  });
});
