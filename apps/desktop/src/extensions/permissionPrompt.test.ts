// @vitest-environment jsdom

import { describe, expect, it } from "vitest";

import { createDesktopPermissionPrompt } from "./permissionPrompt.js";

describe("createDesktopPermissionPrompt", () => {
  it("renders a dialog with permission details and resolves false when denied", async () => {
    document.body.innerHTML = "";

    const prompt = createDesktopPermissionPrompt();
    const resultPromise = prompt({
      extensionId: "acme.example",
      displayName: "Acme Extension",
      permissions: ["network", "clipboard"],
      request: { network: { host: "api.example.com", url: "https://api.example.com/v1/hello" } },
    });

    const dialog = document.querySelector<HTMLDialogElement>('dialog[data-testid="extension-permission-prompt"]');
    expect(dialog).not.toBeNull();
    expect(dialog?.textContent).toContain("Acme Extension");
    expect(dialog?.textContent).toContain("acme.example");
    expect(dialog?.textContent).toContain("network");
    expect(dialog?.textContent).toContain("clipboard");
    expect(dialog?.textContent).toContain("api.example.com");

    const deny = dialog?.querySelector<HTMLButtonElement>('button[data-testid="extension-permission-deny"]');
    expect(deny).not.toBeNull();
    deny?.click();

    await expect(resultPromise).resolves.toBe(false);
    expect(document.querySelector('dialog[data-testid="extension-permission-prompt"]')).toBeNull();
  });

  it("serializes concurrent prompts so only one dialog is shown at a time", async () => {
    document.body.innerHTML = "";

    const prompt = createDesktopPermissionPrompt();

    const first = prompt({
      extensionId: "a.one",
      displayName: "First",
      permissions: ["ui.commands"],
    });

    const second = prompt({
      extensionId: "b.two",
      displayName: "Second",
      permissions: ["network"],
      request: { network: { host: "example.com", url: "https://example.com/" } },
    });

    const firstDialog = document.querySelector<HTMLDialogElement>('dialog[data-testid="extension-permission-prompt"]');
    expect(firstDialog).not.toBeNull();
    expect(firstDialog?.textContent).toContain("First");

    firstDialog?.querySelector<HTMLButtonElement>('button[data-testid="extension-permission-allow"]')?.click();
    await expect(first).resolves.toBe(true);

    // The second prompt should render after the first one closes.
    for (let i = 0; i < 10; i += 1) {
      const maybeDialog = document.querySelector<HTMLDialogElement>('dialog[data-testid="extension-permission-prompt"]');
      if (maybeDialog && maybeDialog.textContent?.includes("Second")) break;
      // eslint-disable-next-line no-await-in-loop
      await new Promise((resolve) => setTimeout(resolve, 0));
    }

    const secondDialog = document.querySelector<HTMLDialogElement>('dialog[data-testid="extension-permission-prompt"]');
    expect(secondDialog).not.toBeNull();
    expect(secondDialog?.textContent).toContain("Second");

    secondDialog?.querySelector<HTMLButtonElement>('button[data-testid="extension-permission-deny"]')?.click();
    await expect(second).resolves.toBe(false);
  });
});

