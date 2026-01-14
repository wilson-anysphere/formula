// @vitest-environment jsdom

import { describe, expect, it } from "vitest";

import { createDesktopPermissionPrompt } from "./permissionPrompt.js";

describe("createDesktopPermissionPrompt", () => {
  it("falls back to a non-modal open attribute when showModal throws", async () => {
    document.body.innerHTML = "";

    const proto = (globalThis as any).HTMLDialogElement?.prototype as any;
    expect(proto).toBeTruthy();

    const originalShowModal = proto.showModal;
    proto.showModal = function showModal() {
      throw new Error("showModal failed");
    };

    try {
      const prompt = createDesktopPermissionPrompt();
      const resultPromise = prompt({
        extensionId: "acme.example",
        displayName: "Acme Extension",
        permissions: ["network"],
        request: { network: { host: "api.example.com", url: "https://api.example.com/v1/hello" } },
      });

      const dialog = document.querySelector<HTMLDialogElement>('dialog[data-testid="extension-permission-prompt"]');
      expect(dialog).not.toBeNull();
      expect(dialog?.hasAttribute("open")).toBe(true);

      dialog?.querySelector<HTMLButtonElement>('button[data-testid="extension-permission-deny"]')?.click();
      await expect(resultPromise).resolves.toBe(false);
    } finally {
      proto.showModal = originalShowModal;
    }
  });

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
    const title = dialog?.querySelector<HTMLElement>(".dialog__title");
    expect(title).not.toBeNull();
    const titleId = title?.id ?? "";
    expect(titleId).not.toEqual("");
    expect(dialog?.getAttribute("aria-labelledby")).toBe(titleId);
    expect(document.getElementById(titleId)).toBe(title);

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

  it("waits for an existing open dialog to close before showing the prompt", async () => {
    document.body.innerHTML = "";

    const blocking = document.createElement("dialog");
    blocking.setAttribute("open", "");
    document.body.appendChild(blocking);

    const prompt = createDesktopPermissionPrompt();
    const resultPromise = prompt({
      extensionId: "acme.example",
      displayName: "Acme Extension",
      permissions: ["network"],
      request: { network: { host: "api.example.com" } },
    });

    // The prompt should not show while another dialog is open.
    expect(document.querySelector('dialog[data-testid="extension-permission-prompt"]')).toBeNull();

    // Close/remove the blocking dialog.
    blocking.removeAttribute("open");
    blocking.dispatchEvent(new Event("close"));
    blocking.remove();

    // Wait a tick for the prompt to render.
    for (let i = 0; i < 10; i += 1) {
      const dialog = document.querySelector<HTMLDialogElement>('dialog[data-testid="extension-permission-prompt"]');
      if (dialog) break;
      // eslint-disable-next-line no-await-in-loop
      await new Promise((resolve) => setTimeout(resolve, 0));
    }

    const dialog = document.querySelector<HTMLDialogElement>('dialog[data-testid="extension-permission-prompt"]');
    expect(dialog).not.toBeNull();

    dialog?.querySelector<HTMLButtonElement>('button[data-testid="extension-permission-deny"]')?.click();
    await expect(resultPromise).resolves.toBe(false);
  });
});
