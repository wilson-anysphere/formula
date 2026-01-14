// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { showInputBox, showQuickPick } from "./ui.js";

beforeEach(() => {
  document.body.innerHTML = "";
});

afterEach(() => {
  document.body.innerHTML = "";
  vi.restoreAllMocks();
});

describe("extensions/ui showInputBox", () => {
  it("resolves with input value when OK is clicked", async () => {
    const promise = showInputBox({ prompt: "Name", value: "default" });
    const dialog = document.querySelector('dialog[data-testid="input-box"]') as HTMLDialogElement | null;
    expect(dialog).not.toBeNull();

    const input = dialog!.querySelector('[data-testid="input-box-field"]') as HTMLInputElement | null;
    expect(input).not.toBeNull();
    input!.value = "hello";

    const ok = dialog!.querySelector('[data-testid="input-box-ok"]') as HTMLButtonElement | null;
    expect(ok).not.toBeNull();
    ok!.click();

    await expect(promise).resolves.toBe("hello");
  });

  it("resolves with null when Cancel is clicked", async () => {
    const promise = showInputBox({ prompt: "Name", value: "default" });
    const dialog = document.querySelector('dialog[data-testid="input-box"]') as HTMLDialogElement | null;
    expect(dialog).not.toBeNull();

    const cancel = dialog!.querySelector('[data-testid="input-box-cancel"]') as HTMLButtonElement | null;
    expect(cancel).not.toBeNull();
    cancel!.click();

    await expect(promise).resolves.toBeNull();
  });

  it("supports password mode", async () => {
    const promise = showInputBox({ prompt: "Password", value: "", type: "password" });
    const dialog = document.querySelector('dialog[data-testid="input-box"]') as HTMLDialogElement | null;
    expect(dialog).not.toBeNull();

    const input = dialog!.querySelector('[data-testid="input-box-field"]') as HTMLInputElement | null;
    expect(input).not.toBeNull();
    expect(input!.tagName).toBe("INPUT");
    expect(input!.type).toBe("password");

    input!.value = "secret";

    const ok = dialog!.querySelector('[data-testid="input-box-ok"]') as HTMLButtonElement | null;
    ok!.click();

    await expect(promise).resolves.toBe("secret");
  });

  it("supports textarea mode (Ctrl+Enter submits)", async () => {
    const promise = showInputBox({ prompt: "JSON", value: "{}", type: "textarea" });
    const dialog = document.querySelector('dialog[data-testid="input-box"]') as HTMLDialogElement | null;
    expect(dialog).not.toBeNull();

    const textarea = dialog!.querySelector('[data-testid="input-box-field"]') as HTMLTextAreaElement | null;
    expect(textarea).not.toBeNull();
    expect(textarea!.tagName).toBe("TEXTAREA");

    textarea!.value = '{"a": 1}';
    textarea!.dispatchEvent(
      new KeyboardEvent("keydown", {
        key: "Enter",
        ctrlKey: true,
        bubbles: true,
        cancelable: true,
      }),
    );

    await expect(promise).resolves.toBe('{"a": 1}');
  });

  it("still works when HTMLDialogElement.showModal throws (best-effort fallback)", async () => {
    const proto = (globalThis as any).HTMLDialogElement?.prototype;
    const original = proto?.showModal;
    if (!proto) throw new Error("Missing HTMLDialogElement prototype");

    try {
      proto.showModal = () => {
        throw new Error("boom");
      };

      const promise = showInputBox({ prompt: "Name", value: "default" });
      const dialog = document.querySelector('dialog[data-testid="input-box"]') as HTMLDialogElement | null;
      expect(dialog).not.toBeNull();

      const input = dialog!.querySelector('[data-testid="input-box-field"]') as HTMLInputElement | null;
      expect(input).not.toBeNull();
      input!.value = "hello";

      const ok = dialog!.querySelector('[data-testid="input-box-ok"]') as HTMLButtonElement | null;
      expect(ok).not.toBeNull();
      ok!.click();

      await expect(promise).resolves.toBe("hello");
    } finally {
      if (original) proto.showModal = original;
      else delete proto.showModal;
      vi.restoreAllMocks();
    }
  });

  it("resolves to null when another modal dialog is already open", async () => {
    const blocking = document.createElement("dialog");
    blocking.setAttribute("open", "");
    document.body.appendChild(blocking);

    await expect(showInputBox({ prompt: "Name" })).resolves.toBeNull();

    // Ensure we did not create an additional input-box dialog.
    expect(document.querySelectorAll('dialog[data-testid="input-box"]').length).toBe(0);
  });
});

describe("extensions/ui showQuickPick", () => {
  it("resolves with the selected value when an item is clicked", async () => {
    const promise = showQuickPick(
      [
        { label: "One", value: 1 },
        { label: "Two", value: 2 },
      ],
      { placeHolder: "Pick" },
    );

    const dialog = document.querySelector('dialog[data-testid="quick-pick"]') as HTMLDialogElement | null;
    expect(dialog).not.toBeNull();

    const item = dialog!.querySelector('[data-testid="quick-pick-item-1"]') as HTMLButtonElement | null;
    expect(item).not.toBeNull();
    item!.click();

    await expect(promise).resolves.toBe(2);
  });

  it("resolves to null when Escape is pressed", async () => {
    const promise = showQuickPick([{ label: "One", value: 1 }], { placeHolder: "Pick" });
    const dialog = document.querySelector('dialog[data-testid="quick-pick"]') as HTMLDialogElement | null;
    expect(dialog).not.toBeNull();

    dialog!.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true, cancelable: true }));
    await expect(promise).resolves.toBeNull();
  });

  it("resolves to null when another modal dialog is already open", async () => {
    const blocking = document.createElement("dialog");
    blocking.setAttribute("open", "");
    document.body.appendChild(blocking);

    await expect(showQuickPick([{ label: "One", value: 1 }], { placeHolder: "Pick" })).resolves.toBeNull();
    expect(document.querySelectorAll('dialog[data-testid="quick-pick"]').length).toBe(0);
  });

  it("still works when HTMLDialogElement.showModal throws (best-effort fallback)", async () => {
    const proto = (globalThis as any).HTMLDialogElement?.prototype;
    const original = proto?.showModal;
    if (!proto) throw new Error("Missing HTMLDialogElement prototype");

    try {
      proto.showModal = () => {
        throw new Error("boom");
      };

      const promise = showQuickPick([{ label: "One", value: 1 }], { placeHolder: "Pick" });
      const dialog = document.querySelector('dialog[data-testid="quick-pick"]') as HTMLDialogElement | null;
      expect(dialog).not.toBeNull();

      const item = dialog!.querySelector('[data-testid="quick-pick-item-0"]') as HTMLButtonElement | null;
      expect(item).not.toBeNull();
      item!.click();

      await expect(promise).resolves.toBe(1);
    } finally {
      if (original) proto.showModal = original;
      else delete proto.showModal;
    }
  });
});
