// @vitest-environment jsdom
import React, { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it, vi } from "vitest";

import { t } from "../../i18n/index.js";

import { CollabVersionHistoryPanel } from "./CollabVersionHistoryPanel.js";

const confirmMock = vi.fn(async () => true);

vi.mock("../../tauri/nativeDialogs.js", () => ({
  confirm: (...args: any[]) => confirmMock(...args),
}));

vi.mock("../useCollabSessionSyncState.js", () => ({
  useCollabSessionSyncState: () => ({ connected: true }),
}));

vi.mock("../collabReservedRootGuard.react.js", () => ({
  clearReservedRootGuardError: () => {},
  useReservedRootGuardError: () => null,
}));

const createCollabVersioningForPanelMock = vi.fn();
vi.mock("./createCollabVersioningForPanel.js", () => ({
  createCollabVersioningForPanel: (...args: any[]) => createCollabVersioningForPanelMock(...args),
}));

function flushPromises() {
  return new Promise<void>((resolve) => setTimeout(resolve, 0));
}

function setNativeValue(el: HTMLInputElement | HTMLTextAreaElement, value: string): void {
  const proto = el instanceof HTMLInputElement ? window.HTMLInputElement.prototype : window.HTMLTextAreaElement.prototype;
  const setter = Object.getOwnPropertyDescriptor(proto, "value")?.set;
  if (setter) {
    setter.call(el, value);
  } else {
    // Best-effort fallback.
    (el as any).value = value;
  }
}

function setNativeChecked(el: HTMLInputElement, checked: boolean): void {
  const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, "checked")?.set;
  if (setter) {
    setter.call(el, checked);
  } else {
    // Best-effort fallback.
    (el as any).checked = checked;
  }
}

async function waitFor(condition: () => boolean, timeoutMs = 2_000) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    if (condition()) return;
    // eslint-disable-next-line no-await-in-loop
    await act(async () => {
      await flushPromises();
    });
  }
  throw new Error("Timed out waiting for condition");
}

afterEach(() => {
  document.body.innerHTML = "";
  vi.restoreAllMocks();
});

describe("CollabVersionHistoryPanel named checkpoints", () => {
  it("creates checkpoints with annotations + locked, supports lock toggle, and blocks deletion while locked", async () => {
    (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

    let versions: any[] = [];
    let idCounter = 0;
    confirmMock.mockClear();

    const collabVersioning = {
      listVersions: vi.fn(async () => versions.map((v) => ({ ...v }))),
      createCheckpoint: vi.fn(async (args: { name: string; annotations?: string; locked?: boolean }) => {
        idCounter += 1;
        const id = `ckpt_${idCounter}`;
        versions = [
          ...versions,
          {
            id,
            kind: "checkpoint",
            timestampMs: Date.now(),
            checkpointName: args.name,
            checkpointAnnotations: args.annotations ?? "",
            checkpointLocked: Boolean(args.locked),
          },
        ];
        return { id };
      }),
      setCheckpointLocked: vi.fn(async (id: string, locked: boolean) => {
        versions = versions.map((v) => (v.id === id ? { ...v, checkpointLocked: locked } : v));
      }),
      deleteVersion: vi.fn(async (id: string) => {
        versions = versions.filter((v) => v.id !== id);
      }),
      restoreVersion: vi.fn(async () => {}),
      destroy: vi.fn(),
    };

    createCollabVersioningForPanelMock.mockResolvedValue(collabVersioning);

    const container = document.createElement("div");
    document.body.appendChild(container);
    const root = createRoot(container);
    act(() => {
      root.render(React.createElement(CollabVersionHistoryPanel, { session: { provider: {} } as any }));
    });

    await waitFor(() => Boolean(container.querySelector(".collab-version-history__create")));

    const nameInput = container.querySelector<HTMLInputElement>(".collab-version-history__input");
    const annotations = container.querySelector<HTMLTextAreaElement>(".collab-version-history__textarea");
    const lockedCheckbox = container.querySelector<HTMLInputElement>('.collab-version-history__checkbox input[type="checkbox"]');
    expect(nameInput).toBeInstanceOf(HTMLInputElement);
    expect(annotations).toBeInstanceOf(HTMLTextAreaElement);
    expect(lockedCheckbox).toBeInstanceOf(HTMLInputElement);

    act(() => {
      setNativeValue(nameInput!, "  My checkpoint  ");
      nameInput!.dispatchEvent(new Event("input", { bubbles: true }));
    });
    act(() => {
      setNativeValue(annotations!, "  Some notes  ");
      annotations!.dispatchEvent(new Event("input", { bubbles: true }));
    });
    act(() => {
      // React's checkbox onChange is wired to click events; use `.click()` so jsdom
      // performs the default toggle behavior and React observes the change.
      lockedCheckbox!.click();
    });

    await act(async () => {
      await flushPromises();
    });

    const createButton = Array.from(container.querySelectorAll("button")).find(
      (b) => b.textContent === t("versionHistory.actions.createCheckpoint"),
    );
    expect(createButton).toBeInstanceOf(HTMLButtonElement);
    expect((createButton as HTMLButtonElement).disabled).toBe(false);

    await act(async () => {
      createButton!.click();
      await flushPromises();
    });

    expect(collabVersioning.createCheckpoint).toHaveBeenCalledTimes(1);
    expect(collabVersioning.createCheckpoint).toHaveBeenCalledWith({
      name: "My checkpoint",
      annotations: "Some notes",
      locked: true,
    });

    await waitFor(() => Boolean(container.querySelector(".collab-version-history__item")));

    const deleteButton = Array.from(container.querySelectorAll("button")).find(
      (b) => b.textContent === t("versionHistory.actions.deleteSelected"),
    );
    expect(deleteButton).toBeInstanceOf(HTMLButtonElement);
    expect((deleteButton as HTMLButtonElement | null)?.disabled).toBe(true);

    const unlockButton = Array.from(container.querySelectorAll("button")).find(
      (b) => b.textContent === t("versionHistory.actions.unlock"),
    );
    expect(unlockButton).toBeInstanceOf(HTMLButtonElement);

    await act(async () => {
      unlockButton!.click();
      await flushPromises();
    });

    expect(collabVersioning.setCheckpointLocked).toHaveBeenCalledWith("ckpt_1", false);

    await waitFor(() => {
      const updated = Array.from(container.querySelectorAll("button")).find(
        (b) => b.textContent === t("versionHistory.actions.deleteSelected"),
      ) as HTMLButtonElement | undefined;
      return Boolean(updated && !updated.disabled);
    });

    const deleteButtonAfterUnlock = Array.from(container.querySelectorAll("button")).find(
      (b) => b.textContent === t("versionHistory.actions.deleteSelected"),
    ) as HTMLButtonElement | undefined;
    expect(deleteButtonAfterUnlock).toBeInstanceOf(HTMLButtonElement);

    await act(async () => {
      deleteButtonAfterUnlock!.click();
      await flushPromises();
    });

    expect(confirmMock).toHaveBeenCalledWith(t("versionHistory.confirm.deleteIrreversible"));
    expect(collabVersioning.deleteVersion).toHaveBeenCalledWith("ckpt_1");
    await waitFor(() => (container.textContent?.includes(t("versionHistory.panel.empty")) ?? false));

    act(() => root.unmount());
  });
});
