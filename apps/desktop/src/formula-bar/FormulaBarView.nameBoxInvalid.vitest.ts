/**
 * @vitest-environment jsdom
 */

import { describe, expect, it, vi } from "vitest";

import { FormulaBarView } from "./FormulaBarView.js";

describe("FormulaBarView name box invalid reference feedback", () => {
  it("keeps focus + sets aria-invalid when navigation fails, then clears on input and navigates", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onGoTo = vi.fn((reference: string) => reference === "B3");
    new FormulaBarView(host, { onCommit: () => {}, onGoTo });

    const address = host.querySelector<HTMLInputElement>('[data-testid="formula-address"]');
    expect(address).not.toBeNull();

    address!.focus();
    address!.value = "NotARef";
    address!.dispatchEvent(new Event("input", { bubbles: true }));
    address!.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true, cancelable: true }));

    expect(onGoTo).toHaveBeenCalledTimes(1);
    expect(onGoTo).toHaveBeenCalledWith("NotARef");
    expect(address!.getAttribute("aria-invalid")).toBe("true");
    expect(document.activeElement).toBe(address);

    address!.value = "B3";
    address!.dispatchEvent(new Event("input", { bubbles: true }));
    expect(address!.getAttribute("aria-invalid")).not.toBe("true");

    address!.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true, cancelable: true }));

    expect(onGoTo).toHaveBeenCalledTimes(2);
    expect(onGoTo).toHaveBeenLastCalledWith("B3");
    expect(address!.getAttribute("aria-invalid")).not.toBe("true");

    host.remove();
  });

  it("clears invalid state on blur and restores the previous name box value", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onGoTo = vi.fn(() => false);
    new FormulaBarView(host, { onCommit: () => {}, onGoTo });

    const address = host.querySelector<HTMLInputElement>('[data-testid="formula-address"]')!;
    address.focus();
    address.value = "NotARef";
    address.dispatchEvent(new Event("input", { bubbles: true }));
    address.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true, cancelable: true }));

    expect(address.getAttribute("aria-invalid")).toBe("true");
    expect(address.value).toBe("NotARef");

    address.blur();

    expect(address.getAttribute("aria-invalid")).not.toBe("true");
    // `FormulaBarView` keeps the name box in sync with the selection; the default selection is A1.
    expect(address.value).toBe("A1");

    host.remove();
  });
});
