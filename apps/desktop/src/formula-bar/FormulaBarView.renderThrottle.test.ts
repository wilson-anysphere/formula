/**
 * @vitest-environment jsdom
 */

import { afterEach, describe, expect, it, vi } from "vitest";

import { FormulaBarView } from "./FormulaBarView.js";

describe("FormulaBarView render throttling", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it("coalesces multiple synchronous input events into one highlight render", () => {
    const rafCallbacks: FrameRequestCallback[] = [];

    const raf = vi.fn((cb: FrameRequestCallback) => {
      rafCallbacks.push(cb);
      return rafCallbacks.length;
    });

    vi.stubGlobal("requestAnimationFrame", raf);
    vi.stubGlobal("cancelAnimationFrame", vi.fn());

    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });
    view.focus({ cursor: "end" });

    // Flush any render scheduled by focusing/beginning edit.
    while (rafCallbacks.length) {
      const cb = rafCallbacks.shift();
      cb?.(0);
    }
    raf.mockClear();

    const highlightedSpy = vi.spyOn(view.model, "highlightedSpans");

    view.textarea.value = "=SUM(A1";
    view.textarea.setSelectionRange(view.textarea.value.length, view.textarea.value.length);
    view.textarea.dispatchEvent(new Event("input"));

    view.textarea.value = "=SUM(A1)";
    view.textarea.setSelectionRange(view.textarea.value.length, view.textarea.value.length);
    view.textarea.dispatchEvent(new Event("input"));

    view.textarea.value = "=SUM(A1)+1";
    view.textarea.setSelectionRange(view.textarea.value.length, view.textarea.value.length);
    view.textarea.dispatchEvent(new Event("input"));

    // All three events should schedule a single render on the next frame.
    expect(raf).toHaveBeenCalledTimes(1);
    expect(highlightedSpy).toHaveBeenCalledTimes(0);

    while (rafCallbacks.length) {
      const cb = rafCallbacks.shift();
      cb?.(0);
    }

    expect(highlightedSpy).toHaveBeenCalledTimes(1);

    host.remove();
  });
});

