/**
 * @vitest-environment jsdom
 */

import { describe, expect, it } from "vitest";

import { pickLocalImageFiles } from "../pickLocalImageFiles.js";

describe("pickLocalImageFiles (web <input> fallback)", () => {
  it("resolves with selected File objects and cleans up the temporary input", async () => {
    // Ensure no Tauri global leaks into this test so we exercise the web fallback.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (globalThis as any).__TAURI__ = undefined;

    const promise = pickLocalImageFiles({ multiple: true });

    const input = document.querySelector<HTMLInputElement>('input[type="file"]');
    expect(input).toBeTruthy();

    const fileA = new File([new Uint8Array([1])], "a.png", { type: "image/png" });
    const fileB = new File([new Uint8Array([2])], "b.jpg", { type: "image/jpeg" });
    Object.defineProperty(input!, "files", { value: [fileA, fileB] });
    input!.dispatchEvent(new Event("change"));

    const result = await promise;
    expect(result.map((f) => f.name)).toEqual(["a.png", "b.jpg"]);
    expect(document.querySelector('input[type="file"]')).toBeNull();
  });
});

