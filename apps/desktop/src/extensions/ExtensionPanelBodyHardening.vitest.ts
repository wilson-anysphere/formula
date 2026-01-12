import { describe, expect, it } from "vitest";

import { injectWebviewCsp } from "./ExtensionPanelBody.js";

function extractHardeningScriptSource(documentHtml: string): string {
  const match = documentHtml.match(/<script[^>]+src="(data:text\/javascript[^"]+)"[^>]*><\/script>/i);
  if (!match) throw new Error("Missing hardening script tag");

  const url = match[1]!;
  const comma = url.indexOf(",");
  if (comma === -1) throw new Error("Invalid hardening script data URL");

  return decodeURIComponent(url.slice(comma + 1));
}

describe("ExtensionPanelBody hardening script", () => {
  it("scrubs non-configurable Tauri globals and locks them down", async () => {
    const html = injectWebviewCsp("<!doctype html><html><head></head><body></body></html>");
    const scriptSource = extractHardeningScriptSource(html);

    const listeners: Record<string, Array<() => void>> = {};
    const timeouts: Array<{ delay: number; callback: () => void }> = [];

    const fakeWindow: any = {
      addEventListener(type: string, cb: () => void) {
        (listeners[type] ||= []).push(cb);
      },
    };
    const fakeDocument: any = {
      addEventListener(type: string, cb: () => void) {
        (listeners[type] ||= []).push(cb);
      },
    };
    const fakeSetTimeout = (cb: () => void, delay?: number) => {
      timeouts.push({ delay: Number(delay ?? 0), callback: cb });
      return timeouts.length;
    };

    Object.defineProperty(fakeWindow, "__TAURI__", {
      value: { secret: "x" },
      writable: true,
      configurable: false,
      enumerable: true,
    });

    const run = new Function("window", "document", "setTimeout", "Promise", scriptSource) as (
      win: any,
      doc: any,
      st: any,
      prom: any,
    ) => void;
    run(fakeWindow, fakeDocument, fakeSetTimeout, Promise);
    await Promise.resolve();

    expect(fakeWindow.__formulaWebviewSandbox).toBeDefined();
    expect(fakeWindow.__formulaWebviewSandbox.tauriGlobalsPresent).toBe(true);

    expect(typeof fakeWindow.__TAURI__).toBe("undefined");
    const desc = Object.getOwnPropertyDescriptor(fakeWindow, "__TAURI__") as any;
    expect(desc).toBeTruthy();
    expect(desc.value).toBeUndefined();
    expect(desc.writable).toBe(false);
    expect(desc.configurable).toBe(false);
    expect(desc.enumerable).toBe(true);
  });

  it("scrubs globals injected after initial evaluation (via delayed scrub passes)", async () => {
    const html = injectWebviewCsp("<!doctype html><html><head></head><body></body></html>");
    const scriptSource = extractHardeningScriptSource(html);

    const listeners: Record<string, Array<() => void>> = {};
    const timeouts: Array<{ delay: number; callback: () => void }> = [];

    const fakeWindow: any = {
      addEventListener(type: string, cb: () => void) {
        (listeners[type] ||= []).push(cb);
      },
    };
    const fakeDocument: any = {
      addEventListener(type: string, cb: () => void) {
        (listeners[type] ||= []).push(cb);
      },
    };
    const fakeSetTimeout = (cb: () => void, delay?: number) => {
      timeouts.push({ delay: Number(delay ?? 0), callback: cb });
      return timeouts.length;
    };

    const run = new Function("window", "document", "setTimeout", "Promise", scriptSource) as (
      win: any,
      doc: any,
      st: any,
      prom: any,
    ) => void;
    run(fakeWindow, fakeDocument, fakeSetTimeout, Promise);
    await Promise.resolve();

    expect(fakeWindow.__formulaWebviewSandbox).toBeDefined();
    expect(fakeWindow.__formulaWebviewSandbox.tauriGlobalsPresent).toBe(false);

    // Run early timeouts (0ms / 50ms / 250ms) before injecting globals, simulating a runtime that
    // injects Tauri globals after initial load.
    const sorted = [...timeouts].sort((a, b) => a.delay - b.delay);
    for (const t of sorted) {
      if (t.delay <= 250) t.callback();
    }
    await Promise.resolve();
    expect(fakeWindow.__formulaWebviewSandbox.tauriGlobalsPresent).toBe(false);

    fakeWindow.__TAURI__ = { injected: true };

    // Run remaining delayed timeouts (including the 1000ms pass) and simulate load events.
    for (const t of sorted) {
      if (t.delay > 250) t.callback();
    }
    for (const cb of listeners.load ?? []) cb();
    await Promise.resolve();

    expect(fakeWindow.__formulaWebviewSandbox.tauriGlobalsPresent).toBe(true);
    expect(typeof fakeWindow.__TAURI__).toBe("undefined");
  });
});

