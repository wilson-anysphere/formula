import { describe, expect, it } from "vitest";

import { injectWebviewCsp } from "./ExtensionPanelBody.js";

describe("injectWebviewCsp", () => {
  it("injects CSP + hardening script inside <head> before any extension head content", () => {
    const html = "<!doctype html><html><head><title>Test</title></head><body><h1>Hi</h1></body></html>";
    const out = injectWebviewCsp(html);

    const cspIdx = out.indexOf('<meta http-equiv="Content-Security-Policy"');
    expect(cspIdx).toBeGreaterThanOrEqual(0);

    const hardeningIdx = out.indexOf("__formulaWebviewSandbox");
    expect(hardeningIdx).toBeGreaterThanOrEqual(0);

    const titleIdx = out.indexOf("<title>Test</title>");
    expect(cspIdx).toBeLessThan(titleIdx);
    expect(hardeningIdx).toBeLessThan(titleIdx);
  });

  it("injects before the first <script> tag when <script> appears before <head>", () => {
    const html = "<!doctype html><html><script>console.log('x')</script><head></head><body></body></html>";
    const out = injectWebviewCsp(html);

    const cspIdx = out.indexOf('<meta http-equiv="Content-Security-Policy"');
    const firstScriptIdx = out.indexOf("<script>console.log('x')</script>");
    expect(cspIdx).toBeGreaterThanOrEqual(0);
    expect(firstScriptIdx).toBeGreaterThanOrEqual(0);
    expect(cspIdx).toBeLessThan(firstScriptIdx);
  });

  it("injects before tags that appear before <html>", () => {
    const html = "<img src=\"data:,\"><html><head></head><body></body></html>";
    const out = injectWebviewCsp(html);

    const cspIdx = out.indexOf('<meta http-equiv="Content-Security-Policy"');
    const imgIdx = out.indexOf("<img");
    expect(cspIdx).toBeGreaterThanOrEqual(0);
    expect(imgIdx).toBeGreaterThanOrEqual(0);
    expect(cspIdx).toBeLessThan(imgIdx);
  });

  it("wraps arbitrary markup in a full document and injects CSP + hardening", () => {
    const html = "<h1>Hello</h1>";
    const out = injectWebviewCsp(html);

    expect(out).toContain("<!doctype html>");
    expect(out).toContain('<meta http-equiv="Content-Security-Policy"');
    expect(out).toContain("__formulaWebviewSandbox");
    expect(out).toContain("<body>");
    expect(out).toContain("<h1>Hello</h1>");
  });

  it("injects before scripts even when the markup is a fragment starting with <script>", () => {
    const html = "<script>window.__TAURI__ = 1;</script><h1>Late</h1>";
    const out = injectWebviewCsp(html);

    const cspIdx = out.indexOf('<meta http-equiv="Content-Security-Policy"');
    const originalScriptIdx = out.indexOf("<script>window.__TAURI__ = 1;</script>");
    expect(cspIdx).toBeGreaterThanOrEqual(0);
    expect(originalScriptIdx).toBeGreaterThanOrEqual(0);
    expect(cspIdx).toBeLessThan(originalScriptIdx);
  });
});
