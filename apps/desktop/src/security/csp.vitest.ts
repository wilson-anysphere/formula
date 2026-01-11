import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

function parseCspDirective(csp: string, directive: string): string[] | null {
  const parts = csp
    .split(";")
    .map((part) => part.trim())
    .filter(Boolean);

  for (const part of parts) {
    if (part === directive) {
      return [];
    }
    if (part.startsWith(`${directive} `)) {
      return part
        .slice(directive.length)
        .trim()
        .split(/\s+/)
        .filter(Boolean);
    }
  }

  return null;
}

describe("Tauri CSP", () => {
  it("allows WASM compilation and module Workers (required for @formula/engine)", () => {
    const tauriConfigPath = fileURLToPath(new URL("../../src-tauri/tauri.conf.json", import.meta.url));
    const config = JSON.parse(readFileSync(tauriConfigPath, "utf8")) as any;
    const csp = config?.app?.security?.csp as unknown;
    expect(typeof csp).toBe("string");

    const scriptSrc = parseCspDirective(csp as string, "script-src");
    expect(scriptSrc, "missing script-src directive").not.toBeNull();
    expect(scriptSrc).toContain("'wasm-unsafe-eval'");

    const workerSrc = parseCspDirective(csp as string, "worker-src");
    expect(workerSrc, "missing worker-src directive").not.toBeNull();
    expect(workerSrc).toContain("'self'");
    expect(workerSrc).toContain("blob:");

    // Older WebKit versions (macOS 10.15 WKWebView / WebKitGTK) can gate Workers
    // behind `child-src` instead of `worker-src`.
    const childSrc = parseCspDirective(csp as string, "child-src");
    expect(childSrc, "missing child-src directive").not.toBeNull();
    expect(childSrc).toContain("'self'");
    expect(childSrc).toContain("blob:");
  });
});
