import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";

function loadTauriConfig(): any {
  const tauriConfUrl = new URL("../../../src-tauri/tauri.conf.json", import.meta.url);
  return JSON.parse(readFileSync(tauriConfUrl, "utf8")) as any;
}

function parseCsp(csp: string): Map<string, string[]> {
  const directives = new Map<string, string[]>();
  for (const rawPart of csp.split(";")) {
    const part = rawPart.trim();
    if (!part) continue;

    const [rawName, ...rawSources] = part.split(/\s+/);
    const name = rawName.toLowerCase();
    const sources = rawSources.map((source) => source.toLowerCase());
    directives.set(name, sources);
  }
  return directives;
}

function requireCspDirective(directives: Map<string, string[]>, directive: string): string[] {
  const sources = directives.get(directive.toLowerCase());
  expect(sources, `missing ${directive} directive in app.security.csp`).not.toBeUndefined();
  return sources ?? [];
}

function isFalsyOrEmpty(value: unknown): boolean {
  if (!value) return true;
  if (Array.isArray(value)) return value.length === 0;
  if (typeof value === "object") return Object.keys(value as Record<string, unknown>).length === 0;
  return false;
}

describe("tauri.conf.json security guardrails", () => {
  it("enables COOP/COEP headers for cross-origin isolation (SharedArrayBuffer)", () => {
    const config = loadTauriConfig();
    const headers = config?.app?.security?.headers as unknown;
    expect(headers && typeof headers === "object", "expected app.security.headers to be present").toBe(true);

    expect((headers as any)["Cross-Origin-Opener-Policy"]).toBe("same-origin");
    expect((headers as any)["Cross-Origin-Embedder-Policy"]).toBe("require-corp");
  });

  it("does not enable dangerous Tauri remote IPC flags", () => {
    const config = loadTauriConfig();
    const security = config?.app?.security as any;
    expect(security && typeof security === "object", "expected app.security to be an object").toBe(true);

    const allowedDangerousKeys = new Set<string>(["dangerousRemoteDomainIpcAccess"]);

    if (Object.prototype.hasOwnProperty.call(security, "dangerousRemoteDomainIpcAccess")) {
      const value = security.dangerousRemoteDomainIpcAccess as unknown;
      expect(
        isFalsyOrEmpty(value),
        "app.security.dangerousRemoteDomainIpcAccess must be unset/empty (remote content must not reach IPC)",
      ).toBe(true);
    }

    const otherDangerousKeys = Object.keys(security).filter(
      (key) => key.toLowerCase().includes("dangerous") && !allowedDangerousKeys.has(key),
    );
    expect(
      otherDangerousKeys,
      `Unexpected "dangerous" keys in app.security: ${otherDangerousKeys.join(", ")}`,
    ).toEqual([]);
  });

  it("enforces a hardened CSP (no framing, no plugins, no outbound network from WebView)", () => {
    const config = loadTauriConfig();
    const csp = config?.app?.security?.csp as unknown;
    expect(typeof csp, "expected app.security.csp to be a string").toBe("string");

    const directives = parseCsp(csp as string);

    const frameAncestors = requireCspDirective(directives, "frame-ancestors");
    expect(frameAncestors).toContain("'none'");

    const objectSrc = requireCspDirective(directives, "object-src");
    expect(objectSrc).toContain("'none'");

    const defaultSrc = requireCspDirective(directives, "default-src");
    expect(defaultSrc).toContain("'self'");

    const connectSrc = requireCspDirective(directives, "connect-src");
    const forbiddenConnectSrcTokens = connectSrc.filter((token) => {
      if (token === "*") return true;
      return (
        token.startsWith("http:") ||
        token.startsWith("https:") ||
        token.startsWith("ws:") ||
        token.startsWith("wss:")
      );
    });
    expect(
      forbiddenConnectSrcTokens,
      `connect-src must not allow remote networking; found: ${forbiddenConnectSrcTokens.join(", ")}`,
    ).toEqual([]);
  });
});

