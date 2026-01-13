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
  it("uses an HTTPS Authenticode timestamp server on Windows", () => {
    const config = loadTauriConfig();
    const timestampUrl = config?.bundle?.windows?.timestampUrl as unknown;
    expect(typeof timestampUrl, "expected bundle.windows.timestampUrl to be a string").toBe("string");
    const parsed = new URL(String(timestampUrl));
    expect(parsed.protocol).toBe("https:");
  });

  it("configures Windows installers to bootstrap the WebView2 runtime when missing", () => {
    const config = loadTauriConfig();
    const mode = config?.bundle?.windows?.webviewInstallMode as unknown;
    expect(mode, "expected bundle.windows.webviewInstallMode to be set (do not rely on WebView2 being preinstalled)").toBeTruthy();

    if (typeof mode === "string") {
      expect(
        mode.toLowerCase(),
        "bundle.windows.webviewInstallMode must not be 'skip' (installer must install WebView2 on fresh machines)",
      ).not.toBe("skip");
      return;
    }

    expect(mode && typeof mode === "object", "expected bundle.windows.webviewInstallMode to be a string or object").toBe(
      true,
    );
    const type = String((mode as any)?.type ?? "").trim();
    expect(type.length > 0, "expected bundle.windows.webviewInstallMode.type to be a non-empty string").toBe(true);
    expect(
      type.toLowerCase(),
      "bundle.windows.webviewInstallMode.type must not be 'skip' (installer must install WebView2 on fresh machines)",
    ).not.toBe("skip");
  });

  it("allows manual Windows downgrades (rollback via Releases page)", () => {
    const config = loadTauriConfig();
    expect(config?.bundle?.windows?.allowDowngrades).toBe(true);
  });

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

  it("enforces a hardened CSP (no framing, no plugins, restricted outbound network from WebView)", () => {
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
    // The desktop CSP allows outbound HTTPS + WebSockets (collaboration/remote APIs) while still
    // disallowing plaintext HTTP and wildcard networking.
    expect(connectSrc).toContain("'self'");
    expect(connectSrc).toContain("https:");
    expect(connectSrc).toContain("ws:");
    expect(connectSrc).toContain("wss:");
    expect(connectSrc).toContain("blob:");
    expect(connectSrc).toContain("data:");

    const forbiddenConnectSrcTokens = connectSrc.filter((token) => {
      if (token === "*") return true;
      // `https:`/`ws:`/`wss:` are allowed; disallow only plaintext `http:`.
      return token.startsWith("http:");
    });
    expect(
      forbiddenConnectSrcTokens,
      `connect-src must not allow wildcard/plaintext HTTP networking; found: ${forbiddenConnectSrcTokens.join(", ")}`,
    ).toEqual([]);

    // Capabilities are always scoped in the capability file via `"windows": [...]` (window labels).
    //
    // Some toolchains also support window-level opt-in via `app.windows[].capabilities` in `tauri.conf.json`. When present,
    // ensure it does not accidentally grant the main capability to new windows.
    const windows = Array.isArray(config?.app?.windows) ? (config.app.windows as Array<Record<string, unknown>>) : [];
    const mainWindow = windows.find((w) => String((w as any)?.label ?? "") === "main") as any;
    expect(mainWindow).toBeTruthy();

    const mainCaps = (mainWindow as any)?.capabilities as unknown;
    if (mainCaps != null) {
      expect(Array.isArray(mainCaps)).toBe(true);
      expect(mainCaps).toContain("main");

      for (const window of windows) {
        const label = String((window as any)?.label ?? "");
        const caps = (window as any)?.capabilities as unknown;
        if (caps == null) continue;
        expect(Array.isArray(caps)).toBe(true);
        if (label !== "main") expect(caps).not.toContain("main");
      }
    } else {
      for (const window of windows) expect((window as any)?.capabilities).toBeUndefined();
    }

    const capUrl = new URL("../../../src-tauri/capabilities/main.json", import.meta.url);
    const cap = JSON.parse(readFileSync(capUrl, "utf8")) as any;
    expect(cap?.identifier).toBe("main");
    expect(Array.isArray(cap?.windows)).toBe(true);
    expect(cap.windows).toContain("main");
    // Avoid wildcard/pattern scoping that could accidentally grant this capability to new windows.
    for (const label of cap.windows as unknown[]) {
      expect(typeof label).toBe("string");
      expect(label).not.toContain("*");
    }
  });
});
