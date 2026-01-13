import { readFileSync } from "node:fs";

import { describe, expect, it } from "vitest";

describe("oauthRedirectIpc wiring", () => {
  it("installs the oauth-redirect IPC readiness handshake in main.ts (prevents cold-start drops)", () => {
    // `main.ts` has many side effects and isn't safe to import in unit tests. Instead, validate
    // (lightly) that it wires the OAuth redirect listener and emits `oauth-redirect-ready`
    // *after* the listener is registered. The Rust host queues `oauth-redirect` URLs until this
    // handshake occurs; breaking it can drop deep-link redirects on cold start.
    const mainUrl = new URL("../../main.ts", import.meta.url);
    const source = readFileSync(mainUrl, "utf8");
    const code = source
      // Strip block comments so commented-out wiring can't satisfy the guardrail.
      .replace(/\/\*[\s\S]*?\*\//g, "")
      // Strip full-line `// ...` comments; we intentionally do *not* attempt to remove
      // inline `//` comments because `main.ts` contains many `https://...` strings.
      .split("\n")
      .filter((line) => !line.trimStart().startsWith("//"))
      .join("\n");

    // 1) Ensure we listen for Rust -> JS oauth redirect events.
    const hasOauthRedirectListener =
      /^\s*(?:const|let)\s+\w+(?:\s*:\s*[^=]+)?\s*=\s*listen\s*\(\s*["']oauth-redirect["']/m.test(code) ||
      /^\s*(?:void\s+)?listen\s*\(\s*["']oauth-redirect["']/m.test(code);
    expect(hasOauthRedirectListener).toBe(true);

    // 2) Ensure we only emit readiness after the listener promise resolves.
    // Accept `const x = listen("oauth-redirect", ...); x.then(() => emit("oauth-redirect-ready"))`
    // and close equivalents (e.g. `await x; emit(...)`).
    const emitReadyMatches = Array.from(code.matchAll(/\bemit\s*\(\s*["']oauth-redirect-ready["']/g));
    expect(emitReadyMatches).toHaveLength(1);

    const escapeForRegExp = (value: string): string => value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const assignment =
      /^(?:\s*)(?:const|let)\s+(\w+)(?:\s*:\s*[^=]+)?\s*=\s*listen\s*\(\s*["']oauth-redirect["']/m.exec(code);
    if (assignment) {
      const varName = escapeForRegExp(assignment[1]!);
      const hasThen =
        new RegExp(
          String.raw`\b${varName}\b\s*\.\s*then\s*\(\s*(?:async\s*)?(?:\(\s*\)\s*=>|function\s*\(\s*\))[\s\S]{0,750}?\bemit\s*\(\s*["']oauth-redirect-ready["']`,
          "m",
        ).test(code);
      const hasAwait =
        new RegExp(
          String.raw`await\s+\b${varName}\b[\s\S]{0,750}?\bemit\s*\(\s*["']oauth-redirect-ready["']`,
          "m",
        ).test(code);
      expect(hasThen || hasAwait).toBe(true);
    } else {
      // Fallback for inlined promise chaining / top-level await.
      const hasThen =
        /\blisten\s*\(\s*["']oauth-redirect["'][\s\S]{0,750}?\)\s*\.?\s*then\s*\(\s*(?:async\s*)?(?:\(\s*\)\s*=>|function\s*\(\s*\))[\s\S]{0,750}?\bemit\s*\(\s*["']oauth-redirect-ready["']/.test(
          code,
        );
      const hasAwait =
        /await\s+listen\s*\(\s*["']oauth-redirect["'][\s\S]{0,750}?\)\s*;?[\s\S]{0,750}?\bemit\s*\(\s*["']oauth-redirect-ready["']/.test(
          code,
        );
      expect(hasThen || hasAwait).toBe(true);
    }

    // 3) Ensure the listener forwards the payload into the OAuth broker.
    expect(code).toMatch(
      /\blisten\s*\(\s*["']oauth-redirect["']\s*,[\s\S]{0,200}?(?:=>|function)[\s\S]{0,1000}?\boauthBroker\.observeRedirect\s*\(/,
    );
  });
});
