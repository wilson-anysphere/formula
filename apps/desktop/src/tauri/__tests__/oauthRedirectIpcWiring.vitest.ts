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

    // 1) Ensure we listen for Rust -> JS oauth redirect events.
    expect(source).toMatch(/\blisten\s*\(\s*["']oauth-redirect["']/);

    // 2) Ensure we only emit readiness after the listener promise resolves.
    // Accept `const x = listen("oauth-redirect", ...); x.then(() => emit("oauth-redirect-ready"))`
    // and close equivalents (e.g. `await x; emit(...)`).
    const escapeForRegExp = (value: string): string => value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const assignment = /(?:const|let)\s+(\w+)\s*=\s*listen\s*\(\s*["']oauth-redirect["']/.exec(source);
    if (assignment) {
      const varName = escapeForRegExp(assignment[1]!);
      expect(source).toMatch(
        new RegExp(
          String.raw`\b${varName}\b\s*\.\s*then\s*\(\s*(?:async\s*)?\(\s*\)\s*=>[\s\S]{0,750}?\bemit\s*\(\s*["']oauth-redirect-ready["']`,
          "m",
        ),
      );
    } else {
      // Fallback for inlined promise chaining.
      expect(source).toMatch(
        /\blisten\s*\(\s*["']oauth-redirect["'][\s\S]{0,750}?\)\s*\.?\s*then\s*\(\s*(?:async\s*)?\(\s*\)\s*=>[\s\S]{0,750}?\bemit\s*\(\s*["']oauth-redirect-ready["']/,
      );
    }

    // 3) Ensure the listener forwards the payload into the OAuth broker.
    expect(source).toMatch(/\blisten\s*\(\s*["']oauth-redirect["'][\s\S]{0,1000}?\boauthBroker\.observeRedirect\s*\(/);
  });
});

