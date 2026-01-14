import { describe, expect, it } from "vitest";

import { isLoopbackRedirectUrl, matchesRedirectUri } from "./oauthBroker.js";

describe("matchesRedirectUri", () => {
  it("matches a custom-scheme deep link redirect (ignoring query params)", () => {
    expect(matchesRedirectUri("formula://oauth/callback", "formula://oauth/callback?code=abc&state=123")).toBe(true);
  });

  it("rejects a deep link with the wrong host", () => {
    expect(matchesRedirectUri("formula://oauth/callback", "formula://not-oauth/callback?code=abc")).toBe(false);
  });

  it("rejects a deep link with the wrong path", () => {
    expect(matchesRedirectUri("formula://oauth/callback", "formula://oauth/not-callback?code=abc")).toBe(false);
  });

  it("matches a loopback redirect (including port) while ignoring query params", () => {
    expect(matchesRedirectUri("http://127.0.0.1:4242/callback", "http://127.0.0.1:4242/callback?code=abc")).toBe(true);
  });

  it("matches a localhost loopback redirect (including port) while ignoring query params", () => {
    expect(matchesRedirectUri("http://localhost:4242/callback", "http://localhost:4242/callback?code=abc")).toBe(true);
  });

  it("matches an IPv6 loopback redirect (including port) while ignoring query params", () => {
    expect(matchesRedirectUri("http://[::1]:4242/callback", "http://[::1]:4242/callback?code=abc")).toBe(true);
  });

  it("rejects a loopback redirect with the wrong port", () => {
    expect(matchesRedirectUri("http://127.0.0.1:4242/callback", "http://127.0.0.1:5555/callback?code=abc")).toBe(false);
  });

  it("returns false for invalid URLs", () => {
    expect(matchesRedirectUri("not a url", "formula://oauth/callback?code=abc")).toBe(false);
    expect(matchesRedirectUri("formula://oauth/callback", "not a url")).toBe(false);
  });
});

describe("isLoopbackRedirectUrl", () => {
  it("rejects loopback URLs with fragments (not observable by HTTP servers)", () => {
    expect(isLoopbackRedirectUrl(new URL("http://127.0.0.1:4242/callback#access_token=abc"))).toBe(false);
    expect(isLoopbackRedirectUrl(new URL("http://[::1]:4242/callback#access_token=abc"))).toBe(false);
  });

  it("rejects loopback URLs with userinfo", () => {
    expect(isLoopbackRedirectUrl(new URL("http://user@127.0.0.1:4242/callback"))).toBe(false);
    expect(isLoopbackRedirectUrl(new URL("http://user:pass@localhost:4242/callback"))).toBe(false);
  });
});
