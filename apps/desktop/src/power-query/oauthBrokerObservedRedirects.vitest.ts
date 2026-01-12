import { describe, expect, it } from "vitest";

import { DesktopOAuthBroker } from "./oauthBroker.js";

describe("DesktopOAuthBroker.observeRedirect", () => {
  it("buffers early redirects and resolves the next waitForRedirect call", async () => {
    const broker = new DesktopOAuthBroker();
    const redirectUrl = "formula://oauth/callback?code=abc";

    broker.setOpenAuthUrlHandler(async () => {});
    await broker.openAuthUrl("https://example.com/oauth/authorize?redirect_uri=formula%3A%2F%2Foauth%2Fcallback");

    // Redirect arrives before the PKCE flow registers the wait.
    expect(broker.observeRedirect(redirectUrl)).toBe(false);

    await expect(broker.waitForRedirect("formula://oauth/callback")).resolves.toBe(redirectUrl);
  });

  it("does not resolve waits for a different redirectUri", async () => {
    const broker = new DesktopOAuthBroker();
    const redirectUrl = "formula://oauth/callback?code=abc";

    broker.setOpenAuthUrlHandler(async () => {});
    await broker.openAuthUrl("https://example.com/oauth/authorize?redirect_uri=formula%3A%2F%2Foauth%2Fcallback");

    broker.observeRedirect(redirectUrl);

    const wait = broker.waitForRedirect("formula://other/callback");
    broker.resolveRedirect("formula://other/callback", "formula://other/callback?code=def");

    await expect(wait).resolves.toBe("formula://other/callback?code=def");
  });

  it("does not buffer early redirects that return a different state", async () => {
    const broker = new DesktopOAuthBroker();
    broker.setOpenAuthUrlHandler(async () => {});
    await broker.openAuthUrl(
      "https://example.com/oauth/authorize?redirect_uri=formula%3A%2F%2Foauth%2Fcallback&state=expected",
    );

    // Mismatched state should not be buffered.
    expect(broker.observeRedirect("formula://oauth/callback?code=abc&state=wrong")).toBe(false);

    const wait = broker.waitForRedirect("formula://oauth/callback");
    broker.resolveRedirect("formula://oauth/callback", "formula://oauth/callback?code=def&state=expected");
    await expect(wait).resolves.toBe("formula://oauth/callback?code=def&state=expected");
  });

  it("ignores pending redirects that return a different state", async () => {
    const broker = new DesktopOAuthBroker();
    broker.setOpenAuthUrlHandler(async () => {});
    await broker.openAuthUrl(
      "https://example.com/oauth/authorize?redirect_uri=formula%3A%2F%2Foauth%2Fcallback&state=expected",
    );

    const wait = broker.waitForRedirect("formula://oauth/callback");

    // Wrong state should be ignored (not resolved).
    expect(broker.observeRedirect("formula://oauth/callback?code=abc&state=wrong")).toBe(false);

    // Correct state should resolve.
    expect(broker.observeRedirect("formula://oauth/callback?code=def&state=expected")).toBe(true);

    await expect(wait).resolves.toBe("formula://oauth/callback?code=def&state=expected");
  });
});
