import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("data queries: desktop OAuth redirect capture", () => {
  test("emitting oauth-redirect resolves a pending PKCE flow without prompting", async ({ page }) => {
    await page.addInitScript(() => {
      const listeners: Record<string, any> = {};
      (window as any).__tauriListeners = listeners;

      const credentialEntries = new Map<string, { id: string; secret: any }>();
      let credentialIdCounter = 0;

      (window as any).__oauthOpenedUrls = [] as string[];
      (window as any).__promptCalls = [] as Array<{ message?: string; defaultValue?: string | null }>;

      // Track prompt usage so tests can assert we didn't fall back to manual copy/paste.
      window.prompt = ((message?: string, defaultValue?: string) => {
        (window as any).__promptCalls.push({ message, defaultValue: defaultValue ?? null });
        return null;
      }) as any;

      const originalFetch = window.fetch.bind(window);
      window.fetch = async (input: any, init?: any) => {
        const url = typeof input === "string" ? input : input?.url;
        if (url === "https://example.com/oauth/token") {
          return new Response(
            JSON.stringify({
              access_token: "at-1",
              refresh_token: "rt-1",
              token_type: "bearer",
              expires_in: 3600,
            }),
            { status: 200, headers: { "content-type": "application/json" } },
          );
        }
        return originalFetch(input, init);
      };

      // Seed a query + OAuth provider config so the Data Queries panel has an OAuth2 row.
      window.localStorage.setItem(
        "formula.desktop.powerQuery.oauthProviders:local-workbook",
        JSON.stringify([
          {
            id: "example",
            clientId: "client-id",
            tokenEndpoint: "https://example.com/oauth/token",
            authorizationEndpoint: "https://example.com/oauth/authorize",
            redirectUri: "formula://oauth/callback",
            defaultScopes: ["scope-a"],
          },
        ]),
      );

      window.localStorage.setItem(
        "formula.desktop.powerQuery.queries:local-workbook",
        JSON.stringify([
          {
            id: "q_oauth",
            name: "OAuth query",
            source: {
              type: "api",
              url: "https://example.com/api",
              method: "GET",
              auth: { type: "oauth2", providerId: "example", scopes: ["scope-a"] },
            },
            steps: [],
            refreshPolicy: { type: "manual" },
          },
        ]),
      );

      (window as any).__TAURI__ = {
        core: {
          invoke: async (cmd: string, args: any) => {
            if (cmd === "open_external_url") {
              const url = String(args?.url ?? "");
              if (url) (window as any).__oauthOpenedUrls.push(url);
              return null;
            }
            if (cmd === "power_query_state_get") return null;
            if (cmd === "power_query_state_set") return null;

            if (cmd === "power_query_credential_get") {
              const scopeKey = String(args?.scope_key ?? "");
              return scopeKey ? credentialEntries.get(scopeKey) ?? null : null;
            }
            if (cmd === "power_query_credential_set") {
              const scopeKey = String(args?.scope_key ?? "");
              if (!scopeKey) throw new Error("missing scope_key");
              const entry = { id: `id-${++credentialIdCounter}`, secret: args?.secret };
              credentialEntries.set(scopeKey, entry);
              return entry;
            }
            if (cmd === "power_query_credential_delete") {
              const scopeKey = String(args?.scope_key ?? "");
              if (scopeKey) credentialEntries.delete(scopeKey);
              return null;
            }
            if (cmd === "power_query_credential_list") {
              return Array.from(credentialEntries.entries()).map(([scopeKey, entry]) => ({ scopeKey, id: entry.id }));
            }

            // Keep the invoke surface flexible; most desktop shell commands aren't
            // needed for this test.
            return null;
          },
        },
        event: {
          listen: async (name: string, handler: any) => {
            listeners[name] = handler;
            return () => {
              delete listeners[name];
            };
          },
          emit: async () => {},
        },
        shell: {
          open: async (url: string) => {
            (window as any).__oauthOpenedUrls.push(url);
          },
        },
        window: {
          getCurrentWebviewWindow: () => ({
            hide: async () => {},
            close: async () => {},
          }),
        },
      };
    });

    await gotoDesktop(page);

    // Open the Data Queries panel.
    await page.evaluate(() => {
      window.dispatchEvent(new CustomEvent("formula:open-panel", { detail: { panelId: "dataQueries" } }));
    });

    // Kick off the PKCE sign-in.
    const signIn = page.getByRole("button", { name: "Sign in" }).first();
    await expect(signIn).toBeVisible();
    await signIn.click();

    // The desktop deep-link redirect capture path should show an awaiting state and should
    // not offer the old manual paste prompt.
    await expect(page.getByText("Awaiting OAuth redirect…")).toBeVisible();
    await expect(page.getByRole("button", { name: /Paste redirect URL/i })).toHaveCount(0);

    // Wait until the auth URL is opened so we can reuse the generated `state` parameter.
    const authUrl = await page.waitForFunction(() => (window as any).__oauthOpenedUrls?.[0] ?? null);
    const authUrlText = await authUrl.jsonValue();
    const parsed = new URL(String(authUrlText));
    const state = parsed.searchParams.get("state");
    const redirectUri = parsed.searchParams.get("redirect_uri");
    if (!state || !redirectUri) {
      throw new Error(`Expected authorization URL to include state + redirect_uri: ${authUrlText}`);
    }

    // Deliver the redirect back to the running app via the stubbed Tauri event channel.
    await page.waitForFunction(() => Boolean((window as any).__tauriListeners?.["oauth-redirect"]));
    const redirectUrl = `${redirectUri}?code=fake_code&state=${encodeURIComponent(state)}`;
    await page.evaluate((url) => {
      (window as any).__tauriListeners["oauth-redirect"]({ payload: url });
    }, redirectUrl);

    // Flow should complete and swap the button to "Sign out".
    await expect(page.getByRole("button", { name: "Sign out" }).first()).toBeVisible();

    const promptCalls = await page.evaluate(() => (window as any).__promptCalls?.length ?? 0);
    expect(promptCalls).toBe(0);
  });

  for (const redirectUri of ["http://localhost:4242/oauth/callback", "http://[::1]:4242/oauth/callback"] as const) {
    test(`loopback redirectUri (${redirectUri}) invokes oauth_loopback_listen and resolves without prompting`, async ({ page }) => {
      await page.addInitScript((redirectUri) => {
      const listeners: Record<string, any> = {};
      (window as any).__tauriListeners = listeners;

      const invokeCalls: Array<{ cmd: string; args: any }> = [];
      (window as any).__tauriInvokeCalls = invokeCalls;

      const credentialEntries = new Map<string, { id: string; secret: any }>();
      let credentialIdCounter = 0;

      (window as any).__oauthOpenedUrls = [] as string[];
      (window as any).__promptCalls = [] as Array<{ message?: string; defaultValue?: string | null }>;

      // Track prompt usage so tests can assert we didn't fall back to manual copy/paste.
      window.prompt = ((message?: string, defaultValue?: string) => {
        (window as any).__promptCalls.push({ message, defaultValue: defaultValue ?? null });
        return null;
      }) as any;

      const originalFetch = window.fetch.bind(window);
      window.fetch = async (input: any, init?: any) => {
        const url = typeof input === "string" ? input : input?.url;
        if (url === "https://example.com/oauth/token") {
          return new Response(
            JSON.stringify({
              access_token: "at-1",
              refresh_token: "rt-1",
              token_type: "bearer",
              expires_in: 3600,
            }),
            { status: 200, headers: { "content-type": "application/json" } },
          );
        }
        return originalFetch(input, init);
      };

      // Seed a query + OAuth provider config so the Data Queries panel has an OAuth2 row.
      // Use an RFC 8252 loopback redirect URI.
      window.localStorage.setItem(
        "formula.desktop.powerQuery.oauthProviders:local-workbook",
        JSON.stringify([
          {
            id: "example",
            clientId: "client-id",
            tokenEndpoint: "https://example.com/oauth/token",
            authorizationEndpoint: "https://example.com/oauth/authorize",
            redirectUri,
            defaultScopes: ["scope-a"],
          },
        ]),
      );

      window.localStorage.setItem(
        "formula.desktop.powerQuery.queries:local-workbook",
        JSON.stringify([
          {
            id: "q_oauth",
            name: "OAuth query",
            source: {
              type: "api",
              url: "https://example.com/api",
              method: "GET",
              auth: { type: "oauth2", providerId: "example", scopes: ["scope-a"] },
            },
            steps: [],
            refreshPolicy: { type: "manual" },
          },
        ]),
      );

      (window as any).__TAURI__ = {
        core: {
          invoke: async (cmd: string, args: any) => {
            invokeCalls.push({ cmd, args });

            if (cmd === "open_external_url") {
              const url = String(args?.url ?? "");
              if (url) (window as any).__oauthOpenedUrls.push(url);
              return null;
            }
            if (cmd === "power_query_state_get") return null;
            if (cmd === "power_query_state_set") return null;

            if (cmd === "oauth_loopback_listen") return null;

            if (cmd === "power_query_credential_get") {
              const scopeKey = String(args?.scope_key ?? "");
              return scopeKey ? credentialEntries.get(scopeKey) ?? null : null;
            }
            if (cmd === "power_query_credential_set") {
              const scopeKey = String(args?.scope_key ?? "");
              if (!scopeKey) throw new Error("missing scope_key");
              const entry = { id: `id-${++credentialIdCounter}`, secret: args?.secret };
              credentialEntries.set(scopeKey, entry);
              return entry;
            }
            if (cmd === "power_query_credential_delete") {
              const scopeKey = String(args?.scope_key ?? "");
              if (scopeKey) credentialEntries.delete(scopeKey);
              return null;
            }
            if (cmd === "power_query_credential_list") {
              return Array.from(credentialEntries.entries()).map(([scopeKey, entry]) => ({ scopeKey, id: entry.id }));
            }

            // Keep the invoke surface flexible; most desktop shell commands aren't
            // needed for this test.
            return null;
          },
        },
        event: {
          listen: async (name: string, handler: any) => {
            listeners[name] = handler;
            return () => {
              delete listeners[name];
            };
          },
          emit: async () => {},
        },
        shell: {
          open: async (url: string) => {
            (window as any).__oauthOpenedUrls.push(url);
          },
        },
        window: {
          getCurrentWebviewWindow: () => ({
            hide: async () => {},
            close: async () => {},
          }),
        },
      };
      }, redirectUri);

      await gotoDesktop(page);

      // Open the Data Queries panel.
      await page.evaluate(() => {
        window.dispatchEvent(new CustomEvent("formula:open-panel", { detail: { panelId: "dataQueries" } }));
      });

      const signIn = page.getByRole("button", { name: "Sign in" }).first();
      await expect(signIn).toBeVisible();
      await signIn.click();

      // Loopback redirect capture should show the awaiting state with no paste prompt.
      await expect(page.getByText("Awaiting OAuth redirect…")).toBeVisible();
      await expect(page.getByRole("button", { name: /Paste redirect URL/i })).toHaveCount(0);

      // Ensure the host was asked to start the loopback listener.
      await page.waitForFunction(
        () => (window as any).__tauriInvokeCalls?.some((c: any) => c?.cmd === "oauth_loopback_listen"),
        undefined,
        { timeout: 10_000 },
      );

      const loopbackArgs = await page.evaluate(() => {
        const calls = (window as any).__tauriInvokeCalls as Array<{ cmd: string; args: any }> | undefined;
        const call = calls?.find((c) => c.cmd === "oauth_loopback_listen");
        return call?.args ?? null;
      });
      expect(loopbackArgs).toEqual({ redirect_uri: redirectUri });

      // Wait until the auth URL is opened so we can reuse the generated `state` parameter.
      const authUrl = await page.waitForFunction(() => (window as any).__oauthOpenedUrls?.[0] ?? null);
      const authUrlText = await authUrl.jsonValue();
      const parsed = new URL(String(authUrlText));
      const state = parsed.searchParams.get("state");
      const openedRedirectUri = parsed.searchParams.get("redirect_uri");
      if (!state || !openedRedirectUri) {
        throw new Error(`Expected authorization URL to include state + redirect_uri: ${authUrlText}`);
      }

      // Ensure loopback listener startup happens before opening the system browser.
      const invokeOrdering = await page.evaluate((authUrlText) => {
        const calls = (window as any).__tauriInvokeCalls as Array<{ cmd: string; args: any }> | undefined;
        if (!Array.isArray(calls)) return null;
        const listenIdx = calls.findIndex((c) => c?.cmd === "oauth_loopback_listen");
        const openIdx = calls.findIndex((c) => c?.cmd === "open_external_url" && String(c?.args?.url ?? "") === String(authUrlText));
        return { listenIdx, openIdx };
      }, authUrlText);
      expect(invokeOrdering).not.toBeNull();
      expect(invokeOrdering!.listenIdx).toBeGreaterThanOrEqual(0);
      expect(invokeOrdering!.openIdx).toBeGreaterThan(invokeOrdering!.listenIdx);

      // Deliver the redirect back to the running app via the stubbed Tauri event channel.
      await page.waitForFunction(() => Boolean((window as any).__tauriListeners?.["oauth-redirect"]));
      const redirectUrl = `${openedRedirectUri}?code=fake_code&state=${encodeURIComponent(state)}`;
      await page.evaluate((url) => {
        (window as any).__tauriListeners["oauth-redirect"]({ payload: url });
      }, redirectUrl);

      await expect(page.getByRole("button", { name: "Sign out" }).first()).toBeVisible();

      const promptCalls = await page.evaluate(() => (window as any).__promptCalls?.length ?? 0);
      expect(promptCalls).toBe(0);
    });
  }

  test("ignores oauth-redirect events that don't match the pending redirectUri", async ({ page }) => {
    await page.addInitScript(() => {
      const listeners: Record<string, any> = {};
      (window as any).__tauriListeners = listeners;

      const credentialEntries = new Map<string, { id: string; secret: any }>();
      let credentialIdCounter = 0;

      (window as any).__oauthOpenedUrls = [] as string[];
      (window as any).__promptCalls = [] as Array<{ message?: string; defaultValue?: string | null }>;

      // Track prompt usage so tests can assert we didn't fall back to manual copy/paste.
      window.prompt = ((message?: string, defaultValue?: string) => {
        (window as any).__promptCalls.push({ message, defaultValue: defaultValue ?? null });
        return null;
      }) as any;

      const originalFetch = window.fetch.bind(window);
      window.fetch = async (input: any, init?: any) => {
        const url = typeof input === "string" ? input : input?.url;
        if (url === "https://example.com/oauth/token") {
          return new Response(
            JSON.stringify({
              access_token: "at-1",
              refresh_token: "rt-1",
              token_type: "bearer",
              expires_in: 3600,
            }),
            { status: 200, headers: { "content-type": "application/json" } },
          );
        }
        return originalFetch(input, init);
      };

      // Seed a query + OAuth provider config so the Data Queries panel has an OAuth2 row.
      window.localStorage.setItem(
        "formula.desktop.powerQuery.oauthProviders:local-workbook",
        JSON.stringify([
          {
            id: "example",
            clientId: "client-id",
            tokenEndpoint: "https://example.com/oauth/token",
            authorizationEndpoint: "https://example.com/oauth/authorize",
            redirectUri: "formula://oauth/callback",
            defaultScopes: ["scope-a"],
          },
        ]),
      );

      window.localStorage.setItem(
        "formula.desktop.powerQuery.queries:local-workbook",
        JSON.stringify([
          {
            id: "q_oauth",
            name: "OAuth query",
            source: {
              type: "api",
              url: "https://example.com/api",
              method: "GET",
              auth: { type: "oauth2", providerId: "example", scopes: ["scope-a"] },
            },
            steps: [],
            refreshPolicy: { type: "manual" },
          },
        ]),
      );

      (window as any).__TAURI__ = {
        core: {
          invoke: async (cmd: string, args: any) => {
            if (cmd === "open_external_url") {
              const url = String(args?.url ?? "");
              if (url) (window as any).__oauthOpenedUrls.push(url);
              return null;
            }
            if (cmd === "power_query_state_get") return null;
            if (cmd === "power_query_state_set") return null;

            if (cmd === "power_query_credential_get") {
              const scopeKey = String(args?.scope_key ?? "");
              return scopeKey ? credentialEntries.get(scopeKey) ?? null : null;
            }
            if (cmd === "power_query_credential_set") {
              const scopeKey = String(args?.scope_key ?? "");
              if (!scopeKey) throw new Error("missing scope_key");
              const entry = { id: `id-${++credentialIdCounter}`, secret: args?.secret };
              credentialEntries.set(scopeKey, entry);
              return entry;
            }
            if (cmd === "power_query_credential_delete") {
              const scopeKey = String(args?.scope_key ?? "");
              if (scopeKey) credentialEntries.delete(scopeKey);
              return null;
            }
            if (cmd === "power_query_credential_list") {
              return Array.from(credentialEntries.entries()).map(([scopeKey, entry]) => ({ scopeKey, id: entry.id }));
            }
            return null;
          },
        },
        event: {
          listen: async (name: string, handler: any) => {
            listeners[name] = handler;
            return () => {
              delete listeners[name];
            };
          },
          emit: async () => {},
        },
        shell: {
          open: async (url: string) => {
            (window as any).__oauthOpenedUrls.push(url);
          },
        },
        window: {
          getCurrentWebviewWindow: () => ({
            hide: async () => {},
            close: async () => {},
          }),
        },
      };
    });

    await gotoDesktop(page);

    // Open the Data Queries panel.
    await page.evaluate(() => {
      window.dispatchEvent(new CustomEvent("formula:open-panel", { detail: { panelId: "dataQueries" } }));
    });

    // Kick off the PKCE sign-in.
    const signIn = page.getByRole("button", { name: "Sign in" }).first();
    await expect(signIn).toBeVisible();
    await signIn.click();

    await expect(page.getByText("Awaiting OAuth redirect…")).toBeVisible();

    // Wait until the auth URL is opened so we can reuse the generated `state` parameter.
    const authUrl = await page.waitForFunction(() => (window as any).__oauthOpenedUrls?.[0] ?? null);
    const authUrlText = await authUrl.jsonValue();
    const parsed = new URL(String(authUrlText));
    const state = parsed.searchParams.get("state");
    const redirectUri = parsed.searchParams.get("redirect_uri");
    if (!state || !redirectUri) {
      throw new Error(`Expected authorization URL to include state + redirect_uri: ${authUrlText}`);
    }

    // Deliver a redirect that *does not* match the pending redirect URI. The flow should
    // remain pending (no sign-out button yet).
    await page.waitForFunction(() => Boolean((window as any).__tauriListeners?.["oauth-redirect"]));
    await page.evaluate(() => {
      (window as any).__tauriListeners["oauth-redirect"]({ payload: "formula://evil/callback?code=x&state=y" });
    });

    await expect(page.getByRole("button", { name: "Sign out" })).toHaveCount(0);
    await expect(page.getByText("Awaiting OAuth redirect…")).toBeVisible();

    // Now deliver the correct redirect URL.
    const correctRedirectUrl = `${redirectUri}?code=fake_code&state=${encodeURIComponent(state)}`;
    await page.evaluate((url) => {
      (window as any).__tauriListeners["oauth-redirect"]({ payload: url });
    }, correctRedirectUrl);

    await expect(page.getByRole("button", { name: "Sign out" }).first()).toBeVisible();

    const promptCalls = await page.evaluate(() => (window as any).__promptCalls?.length ?? 0);
    expect(promptCalls).toBe(0);
  });

  test("can cancel a pending PKCE redirect without showing an error", async ({ page }) => {
    await page.addInitScript(() => {
      const listeners: Record<string, any> = {};
      (window as any).__tauriListeners = listeners;

      const credentialEntries = new Map<string, { id: string; secret: any }>();
      let credentialIdCounter = 0;

      (window as any).__oauthOpenedUrls = [] as string[];
      (window as any).__promptCalls = [] as Array<{ message?: string; defaultValue?: string | null }>;

      // Track prompt usage so tests can assert we didn't fall back to manual copy/paste.
      window.prompt = ((message?: string, defaultValue?: string) => {
        (window as any).__promptCalls.push({ message, defaultValue: defaultValue ?? null });
        return null;
      }) as any;

      // Seed a query + OAuth provider config so the Data Queries panel has an OAuth2 row.
      window.localStorage.setItem(
        "formula.desktop.powerQuery.oauthProviders:local-workbook",
        JSON.stringify([
          {
            id: "example",
            clientId: "client-id",
            tokenEndpoint: "https://example.com/oauth/token",
            authorizationEndpoint: "https://example.com/oauth/authorize",
            redirectUri: "formula://oauth/callback",
            defaultScopes: ["scope-a"],
          },
        ]),
      );

      window.localStorage.setItem(
        "formula.desktop.powerQuery.queries:local-workbook",
        JSON.stringify([
          {
            id: "q_oauth",
            name: "OAuth query",
            source: {
              type: "api",
              url: "https://example.com/api",
              method: "GET",
              auth: { type: "oauth2", providerId: "example", scopes: ["scope-a"] },
            },
            steps: [],
            refreshPolicy: { type: "manual" },
          },
        ]),
      );

      (window as any).__TAURI__ = {
        core: {
          invoke: async (cmd: string, args: any) => {
            if (cmd === "open_external_url") {
              const url = String(args?.url ?? "");
              if (url) (window as any).__oauthOpenedUrls.push(url);
              return null;
            }
            if (cmd === "power_query_state_get") return null;
            if (cmd === "power_query_state_set") return null;

            if (cmd === "power_query_credential_get") {
              const scopeKey = String(args?.scope_key ?? "");
              return scopeKey ? credentialEntries.get(scopeKey) ?? null : null;
            }
            if (cmd === "power_query_credential_set") {
              const scopeKey = String(args?.scope_key ?? "");
              if (!scopeKey) throw new Error("missing scope_key");
              const entry = { id: `id-${++credentialIdCounter}`, secret: args?.secret };
              credentialEntries.set(scopeKey, entry);
              return entry;
            }
            if (cmd === "power_query_credential_delete") {
              const scopeKey = String(args?.scope_key ?? "");
              if (scopeKey) credentialEntries.delete(scopeKey);
              return null;
            }
            if (cmd === "power_query_credential_list") {
              return Array.from(credentialEntries.entries()).map(([scopeKey, entry]) => ({ scopeKey, id: entry.id }));
            }
            return null;
          },
        },
        event: {
          listen: async (name: string, handler: any) => {
            listeners[name] = handler;
            return () => {
              delete listeners[name];
            };
          },
          emit: async () => {},
        },
        shell: {
          open: async (url: string) => {
            (window as any).__oauthOpenedUrls.push(url);
          },
        },
        window: {
          getCurrentWebviewWindow: () => ({
            hide: async () => {},
            close: async () => {},
          }),
        },
      };
    });

    await gotoDesktop(page);

    // Open the Data Queries panel.
    await page.evaluate(() => {
      window.dispatchEvent(new CustomEvent("formula:open-panel", { detail: { panelId: "dataQueries" } }));
    });

    const signIn = page.getByRole("button", { name: "Sign in" }).first();
    await expect(signIn).toBeVisible();
    await signIn.click();

    await expect(page.getByText("Awaiting OAuth redirect…")).toBeVisible();

    const cancelButton = page.getByRole("button", { name: "Cancel sign-in" }).first();
    await expect(cancelButton).toBeVisible();
    await cancelButton.click();

    await expect(page.getByText("Awaiting OAuth redirect…")).toHaveCount(0);
    await expect(page.getByRole("button", { name: "Sign in" }).first()).toBeVisible();

    // No prompt should be shown and cancellation should not surface a global error.
    const promptCalls = await page.evaluate(() => (window as any).__promptCalls?.length ?? 0);
    expect(promptCalls).toBe(0);
    await expect(page.getByText(/cancelled/i)).toHaveCount(0);
  });

  test("redirect emitted during openAuthUrl (before waitForRedirect) still completes PKCE", async ({ page }) => {
    await page.addInitScript(() => {
      const listeners: Record<string, any> = {};
      const emitted: Array<{ event: string; payload: any }> = [];
      const callOrder: Array<{ kind: "listen" | "listen-registered" | "emit"; name: string; seq: number }> = [];
      let seq = 0;

      (window as any).__tauriListeners = listeners;
      (window as any).__tauriEmittedEvents = emitted;
      (window as any).__tauriCallOrder = callOrder;

      const credentialEntries = new Map<string, { id: string; secret: any }>();
      let credentialIdCounter = 0;

      (window as any).__oauthOpenedUrls = [] as string[];
      (window as any).__promptCalls = [] as Array<{ message?: string; defaultValue?: string | null }>;

      // Track prompt usage so tests can assert we didn't fall back to manual copy/paste.
      window.prompt = ((message?: string, defaultValue?: string) => {
        (window as any).__promptCalls.push({ message, defaultValue: defaultValue ?? null });
        return null;
      }) as any;

      const originalFetch = window.fetch.bind(window);
      window.fetch = async (input: any, init?: any) => {
        const url = typeof input === "string" ? input : input?.url;
        if (url === "https://example.com/oauth/token") {
          return new Response(
            JSON.stringify({
              access_token: "at-1",
              refresh_token: "rt-1",
              token_type: "bearer",
              expires_in: 3600,
            }),
            { status: 200, headers: { "content-type": "application/json" } },
          );
        }
        return originalFetch(input, init);
      };

      window.localStorage.setItem(
        "formula.desktop.powerQuery.oauthProviders:local-workbook",
        JSON.stringify([
          {
            id: "example",
            clientId: "client-id",
            tokenEndpoint: "https://example.com/oauth/token",
            authorizationEndpoint: "https://example.com/oauth/authorize",
            redirectUri: "formula://oauth/callback",
            defaultScopes: ["scope-a"],
          },
        ]),
      );

      window.localStorage.setItem(
        "formula.desktop.powerQuery.queries:local-workbook",
        JSON.stringify([
          {
            id: "q_oauth",
            name: "OAuth query",
            source: {
              type: "api",
              url: "https://example.com/api",
              method: "GET",
              auth: { type: "oauth2", providerId: "example", scopes: ["scope-a"] },
            },
            steps: [],
            refreshPolicy: { type: "manual" },
          },
        ]),
      );

      (window as any).__TAURI__ = {
        core: {
          invoke: async (cmd: string, args: any) => {
            if (cmd === "open_external_url") {
              const url = String(args?.url ?? "");
              if (url) (window as any).__oauthOpenedUrls.push(url);
              // Immediately bounce the redirect back via the event channel *before*
              // the PKCE flow registers `waitForRedirect(...)`.
              try {
                const parsed = new URL(url);
                const state = parsed.searchParams.get("state");
                const redirectUri = parsed.searchParams.get("redirect_uri");
                if (state && redirectUri && typeof listeners["oauth-redirect"] === "function") {
                  const redirectUrl = `${redirectUri}?code=fake_code&state=${encodeURIComponent(state)}`;
                  listeners["oauth-redirect"]({ payload: redirectUrl });
                }
              } catch {
                // ignore invalid URLs in tests
              }
              return null;
            }
            if (cmd === "power_query_state_get") return null;
            if (cmd === "power_query_state_set") return null;

            if (cmd === "power_query_credential_get") {
              const scopeKey = String(args?.scope_key ?? "");
              return scopeKey ? credentialEntries.get(scopeKey) ?? null : null;
            }
            if (cmd === "power_query_credential_set") {
              const scopeKey = String(args?.scope_key ?? "");
              if (!scopeKey) throw new Error("missing scope_key");
              const entry = { id: `id-${++credentialIdCounter}`, secret: args?.secret };
              credentialEntries.set(scopeKey, entry);
              return entry;
            }
            if (cmd === "power_query_credential_delete") {
              const scopeKey = String(args?.scope_key ?? "");
              if (scopeKey) credentialEntries.delete(scopeKey);
              return null;
            }
            if (cmd === "power_query_credential_list") {
              return Array.from(credentialEntries.entries()).map(([scopeKey, entry]) => ({ scopeKey, id: entry.id }));
            }

            return null;
          },
        },
        event: {
          listen: async (name: string, handler: any) => {
            // Simulate async backend confirmation for handler registration.
            callOrder.push({ kind: "listen", name, seq: ++seq });
            await Promise.resolve();
            listeners[name] = handler;
            callOrder.push({ kind: "listen-registered", name, seq: ++seq });
            return () => {
              delete listeners[name];
            };
          },
          emit: async (event: string, payload?: any) => {
            callOrder.push({ kind: "emit", name: event, seq: ++seq });
            emitted.push({ event, payload });
          },
        },
        shell: {
          open: async (url: string) => {
            // Immediately bounce the redirect back via the event channel *before*
            // the PKCE flow registers `waitForRedirect(...)`.
            const parsed = new URL(url);
            const state = parsed.searchParams.get("state");
            const redirectUri = parsed.searchParams.get("redirect_uri");
            if (state && redirectUri && typeof listeners["oauth-redirect"] === "function") {
              const redirectUrl = `${redirectUri}?code=fake_code&state=${encodeURIComponent(state)}`;
              listeners["oauth-redirect"]({ payload: redirectUrl });
            }
          },
        },
        window: {
          getCurrentWebviewWindow: () => ({
            hide: async () => {},
            close: async () => {},
          }),
        },
      };
    });

    await gotoDesktop(page);

    // Ensure the frontend handshake fired (used by the Rust side to flush queued redirects).
    await page.waitForFunction(() =>
      Boolean((window as any).__tauriEmittedEvents?.some((entry: any) => entry?.event === "oauth-redirect-ready")),
    );
    const ordering = await page.evaluate(() => {
      const calls = (window as any).__tauriCallOrder as Array<{ kind: string; name: string; seq: number }> | undefined;
      if (!Array.isArray(calls)) return null;
      const redirectRegistered = calls.find((c) => c.kind === "listen-registered" && c.name === "oauth-redirect")?.seq ?? null;
      const readyEmitted = calls.find((c) => c.kind === "emit" && c.name === "oauth-redirect-ready")?.seq ?? null;
      return { redirectRegistered, readyEmitted };
    });
    expect(ordering).not.toBeNull();
    expect(ordering!.redirectRegistered).not.toBeNull();
    expect(ordering!.readyEmitted).not.toBeNull();
    expect(ordering!.readyEmitted!).toBeGreaterThan(ordering!.redirectRegistered!);

    // Open the Data Queries panel.
    await page.evaluate(() => {
      window.dispatchEvent(new CustomEvent("formula:open-panel", { detail: { panelId: "dataQueries" } }));
    });

    const signIn = page.getByRole("button", { name: "Sign in" }).first();
    await expect(signIn).toBeVisible();
    await signIn.click();

    // Opening the auth URL triggers the oauth-redirect event immediately; the flow should still
    // complete (without manual paste/prompt) even though the redirect arrives before
    // `waitForRedirect` is registered.
    await expect(page.getByRole("button", { name: "Sign out" }).first()).toBeVisible();

    const promptCalls = await page.evaluate(() => (window as any).__promptCalls?.length ?? 0);
    expect(promptCalls).toBe(0);
  });
});
