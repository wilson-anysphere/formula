import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("desktop OAuth redirect capture", () => {
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
});
