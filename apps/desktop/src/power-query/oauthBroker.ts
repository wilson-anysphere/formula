import { shellOpen } from "../tauri/shellOpen";

export type OAuthBroker = {
  /**
   * Open a URL for the user to authenticate (system browser, in-app webview, etc).
   */
  openAuthUrl(url: string): Promise<void> | void;

  /**
   * Wait for a redirect to the provided redirect URI and resolve with the full
   * redirect URL (including query parameters).
   */
  waitForRedirect?(redirectUri: string): Promise<string>;

  /**
   * Display a device code prompt to the user.
   */
  deviceCodePrompt?(code: string, verificationUri: string): Promise<void> | void;
};

/**
 * Returns true if `redirectUrl` looks like an OAuth redirect to `redirectUri`.
 *
 * Security note: we intentionally match only the parts that identify the redirect
 * endpoint (scheme + host + path) and ignore query/fragment parameters. This lets
 * callers pass the full URL from a deep link / loopback handler while preventing
 * arbitrary URLs from being treated as redirects.
 */
export function matchesRedirectUri(redirectUri: string, redirectUrl: string): boolean {
  if (typeof redirectUri !== "string" || typeof redirectUrl !== "string") return false;
  const expectedText = redirectUri.trim();
  const actualText = redirectUrl.trim();
  if (!expectedText || !actualText) return false;

  let expected: URL;
  let actual: URL;
  try {
    expected = new URL(expectedText);
    actual = new URL(actualText);
  } catch {
    return false;
  }

  // Match the redirect "endpoint" exactly.
  // OAuth providers should treat redirects as exact matches; we follow that
  // expectation here (case-insensitive host comparison is handled by URL).
  return (
    expected.protocol === actual.protocol &&
    expected.hostname === actual.hostname &&
    expected.port === actual.port &&
    expected.pathname === actual.pathname
  );
}

type Deferred<T> = {
  promise: Promise<T>;
  resolve: (value: T) => void;
  reject: (reason?: unknown) => void;
};

function defer<T>(): Deferred<T> {
  let resolve!: (value: T) => void;
  let reject!: (reason?: unknown) => void;
  const promise = new Promise<T>((res, rej) => {
    resolve = res;
    reject = rej;
  });
  return { promise, resolve, reject };
}

/**
 * Minimal in-process broker implementation suitable for early prototypes.
 *
 * UI can wire into this broker by:
 * - setting `setOpenAuthUrlHandler`
 * - calling `resolveRedirect(...)` once the app observes the redirect
 */
export class DesktopOAuthBroker implements OAuthBroker {
  private openAuthUrlHandler: ((url: string) => Promise<void> | void) | null = null;
  private deviceCodePromptHandler: ((code: string, verificationUri: string) => Promise<void> | void) | null = null;
  private pendingRedirects = new Map<string, Deferred<string>>();

  setOpenAuthUrlHandler(handler: ((url: string) => Promise<void> | void) | null) {
    this.openAuthUrlHandler = handler;
  }

  setDeviceCodePromptHandler(handler: ((code: string, verificationUri: string) => Promise<void> | void) | null) {
    this.deviceCodePromptHandler = handler;
  }

  async openAuthUrl(url: string) {
    if (!this.openAuthUrlHandler) {
      await shellOpen(url);
      return;
    }
    await this.openAuthUrlHandler(url);
  }

  waitForRedirect(redirectUri: string): Promise<string> {
    const existing = this.pendingRedirects.get(redirectUri);
    if (existing) return existing.promise;
    const d = defer<string>();
    this.pendingRedirects.set(redirectUri, d);
    return d.promise;
  }

  async deviceCodePrompt(code: string, verificationUri: string) {
    if (!this.deviceCodePromptHandler) return;
    await this.deviceCodePromptHandler(code, verificationUri);
  }

  resolveRedirect(redirectUri: string, redirectUrl: string) {
    // Security: ensure the observed URL matches the pending redirect endpoint before resolving.
    if (!matchesRedirectUri(redirectUri, redirectUrl)) return;
    const pending = this.pendingRedirects.get(redirectUri);
    if (!pending) return;
    this.pendingRedirects.delete(redirectUri);
    pending.resolve(redirectUrl);
  }

  /**
   * Find a pending redirect URI (registered via `waitForRedirect`) that matches
   * the provided full redirect URL. Returns `null` when no pending redirect
   * matches.
   */
  findPendingRedirectUri(redirectUrl: string): string | null {
    for (const redirectUri of this.pendingRedirects.keys()) {
      if (matchesRedirectUri(redirectUri, redirectUrl)) return redirectUri;
    }
    return null;
  }

  rejectRedirect(redirectUri: string, reason?: unknown) {
    const pending = this.pendingRedirects.get(redirectUri);
    if (!pending) return;
    this.pendingRedirects.delete(redirectUri);
    pending.reject(reason);
  }
}

export const oauthBroker = new DesktopOAuthBroker();
