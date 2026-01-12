import { shellOpen } from "../tauri/shellOpen.ts";

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
  private observedRedirects: Array<{ url: string; observedAtMs: number }> = [];
  private lastAuthUrlOpenedAtMs: number | null = null;

  // Small buffer to avoid dropping redirects that arrive before `waitForRedirect(...)`
  // is registered (e.g. fast redirects, or deep-link events emitted at app startup).
  private static readonly OBSERVED_REDIRECT_LIMIT = 10;
  // Keep this short: it only exists to bridge a tiny race between `openAuthUrl(...)`
  // and `waitForRedirect(...)`. Longer TTLs risk consuming stale redirects in a
  // future PKCE attempt and causing a confusing "unexpected state value" error.
  private static readonly OBSERVED_REDIRECT_TTL_MS = 30_000;

  setOpenAuthUrlHandler(handler: ((url: string) => Promise<void> | void) | null) {
    this.openAuthUrlHandler = handler;
  }

  setDeviceCodePromptHandler(handler: ((code: string, verificationUri: string) => Promise<void> | void) | null) {
    this.deviceCodePromptHandler = handler;
  }

  async openAuthUrl(url: string) {
    // Security: auth URLs should always be opened via a trusted web protocol.
    // This prevents malicious callers from using the OAuth broker as a generic
    // "open arbitrary protocol" primitive.
    let parsed: URL;
    try {
      parsed = new URL(url);
    } catch {
      throw new Error(`Invalid OAuth auth URL: ${url}`);
    }

    const protocol = parsed.protocol.replace(":", "").toLowerCase();
    if (protocol !== "http" && protocol !== "https") {
      throw new Error(`Refusing to open OAuth auth URL with untrusted protocol "${protocol}:"`);
    }

    // Used to gate buffering of early redirects. We only expect redirects very
    // shortly after opening an auth URL (PKCE flow), so avoid holding onto deep
    // links indefinitely if they're delivered at unrelated times.
    this.lastAuthUrlOpenedAtMs = parsed.searchParams.has("redirect_uri") ? Date.now() : null;

    if (!this.openAuthUrlHandler) {
      await shellOpen(url);
      return;
    }
    await this.openAuthUrlHandler(url);
  }

  waitForRedirect(redirectUri: string): Promise<string> {
    // Once a redirect wait is registered, there's no longer a race window between
    // openAuthUrl(...) and waitForRedirect(...), so we can disable early-redirect
    // buffering for this flow.
    this.lastAuthUrlOpenedAtMs = null;

    // If we observed a redirect before the caller registered the wait, resolve immediately.
    const observed = this.shiftObservedRedirect(redirectUri);
    if (observed) return Promise.resolve(observed);

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
   * Observe a redirect URL coming from an external source (e.g. a Tauri deep-link
   * handler) and either resolve an in-flight redirect or buffer it briefly until
   * a future `waitForRedirect(...)` call registers interest.
   *
   * Returns true when the redirect was accepted and applied to an in-flight wait.
   */
  observeRedirect(redirectUrl: string): boolean {
    if (typeof redirectUrl !== "string" || redirectUrl.trim() === "") return false;

    // Try to resolve an existing pending redirect first.
    const expectedRedirectUri = this.findPendingRedirectUri(redirectUrl);
    if (expectedRedirectUri) {
      this.resolveRedirect(expectedRedirectUri, redirectUrl);
      return true;
    }

    // Otherwise, store it briefly so we don't drop redirects that race ahead of
    // `waitForRedirect`.
    //
    // Only buffer if we've opened an auth URL very recently; this prevents an
    // unrelated deep link (or a redirect from an old flow) from being consumed by
    // a future PKCE attempt.
    const lastOpen = this.lastAuthUrlOpenedAtMs;
    if (typeof lastOpen !== "number" || Date.now() - lastOpen > DesktopOAuthBroker.OBSERVED_REDIRECT_TTL_MS) {
      return false;
    }

    this.pruneObservedRedirects();
    this.observedRedirects.push({ url: redirectUrl, observedAtMs: Date.now() });
    if (this.observedRedirects.length > DesktopOAuthBroker.OBSERVED_REDIRECT_LIMIT) {
      this.observedRedirects.splice(0, this.observedRedirects.length - DesktopOAuthBroker.OBSERVED_REDIRECT_LIMIT);
    }
    return false;
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

  private pruneObservedRedirects(): void {
    const cutoff = Date.now() - DesktopOAuthBroker.OBSERVED_REDIRECT_TTL_MS;
    this.observedRedirects = this.observedRedirects.filter((e) => e.observedAtMs >= cutoff);
  }

  private shiftObservedRedirect(redirectUri: string): string | null {
    this.pruneObservedRedirects();
    const idx = this.observedRedirects.findIndex((e) => matchesRedirectUri(redirectUri, e.url));
    if (idx === -1) return null;
    const [entry] = this.observedRedirects.splice(idx, 1);
    return entry?.url ?? null;
  }

  rejectRedirect(redirectUri: string, reason?: unknown) {
    const pending = this.pendingRedirects.get(redirectUri);
    if (!pending) return;
    this.pendingRedirects.delete(redirectUri);
    pending.reject(reason);
  }
}

export const oauthBroker = new DesktopOAuthBroker();
