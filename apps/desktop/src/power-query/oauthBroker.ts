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
      throw new Error("No OAuth openAuthUrl handler registered");
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
    const pending = this.pendingRedirects.get(redirectUri);
    if (!pending) return;
    this.pendingRedirects.delete(redirectUri);
    pending.resolve(redirectUrl);
  }

  rejectRedirect(redirectUri: string, reason?: unknown) {
    const pending = this.pendingRedirects.get(redirectUri);
    if (!pending) return;
    this.pendingRedirects.delete(redirectUri);
    pending.reject(reason);
  }
}

export const oauthBroker = new DesktopOAuthBroker();

