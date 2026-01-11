/**
 * Low-level OAuth2 token endpoint client.
 *
 * This module implements the protocol mechanics (token exchange, refresh, device
 * code). Higher-level caching/persistence lives in `manager.js`.
 *
 * The client is fetch-based and therefore can run in Node, browsers, and
 * workers as long as a compatible `fetch` implementation is provided.
 */

/**
 * @typedef {Object} OAuth2TokenResponse
 * @property {string} access_token
 * @property {string} token_type
 * @property {number | undefined} [expires_in]
 * @property {string | undefined} [refresh_token]
 * @property {string | undefined} [scope]
 * @property {string | undefined} [id_token]
 */

/**
 * @typedef {Object} OAuth2DeviceCodeResponse
 * @property {string} device_code
 * @property {string} user_code
 * @property {string} verification_uri
 * @property {string | undefined} [verification_uri_complete]
 * @property {number} expires_in
 * @property {number | undefined} [interval]
 */

/**
 * @typedef {Object} OAuth2ErrorResponse
 * @property {string} error
 * @property {string | undefined} [error_description]
 * @property {string | undefined} [error_uri]
 */

export class OAuth2TokenError extends Error {
  /**
   * @param {string} message
   * @param {{ status?: number, error?: string, errorDescription?: string, payload?: unknown } | undefined} [options]
   */
  constructor(message, options = {}) {
    super(message);
    this.name = "OAuth2TokenError";
    this.status = options.status;
    this.error = options.error;
    this.errorDescription = options.errorDescription;
    this.payload = options.payload;
  }
}

/**
 * @param {Response} response
 * @returns {Promise<unknown>}
 */
async function readJsonOrText(response) {
  const contentType = response.headers.get("content-type") ?? "";
  if (contentType.includes("application/json") || contentType.includes("+json")) {
    return await response.json().catch(() => null);
  }
  return await response.text().catch(() => "");
}

/**
 * @param {typeof fetch | null} fetchFn
 * @returns {typeof fetch}
 */
function assertFetch(fetchFn) {
  const fn = fetchFn ?? (typeof fetch === "function" ? fetch.bind(globalThis) : null);
  if (!fn) throw new Error("OAuth2 token client requires a fetch implementation");
  return fn;
}

/**
 * @param {typeof fetch} fetchFn
 * @param {string} url
 * @param {Record<string, string>} params
 * @param {{ signal?: AbortSignal } | undefined} [options]
 * @returns {Promise<OAuth2TokenResponse>}
 */
async function postForm(fetchFn, url, params, options = {}) {
  const body = new URLSearchParams();
  for (const [k, v] of Object.entries(params)) {
    if (v == null) continue;
    body.set(k, v);
  }

  const response = await fetchFn(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
      Accept: "application/json",
    },
    body,
    signal: options.signal,
  });

  const payload = await readJsonOrText(response);
  if (!response.ok) {
    const maybeError = /** @type {OAuth2ErrorResponse | null} */ (
      payload && typeof payload === "object" && "error" in payload ? payload : null
    );
    throw new OAuth2TokenError(`OAuth2 token request failed (${response.status})`, {
      status: response.status,
      error: maybeError?.error,
      errorDescription: maybeError?.error_description,
      payload,
    });
  }

  if (!payload || typeof payload !== "object") {
    throw new OAuth2TokenError("OAuth2 token endpoint returned an invalid response", { status: response.status, payload });
  }

  // @ts-ignore - runtime structure
  return payload;
}

/**
 * @param {number} ms
 * @param {AbortSignal | undefined} signal
 */
function sleep(ms, signal) {
  if (ms <= 0) return Promise.resolve();
  return new Promise((resolve, reject) => {
    if (signal?.aborted) {
      const err = new Error("Aborted");
      err.name = "AbortError";
      reject(err);
      return;
    }
    const timer = setTimeout(() => {
      signal?.removeEventListener("abort", onAbort);
      resolve(undefined);
    }, ms);
    const onAbort = () => {
      clearTimeout(timer);
      const err = new Error("Aborted");
      err.name = "AbortError";
      reject(err);
    };
    signal?.addEventListener("abort", onAbort, { once: true });
  });
}

export class OAuth2TokenClient {
  /**
   * @param {{ fetch?: typeof fetch | undefined, now?: (() => number) | undefined } | undefined} [options]
   */
  constructor(options = {}) {
    this.fetchFn = assertFetch(options.fetch ?? null);
    this.now = options.now ?? (() => Date.now());
  }

  /**
   * Authorization code exchange (PKCE-friendly).
   *
   * @param {{
   *   tokenEndpoint: string;
   *   clientId: string;
   *   clientSecret?: string;
   *   code: string;
   *   redirectUri: string;
   *   codeVerifier?: string;
   *   signal?: AbortSignal;
   * }} params
   * @returns {Promise<OAuth2TokenResponse>}
   */
  async exchangeCode(params) {
    const body = {
      grant_type: "authorization_code",
      code: params.code,
      redirect_uri: params.redirectUri,
      client_id: params.clientId,
      client_secret: params.clientSecret ?? "",
      code_verifier: params.codeVerifier ?? "",
    };
    // Remove empty optional params so we don't send `client_secret=` etc.
    if (!params.clientSecret) delete body.client_secret;
    if (!params.codeVerifier) delete body.code_verifier;
    return await postForm(this.fetchFn, params.tokenEndpoint, body, { signal: params.signal });
  }

  /**
   * Refresh token exchange.
   *
   * @param {{
   *   tokenEndpoint: string;
   *   clientId: string;
   *   clientSecret?: string;
   *   refreshToken: string;
   *   scopes?: string[];
   *   signal?: AbortSignal;
   * }} params
   * @returns {Promise<OAuth2TokenResponse>}
   */
  async refreshToken(params) {
    const body = {
      grant_type: "refresh_token",
      refresh_token: params.refreshToken,
      client_id: params.clientId,
      client_secret: params.clientSecret ?? "",
      scope: params.scopes && params.scopes.length > 0 ? params.scopes.join(" ") : "",
    };
    if (!params.clientSecret) delete body.client_secret;
    if (!params.scopes || params.scopes.length === 0) delete body.scope;
    return await postForm(this.fetchFn, params.tokenEndpoint, body, { signal: params.signal });
  }

  /**
   * RFC 8628 Device Authorization Request.
   *
   * @param {{
   *   deviceAuthorizationEndpoint: string;
   *   clientId: string;
   *   scopes?: string[];
   *   signal?: AbortSignal;
   * }} params
   * @returns {Promise<OAuth2DeviceCodeResponse>}
   */
  async deviceCodeStart(params) {
    const body = {
      client_id: params.clientId,
      scope: params.scopes && params.scopes.length > 0 ? params.scopes.join(" ") : "",
    };
    if (!params.scopes || params.scopes.length === 0) delete body.scope;
    // @ts-ignore - device code response is a subset of token response keys
    return await postForm(this.fetchFn, params.deviceAuthorizationEndpoint, body, { signal: params.signal });
  }

  /**
   * RFC 8628 token polling loop.
   *
   * The method resolves with an access token response or rejects with an
   * `OAuth2TokenError` once the device code expires or is denied.
   *
   * @param {{
   *   tokenEndpoint: string;
   *   clientId: string;
   *   clientSecret?: string;
   *   deviceCode: string;
   *   intervalMs?: number;
   *   expiresAtMs: number;
   *   signal?: AbortSignal;
   * }} params
   * @returns {Promise<OAuth2TokenResponse>}
   */
  async deviceCodePoll(params) {
    let intervalMs = Math.max(1000, params.intervalMs ?? 5000);
    while (this.now() < params.expiresAtMs) {
      try {
        const body = {
          grant_type: "urn:ietf:params:oauth:grant-type:device_code",
          device_code: params.deviceCode,
          client_id: params.clientId,
          client_secret: params.clientSecret ?? "",
        };
        if (!params.clientSecret) delete body.client_secret;
        return await postForm(this.fetchFn, params.tokenEndpoint, body, { signal: params.signal });
      } catch (err) {
        if (!(err instanceof OAuth2TokenError)) throw err;
        // Device code polling uses 400 errors with OAuth2 error codes for control flow.
        if (err.error === "authorization_pending") {
          await sleep(intervalMs, params.signal);
          continue;
        }
        if (err.error === "slow_down") {
          intervalMs += 5000;
          await sleep(intervalMs, params.signal);
          continue;
        }
        if (err.error === "expired_token") {
          throw new OAuth2TokenError("Device code expired", { status: err.status, error: err.error, payload: err.payload });
        }
        throw err;
      }
    }

    throw new OAuth2TokenError("Device code expired", { error: "expired_token" });
  }
}

