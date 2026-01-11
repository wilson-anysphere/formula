import { OAuth2TokenClient } from "./tokenClient.js";
import { InMemoryOAuthTokenStore, normalizeScopes } from "./tokenStore.js";
import { createCodeChallenge, createCodeVerifier } from "./pkce.js";

const DEFAULT_CLOCK_SKEW_MS = 60_000;

/**
 * @typedef {import("./tokenStore.js").OAuth2TokenStore} OAuth2TokenStore
 * @typedef {import("./tokenStore.js").OAuth2TokenStoreEntry} OAuth2TokenStoreEntry
 * @typedef {import("./tokenStore.js").OAuth2TokenStoreKey} OAuth2TokenStoreKey
 */

/**
 * @typedef {Object} OAuth2ProviderConfig
 * @property {string} id Stable provider/config identifier.
 * @property {string} clientId
 * @property {string | undefined} [clientSecret]
 * @property {string} tokenEndpoint
 * @property {string | undefined} [authorizationEndpoint] Authorization endpoint for auth-code flows.
 * @property {string | undefined} [redirectUri] Redirect URI for auth-code flows.
 * @property {string | undefined} [deviceAuthorizationEndpoint] Device authorization endpoint (RFC 8628).
 * @property {string[] | undefined} [defaultScopes]
 * @property {Record<string, string> | undefined} [authorizationParams]
 */

/**
 * @typedef {Object} OAuth2Broker
 * @property {(url: string) => void | Promise<void>} openAuthUrl
 * @property {(redirectUri: string) => Promise<string>} [waitForRedirect]
 * @property {(code: string, verificationUri: string) => void | Promise<void>} [deviceCodePrompt]
 */

/**
 * @typedef {Object} CachedTokenState
 * @property {string | null} accessToken
 * @property {number | null} expiresAtMs
 * @property {string | null} refreshToken
 * @property {string[]} scopes
 */

/**
 * @typedef {{
 *   providerId: string;
 *   scopes?: string[];
 *   signal?: AbortSignal;
 *   now?: (() => number);
 *   forceRefresh?: boolean;
 * }} GetAccessTokenOptions
 */

/**
 * @typedef {{
 *   accessToken: string;
 *   expiresAtMs: number | null;
 *   refreshToken: string | null;
 * }} OAuth2AccessTokenResult
 */

/**
 * High-level OAuth2 manager with:
 * - access token caching (with clock skew)
 * - refresh token rotation
 * - optional persistence via a pluggable token store
 * - concurrency dedupe (one refresh per provider+scopes at a time)
 */
export class OAuth2Manager {
  /**
   * @param {{
   *   tokenStore?: OAuth2TokenStore | undefined;
   *   fetch?: typeof fetch | undefined;
   *   now?: (() => number) | undefined;
   *   clockSkewMs?: number | undefined;
   *   persistAccessToken?: boolean | undefined;
   * } | undefined} [options]
   */
  constructor(options = {}) {
    /** @type {OAuth2TokenStore} */
    this.tokenStore = options.tokenStore ?? new InMemoryOAuthTokenStore();
    this.now = options.now ?? (() => Date.now());
    this.clockSkewMs = options.clockSkewMs ?? DEFAULT_CLOCK_SKEW_MS;
    this.persistAccessToken = options.persistAccessToken ?? false;

    this.client = new OAuth2TokenClient({ fetch: options.fetch, now: this.now });

    /** @type {Map<string, OAuth2ProviderConfig>} */
    this.providers = new Map();

    /** @type {Map<string, CachedTokenState>} */
    this.cache = new Map();

    /** @type {Map<string, Promise<OAuth2AccessTokenResult>>} */
    this.inFlight = new Map();
  }

  /**
   * @param {OAuth2ProviderConfig} config
   */
  registerProvider(config) {
    if (!config || typeof config !== "object") throw new Error("OAuth2 provider config is required");
    if (!config.id) throw new Error("OAuth2 provider config requires an id");
    if (!config.clientId) throw new Error(`OAuth2 provider '${config.id}' requires a clientId`);
    if (!config.tokenEndpoint) throw new Error(`OAuth2 provider '${config.id}' requires a tokenEndpoint`);
    this.providers.set(config.id, config);
  }

  /**
   * @param {string} providerId
   * @returns {OAuth2ProviderConfig}
   */
  getProvider(providerId) {
    const provider = this.providers.get(providerId);
    if (!provider) throw new Error(`Unknown OAuth2 provider '${providerId}'`);
    return provider;
  }

  /**
   * @param {string} providerId
   * @param {string[]} scopes
   * @returns {OAuth2TokenStoreKey}
   */
  makeStoreKey(providerId, scopes) {
    const { scopesHash } = normalizeScopes(scopes);
    return { providerId, scopesHash };
  }

  /**
   * @param {OAuth2TokenStoreKey} key
   * @returns {string}
   */
  static keyString(key) {
    return `${key.providerId}:${key.scopesHash}`;
  }

  /**
   * @param {CachedTokenState | undefined} state
   * @param {number} nowMs
   */
  isAccessTokenValid(state, nowMs) {
    if (!state?.accessToken) return false;
    if (state.expiresAtMs == null) return true;
    return nowMs < state.expiresAtMs - this.clockSkewMs;
  }

  /**
   * Load any persisted token data into the in-memory cache.
   *
   * @param {OAuth2TokenStoreKey} key
   * @param {string[]} scopes
   * @returns {Promise<CachedTokenState>}
   */
  async hydrateCache(key, scopes) {
    const cacheKey = OAuth2Manager.keyString(key);
    const existing = this.cache.get(cacheKey);
    if (existing) return existing;

    const entry = await this.tokenStore.get(key);
    /** @type {CachedTokenState} */
    const state = {
      accessToken: entry?.accessToken ?? null,
      expiresAtMs: entry?.expiresAtMs ?? null,
      refreshToken: entry?.refreshToken ?? null,
      scopes,
    };
    this.cache.set(cacheKey, state);
    return state;
  }

  /**
   * Obtain a valid access token for a provider + scope set.
   *
   * @param {GetAccessTokenOptions} options
   * @returns {Promise<OAuth2AccessTokenResult>}
   */
  async getAccessToken(options) {
    const provider = this.getProvider(options.providerId);
    const normalized = normalizeScopes(options.scopes ?? provider.defaultScopes);
    const storeKey = { providerId: options.providerId, scopesHash: normalized.scopesHash };
    const cacheKey = OAuth2Manager.keyString(storeKey);
    const now = options.now ?? this.now;
    const nowMs = now();

    const existingInFlight = this.inFlight.get(cacheKey);
    if (existingInFlight) return await existingInFlight;

    const cached = this.cache.get(cacheKey);
    if (!options.forceRefresh && this.isAccessTokenValid(cached, nowMs) && cached?.accessToken) {
      return { accessToken: cached.accessToken, expiresAtMs: cached.expiresAtMs, refreshToken: cached.refreshToken };
    }

    const promise = this.acquireAccessToken({
      provider,
      storeKey,
      scopes: normalized.scopes,
      forceRefresh: options.forceRefresh ?? false,
      signal: options.signal,
      now,
    });

    this.inFlight.set(cacheKey, promise);
    try {
      return await promise;
    } finally {
      if (this.inFlight.get(cacheKey) === promise) this.inFlight.delete(cacheKey);
    }
  }

  /**
   * @private
   * @param {{
   *   provider: OAuth2ProviderConfig;
   *   storeKey: OAuth2TokenStoreKey;
   *   scopes: string[];
   *   forceRefresh: boolean;
   *   signal?: AbortSignal;
   *   now: (() => number);
   * }} params
   * @returns {Promise<OAuth2AccessTokenResult>}
   */
  async acquireAccessToken(params) {
    const cacheKey = OAuth2Manager.keyString(params.storeKey);
    const nowMs = params.now();

    const cached = await this.hydrateCache(params.storeKey, params.scopes);
    if (!params.forceRefresh && this.isAccessTokenValid(cached, nowMs) && cached.accessToken) {
      return { accessToken: cached.accessToken, expiresAtMs: cached.expiresAtMs, refreshToken: cached.refreshToken };
    }

    if (!cached.refreshToken) {
      throw new Error(
        `No refresh token available for OAuth2 provider '${params.provider.id}'. ` +
          "Authenticate via an authorization-code/device-code flow before making requests.",
      );
    }

    const refreshed = await this.client.refreshToken({
      tokenEndpoint: params.provider.tokenEndpoint,
      clientId: params.provider.clientId,
      clientSecret: params.provider.clientSecret,
      refreshToken: cached.refreshToken,
      scopes: params.scopes,
      signal: params.signal,
    });

    const accessToken = refreshed.access_token;
    const expiresInSeconds = parsePositiveInt(refreshed.expires_in);
    const expiresAtMs = expiresInSeconds != null ? params.now() + expiresInSeconds * 1000 : null;
    const nextRefreshToken = refreshed.refresh_token ?? cached.refreshToken;

    cached.accessToken = accessToken;
    cached.expiresAtMs = expiresAtMs;
    cached.refreshToken = nextRefreshToken;
    cached.scopes = params.scopes;
    this.cache.set(cacheKey, cached);

    await this.persistTokens(params.storeKey, {
      providerId: params.provider.id,
      scopesHash: params.storeKey.scopesHash,
      scopes: params.scopes,
      refreshToken: nextRefreshToken,
      accessToken: this.persistAccessToken ? accessToken : undefined,
      expiresAtMs: this.persistAccessToken ? expiresAtMs : undefined,
    });

    return { accessToken, expiresAtMs, refreshToken: nextRefreshToken };
  }

  /**
   * Exchange an authorization code for tokens and persist the result.
   *
   * @param {{
   *   providerId: string;
   *   code: string;
   *   redirectUri?: string;
   *   codeVerifier?: string;
   *   scopes?: string[];
   *   signal?: AbortSignal;
   *   now?: (() => number);
   * }} options
   * @returns {Promise<OAuth2AccessTokenResult>}
   */
  async exchangeAuthorizationCode(options) {
    const provider = this.getProvider(options.providerId);
    const normalized = normalizeScopes(options.scopes ?? provider.defaultScopes);
    const storeKey = { providerId: options.providerId, scopesHash: normalized.scopesHash };
    const cacheKey = OAuth2Manager.keyString(storeKey);

    const existingInFlight = this.inFlight.get(cacheKey);
    if (existingInFlight) return await existingInFlight;

    const promise = this.acquireFromAuthorizationCode({
      provider,
      storeKey,
      code: options.code,
      redirectUri: options.redirectUri ?? provider.redirectUri ?? "",
      codeVerifier: options.codeVerifier,
      scopes: normalized.scopes,
      signal: options.signal,
      now: options.now ?? this.now,
    });

    this.inFlight.set(cacheKey, promise);
    try {
      return await promise;
    } finally {
      if (this.inFlight.get(cacheKey) === promise) this.inFlight.delete(cacheKey);
    }
  }

  /**
   * @private
   * @param {{
   *   provider: OAuth2ProviderConfig;
   *   storeKey: OAuth2TokenStoreKey;
   *   code: string;
   *   redirectUri: string;
   *   codeVerifier?: string;
   *   scopes: string[];
   *   signal?: AbortSignal;
   *   now: (() => number);
   * }} params
   * @returns {Promise<OAuth2AccessTokenResult>}
   */
  async acquireFromAuthorizationCode(params) {
    if (!params.redirectUri) {
      throw new Error(`OAuth2 provider '${params.provider.id}' requires a redirectUri for authorization-code exchange`);
    }

    const token = await this.client.exchangeCode({
      tokenEndpoint: params.provider.tokenEndpoint,
      clientId: params.provider.clientId,
      clientSecret: params.provider.clientSecret,
      code: params.code,
      redirectUri: params.redirectUri,
      codeVerifier: params.codeVerifier,
      signal: params.signal,
    });

    const accessToken = token.access_token;
    const expiresInSeconds = parsePositiveInt(token.expires_in);
    const expiresAtMs = expiresInSeconds != null ? params.now() + expiresInSeconds * 1000 : null;
    const refreshToken = token.refresh_token ?? null;

    const cacheKey = OAuth2Manager.keyString(params.storeKey);
    this.cache.set(cacheKey, { accessToken, expiresAtMs, refreshToken, scopes: params.scopes });

    await this.persistTokens(params.storeKey, {
      providerId: params.provider.id,
      scopesHash: params.storeKey.scopesHash,
      scopes: params.scopes,
      refreshToken,
      accessToken: this.persistAccessToken ? accessToken : undefined,
      expiresAtMs: this.persistAccessToken ? expiresAtMs : undefined,
    });

    return { accessToken, expiresAtMs, refreshToken };
  }

  /**
   * Start an interactive Authorization Code + PKCE flow.
   *
   * @param {{
   *   providerId: string;
   *   scopes?: string[];
   *   broker: OAuth2Broker;
   *   signal?: AbortSignal;
   *   now?: (() => number);
   * }} options
   * @returns {Promise<OAuth2AccessTokenResult>}
   */
  async authorizeWithPkce(options) {
    const provider = this.getProvider(options.providerId);
    if (!provider.authorizationEndpoint) {
      throw new Error(`OAuth2 provider '${provider.id}' does not define an authorizationEndpoint`);
    }
    if (!provider.redirectUri) {
      throw new Error(`OAuth2 provider '${provider.id}' does not define a redirectUri`);
    }
    if (!options.broker?.openAuthUrl) {
      throw new Error("OAuth2 PKCE authorization requires a broker with openAuthUrl(url)");
    }
    if (!options.broker.waitForRedirect) {
      throw new Error("OAuth2 PKCE authorization requires a broker with waitForRedirect(redirectUri)");
    }

    const normalized = normalizeScopes(options.scopes ?? provider.defaultScopes);
    const codeVerifier = await createCodeVerifier();
    const codeChallenge = await createCodeChallenge(codeVerifier);
    // OAuth2 `state` is a CSRF token; it does not need to follow PKCE verifier
    // length constraints. Reuse the same base64url random generator by using a
    // valid PKCE-sized nonce.
    const state = await createCodeVerifier();

    const authUrl = new URL(provider.authorizationEndpoint);
    authUrl.searchParams.set("response_type", "code");
    authUrl.searchParams.set("client_id", provider.clientId);
    authUrl.searchParams.set("redirect_uri", provider.redirectUri);
    if (normalized.scopes.length > 0) authUrl.searchParams.set("scope", normalized.scopes.join(" "));
    authUrl.searchParams.set("code_challenge", codeChallenge);
    authUrl.searchParams.set("code_challenge_method", "S256");
    authUrl.searchParams.set("state", state);
    if (provider.authorizationParams) {
      for (const [k, v] of Object.entries(provider.authorizationParams)) authUrl.searchParams.set(k, v);
    }

    await options.broker.openAuthUrl(authUrl.toString());
    const redirectUrl = await options.broker.waitForRedirect(provider.redirectUri);
    const parsedRedirect = new URL(redirectUrl);
    const returnedState = parsedRedirect.searchParams.get("state");
    if (returnedState && returnedState !== state) {
      throw new Error("OAuth2 redirect returned an unexpected state value");
    }

    const code = parsedRedirect.searchParams.get("code");
    if (!code) {
      const error = parsedRedirect.searchParams.get("error");
      if (error) {
        throw new Error(`OAuth2 authorization failed: ${error}`);
      }
      throw new Error("OAuth2 redirect did not include an authorization code");
    }

    return await this.exchangeAuthorizationCode({
      providerId: provider.id,
      code,
      redirectUri: provider.redirectUri,
      codeVerifier,
      scopes: normalized.scopes,
      signal: options.signal,
      now: options.now,
    });
  }

  /**
   * Start a device code flow and poll until completion.
   *
   * @param {{
   *   providerId: string;
   *   scopes?: string[];
   *   broker: OAuth2Broker;
   *   signal?: AbortSignal;
   *   now?: (() => number);
   * }} options
   * @returns {Promise<OAuth2AccessTokenResult>}
   */
  async authorizeWithDeviceCode(options) {
    const provider = this.getProvider(options.providerId);
    if (!provider.deviceAuthorizationEndpoint) {
      throw new Error(`OAuth2 provider '${provider.id}' does not define a deviceAuthorizationEndpoint`);
    }
    if (!options.broker?.openAuthUrl) {
      throw new Error("OAuth2 device code authorization requires a broker with openAuthUrl(url)");
    }

    const now = options.now ?? this.now;
    const normalized = normalizeScopes(options.scopes ?? provider.defaultScopes);
    const started = await this.client.deviceCodeStart({
      deviceAuthorizationEndpoint: provider.deviceAuthorizationEndpoint,
      clientId: provider.clientId,
      scopes: normalized.scopes,
      signal: options.signal,
    });

    const verificationUri = started.verification_uri_complete ?? started.verification_uri;
    if (options.broker.deviceCodePrompt) {
      await options.broker.deviceCodePrompt(started.user_code, verificationUri);
    }
    await options.broker.openAuthUrl(verificationUri);

    const deviceExpiresIn = parsePositiveInt(started.expires_in);
    if (deviceExpiresIn == null) {
      throw new Error("OAuth2 device code response missing expires_in");
    }
    const expiresAtMs = now() + deviceExpiresIn * 1000;
    const token = await this.client.deviceCodePoll({
      tokenEndpoint: provider.tokenEndpoint,
      clientId: provider.clientId,
      clientSecret: provider.clientSecret,
      deviceCode: started.device_code,
      intervalMs: (parsePositiveInt(started.interval) ?? 5) * 1000,
      expiresAtMs,
      signal: options.signal,
    });

    const accessToken = token.access_token;
    const tokenExpiresIn = parsePositiveInt(token.expires_in);
    const accessExpiresAtMs = tokenExpiresIn != null ? now() + tokenExpiresIn * 1000 : null;
    const refreshToken = token.refresh_token ?? null;

    const storeKey = { providerId: provider.id, scopesHash: normalized.scopesHash };
    const cacheKey = OAuth2Manager.keyString(storeKey);
    this.cache.set(cacheKey, { accessToken, expiresAtMs: accessExpiresAtMs, refreshToken, scopes: normalized.scopes });

    await this.persistTokens(storeKey, {
      providerId: provider.id,
      scopesHash: storeKey.scopesHash,
      scopes: normalized.scopes,
      refreshToken,
      accessToken: this.persistAccessToken ? accessToken : undefined,
      expiresAtMs: this.persistAccessToken ? accessExpiresAtMs : undefined,
    });

    return { accessToken, expiresAtMs: accessExpiresAtMs, refreshToken };
  }

  /**
   * @param {OAuth2TokenStoreKey} key
   * @param {OAuth2TokenStoreEntry} entry
   */
  async persistTokens(key, entry) {
    await this.tokenStore.set(key, entry);
  }

  /**
   * Clear cached + persisted tokens for a provider+scopes set.
   *
   * @param {{ providerId: string; scopes?: string[] }} options
   */
  async clearTokens(options) {
    const provider = this.getProvider(options.providerId);
    const normalized = normalizeScopes(options.scopes ?? provider.defaultScopes);
    const storeKey = { providerId: provider.id, scopesHash: normalized.scopesHash };
    const cacheKey = OAuth2Manager.keyString(storeKey);
    this.cache.delete(cacheKey);
    await this.tokenStore.delete(storeKey);
  }
}

/**
 * @param {unknown} value
 * @returns {number | null}
 */
function parsePositiveInt(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value === "string" && value.trim() !== "") {
    const parsed = Number(value);
    if (Number.isFinite(parsed)) return parsed;
  }
  return null;
}
