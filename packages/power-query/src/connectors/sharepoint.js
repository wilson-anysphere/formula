import { DataTable } from "../table.js";

/**
 * @typedef {import("./types.js").ConnectorExecuteOptions} ConnectorExecuteOptions
 * @typedef {import("./types.js").ConnectorResult} ConnectorResult
 */

/**
 * @typedef {{ type: "oauth2"; providerId: string; scopes?: string[] | string }} SharePointConnectorOAuth2Config
 */

/**
 * @typedef {{
 *   siteUrl: string;
 *   mode: "contents" | "files";
 *   options?: {
 *     auth?: SharePointConnectorOAuth2Config | null;
 *     includeContent?: boolean;
 *     recursive?: boolean;
 *   };
 *   // Compatibility with `http:request` permission prompts (e.g. desktop app).
 *   url?: string;
 *   method?: string;
 * }} SharePointConnectorRequest
 */

/**
 * @typedef {Object} SharePointConnectorOptions
 * @property {typeof fetch | undefined} [fetch]
 * @property {import("../oauth2/manager.js").OAuth2Manager | undefined} [oauth2Manager]
 * @property {number[] | undefined} [oauth2RetryStatusCodes]
 */

/**
 * Normalize a SharePoint site URL into a canonical, stable string.
 *
 * - Forces `https:`
 * - Lower-cases the hostname
 * - Removes query/hash
 * - Removes a trailing slash (except for `/`)
 *
 * @param {string} input
 * @returns {string}
 */
function normalizeSiteUrl(input) {
  const parsed = new URL(String(input));
  const protocol = "https:";
  const hostname = parsed.hostname.toLowerCase();
  const port = parsed.port;
  const defaultPort = "443";
  const portSuffix = port && port !== defaultPort ? `:${port}` : "";
  let pathname = parsed.pathname || "/";
  pathname = pathname.replace(/\/{2,}/g, "/");
  if (pathname.length > 1 && pathname.endsWith("/")) pathname = pathname.slice(0, -1);
  return `${protocol}//${hostname}${portSuffix}${pathname}`;
}

/**
 * @param {string} siteUrl
 * @returns {string}
 */
function buildGraphSiteEndpoint(siteUrl) {
  const parsed = new URL(siteUrl);
  const hostname = parsed.hostname;
  let pathname = parsed.pathname || "/";
  pathname = pathname.replace(/\/{2,}/g, "/");
  if (pathname.length > 1 && pathname.endsWith("/")) pathname = pathname.slice(0, -1);
  // `encodeURI` preserves slashes but escapes spaces and other special characters.
  const encodedPath = encodeURI(pathname);
  return `https://graph.microsoft.com/v1.0/sites/${hostname}:${encodedPath}`;
}

/**
 * @param {unknown} value
 * @returns {Date | undefined}
 */
function parseDateTime(value) {
  if (typeof value !== "string") return undefined;
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return undefined;
  return parsed;
}

/**
 * @param {unknown} scopes
 * @returns {string[] | undefined}
 */
function coerceScopes(scopes) {
  if (Array.isArray(scopes)) return scopes;
  if (typeof scopes === "string") {
    const parts = scopes
      .split(/[\s,]+/)
      .map((s) => s.trim())
      .filter((s) => s.length > 0);
    return parts.length > 0 ? parts : undefined;
  }
  return undefined;
}

/**
 * @param {unknown} scopes
 * @returns {string[]}
 */
function normalizeScopesForCache(scopes) {
  const raw = Array.isArray(scopes) ? scopes : typeof scopes === "string" ? scopes.split(/[\s,]+/).filter(Boolean) : [];
  const cleaned = raw
    .filter((s) => typeof s === "string")
    .map((s) => s.trim())
    .filter((s) => s.length > 0);
  cleaned.sort();
  return Array.from(new Set(cleaned));
}

/**
 * @param {AbortSignal | undefined} signal
 */
function throwIfAborted(signal) {
  if (!signal?.aborted) return;
  const err = new Error("Aborted");
  err.name = "AbortError";
  throw err;
}

export class SharePointConnector {
  /**
   * @param {SharePointConnectorOptions} [options]
   */
  constructor(options = {}) {
    this.id = "sharepoint";
    // Keep permission semantics aligned with the underlying HTTP requests so host
    // apps can reuse their existing DLP + prompting UX.
    this.permissionKind = "http:request";
    this.fetchFn = options.fetch ?? (typeof fetch === "function" ? fetch.bind(globalThis) : null);
    this.oauth2Manager = options.oauth2Manager ?? null;
    this.oauth2RetryStatusCodes = Array.isArray(options.oauth2RetryStatusCodes) ? options.oauth2RetryStatusCodes : [401, 403];
  }

  /**
   * @param {SharePointConnectorRequest} request
   * @returns {unknown}
   */
  getCacheKey(request) {
    const siteUrl = normalizeSiteUrl(request.siteUrl);
    const mode = request.mode;
    const includeContent = request.options?.includeContent ?? false;
    const recursive = request.options?.recursive ?? false;
    const auth =
      request.options?.auth?.type === "oauth2"
        ? { type: "oauth2", providerId: request.options.auth.providerId, scopes: normalizeScopesForCache(request.options.auth.scopes) }
        : null;
    /** @type {any} */
    const key = {
      connector: "sharepoint",
      siteUrl,
      mode,
      options: {
        includeContent,
        recursive,
      },
    };
    // Avoid changing cache keys for unauthenticated callers by only including
    // `auth` when present.
    if (auth) {
      key.options.auth = auth;
    }
    return key;
  }

  /**
   * Lightweight source-state probe for cache validation.
   *
   * Uses a single request to the Graph "site" endpoint and extracts `etag` (when
   * present) plus `lastModifiedDateTime` from the response payload.
   *
   * @param {SharePointConnectorRequest} request
   * @param {ConnectorExecuteOptions} [options]
   * @returns {Promise<import("./types.js").SourceState>}
   */
  async getSourceState(request, options = {}) {
    const now = options.now ?? (() => Date.now());
    const signal = options.signal;
    throwIfAborted(signal);

    if (!this.fetchFn) return {};

    /** @type {Record<string, string>} */
    const headers = { Accept: "application/json" };

    let credentials = options.credentials;
    if (
      credentials &&
      typeof credentials === "object" &&
      !Array.isArray(credentials) &&
      // @ts-ignore - runtime access
      typeof credentials.getSecret === "function"
    ) {
      // @ts-ignore - runtime call
      credentials = await credentials.getSecret();
    }

    /** @type {{ providerId: string; scopes?: string[] } | null} */
    let credentialOAuth2 = null;
    if (credentials && typeof credentials === "object" && !Array.isArray(credentials)) {
      // @ts-ignore - runtime merge
      const extraHeaders = credentials.headers;
      if (extraHeaders && typeof extraHeaders === "object") {
        Object.assign(headers, extraHeaders);
      }

      // @ts-ignore - runtime access
      const oauth2 = credentials.oauth2;
      if (oauth2 && typeof oauth2 === "object") {
        // @ts-ignore - runtime access
        const providerId = oauth2.providerId;
        // @ts-ignore - runtime access
        const scopes = coerceScopes(oauth2.scopes);
        if (typeof providerId === "string" && providerId) {
          credentialOAuth2 = { providerId, scopes };
        }
      }
    }

    /** @type {{ providerId: string; scopes?: string[] } | null} */
    let oauth2 = null;
    if (request.options?.auth?.type === "oauth2") {
      oauth2 = { providerId: request.options.auth.providerId, scopes: coerceScopes(request.options.auth.scopes) };
    } else if (credentialOAuth2) {
      oauth2 = credentialOAuth2;
    }

    const applyOAuthHeader = async (forceRefresh = false) => {
      if (!oauth2) return;
      if (!this.oauth2Manager) {
        throw new Error("SharePoint OAuth2 requests require configuring SharePointConnector with an OAuth2Manager");
      }
      const token = await this.oauth2Manager.getAccessToken({
        providerId: oauth2.providerId,
        scopes: oauth2.scopes,
        signal,
        now,
        forceRefresh,
      });
      headers.Authorization = `Bearer ${token.accessToken}`;
    };

    await applyOAuthHeader(false);

    const siteUrl = normalizeSiteUrl(request.siteUrl);
    const siteEndpoint = `${buildGraphSiteEndpoint(siteUrl)}?$select=id,lastModifiedDateTime,webUrl`;

    const fetchOnce = async () => {
      const response = await this.fetchFn(siteEndpoint, { method: "GET", headers, signal });
      return response;
    };

    let response;
    try {
      response = await fetchOnce();
    } catch {
      return {};
    }

    if (!response.ok && oauth2 && this.oauth2RetryStatusCodes.includes(response.status)) {
      await applyOAuthHeader(true);
      try {
        response = await fetchOnce();
      } catch {
        return {};
      }
    }

    if (!response.ok) return {};

    const etag = response.headers.get("etag") ?? undefined;
    let sourceTimestamp;
    try {
      const json = await response.json();
      // @ts-ignore - runtime access
      sourceTimestamp = parseDateTime(json?.lastModifiedDateTime);
    } catch {
      sourceTimestamp = undefined;
    }

    return { etag, sourceTimestamp };
  }

  /**
   * @param {SharePointConnectorRequest} request
   * @param {ConnectorExecuteOptions} [options]
   * @returns {Promise<ConnectorResult>}
   */
  async execute(request, options = {}) {
    const now = options.now ?? (() => Date.now());
    const signal = options.signal;
    throwIfAborted(signal);

    if (!this.fetchFn) {
      throw new Error("SharePoint source requires a global fetch implementation or a SharePointConnector fetch adapter");
    }

    let credentials = options.credentials;
    if (
      credentials &&
      typeof credentials === "object" &&
      !Array.isArray(credentials) &&
      // @ts-ignore - runtime access
      typeof credentials.getSecret === "function"
    ) {
      // @ts-ignore - runtime call
      credentials = await credentials.getSecret();
    }

    /** @type {Record<string, string>} */
    const headers = { Accept: "application/json" };

    /** @type {{ providerId: string; scopes?: string[] } | null} */
    let credentialOAuth2 = null;
    if (credentials && typeof credentials === "object" && !Array.isArray(credentials)) {
      // @ts-ignore - runtime merge
      const extraHeaders = credentials.headers;
      if (extraHeaders && typeof extraHeaders === "object") {
        Object.assign(headers, extraHeaders);
      }

      // @ts-ignore - runtime access
      const oauth2 = credentials.oauth2;
      if (oauth2 && typeof oauth2 === "object") {
        // @ts-ignore - runtime access
        const providerId = oauth2.providerId;
        // @ts-ignore - runtime access
        const scopes = coerceScopes(oauth2.scopes);
        if (typeof providerId === "string" && providerId) {
          credentialOAuth2 = { providerId, scopes };
        }
      }
    }

    /** @type {{ providerId: string; scopes?: string[] } | null} */
    let oauth2 = null;
    if (request.options?.auth?.type === "oauth2") {
      oauth2 = { providerId: request.options.auth.providerId, scopes: coerceScopes(request.options.auth.scopes) };
    } else if (credentialOAuth2) {
      oauth2 = credentialOAuth2;
    }

    const applyOAuthHeader = async (forceRefresh = false) => {
      if (!oauth2) return;
      if (!this.oauth2Manager) {
        throw new Error("SharePoint OAuth2 requests require configuring SharePointConnector with an OAuth2Manager");
      }
      const token = await this.oauth2Manager.getAccessToken({
        providerId: oauth2.providerId,
        scopes: oauth2.scopes,
        signal,
        now,
        forceRefresh,
      });
      headers.Authorization = `Bearer ${token.accessToken}`;
    };

    await applyOAuthHeader(false);

    const fetchResponseWithRetry = async (url, init) => {
      const response = await this.fetchFn(url, init);
      if (!response.ok && oauth2 && this.oauth2RetryStatusCodes.includes(response.status)) {
        await applyOAuthHeader(true);
        return await this.fetchFn(url, init);
      }
      return response;
    };

    /**
     * @param {string} url
     * @returns {Promise<any>}
     */
    const fetchJson = async (url) => {
      const response = await fetchResponseWithRetry(url, { method: "GET", headers, signal });
      if (!response.ok) throw new Error(`HTTP ${response.status} for ${url}`);
      return await response.json();
    };

    /**
     * @param {string} url
     * @returns {Promise<any[]>}
     */
    const fetchPaged = async (url) => {
      /** @type {any[]} */
      const items = [];
      let next = url;
      while (next) {
        const data = await fetchJson(next);
        const page = Array.isArray(data?.value) ? data.value : [];
        items.push(...page);
        // @ts-ignore - Graph uses `@odata.nextLink`
        const nextLink = data?.["@odata.nextLink"];
        next = typeof nextLink === "string" && nextLink.length > 0 ? nextLink : "";
      }
      return items;
    };

    const siteUrl = normalizeSiteUrl(request.siteUrl);
    const siteEndpoint = `${buildGraphSiteEndpoint(siteUrl)}?$select=id,lastModifiedDateTime,webUrl`;

    /** @type {{ id?: string; lastModifiedDateTime?: string }} */
    const site = await fetchJson(siteEndpoint);
    const siteId = typeof site?.id === "string" ? site.id : null;
    if (!siteId) throw new Error(`SharePoint site resolution failed for ${siteUrl}`);

    const siteLastModified = parseDateTime(site?.lastModifiedDateTime);

    if (request.mode === "contents") {
      const drives = await fetchPaged(`https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(siteId)}/drives?$select=id,name,webUrl,driveType`);
      const rows = drives.map((d) => [
        typeof d?.name === "string" ? d.name : null,
        typeof d?.id === "string" ? d.id : null,
        typeof d?.webUrl === "string" ? d.webUrl : null,
        typeof d?.driveType === "string" ? d.driveType : null,
      ]);
      const table = new DataTable(
        [
          { name: "Name", type: "string" },
          { name: "Id", type: "string" },
          { name: "WebUrl", type: "string" },
          { name: "DriveType", type: "string" },
        ],
        rows,
      );
      return {
        table,
        meta: {
          refreshedAt: new Date(now()),
          sourceTimestamp: siteLastModified,
          schema: { columns: table.columns, inferred: true },
          rowCount: table.rows.length,
          rowCountEstimate: table.rows.length,
          provenance: { kind: "sharepoint", siteUrl, mode: "contents" },
        },
      };
    }

    if (request.mode === "files") {
      const includeContent = request.options?.includeContent ?? false;
      const recursive = request.options?.recursive ?? false;

      const drives = await fetchPaged(`https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(siteId)}/drives?$select=id,name,webUrl,driveType`);

      /** @type {unknown[][]} */
      const outRows = [];

      /**
       * @param {string} driveId
       * @param {string} itemId
       * @param {string} folderPath
       */
      const walkFolder = async (driveId, itemId, folderPath) => {
        const base =
          itemId === "root"
            ? `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/root/children?$select=id,name,webUrl,size,file,folder,parentReference,lastModifiedDateTime,createdDateTime`
            : `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/children?$select=id,name,webUrl,size,file,folder,parentReference,lastModifiedDateTime,createdDateTime`;
        const children = await fetchPaged(base);
        for (const child of children) {
          const name = typeof child?.name === "string" ? child.name : "";
          const isFolder = Boolean(child?.folder);
          const isFile = Boolean(child?.file);
          if (isFolder && recursive) {
            await walkFolder(driveId, String(child.id ?? ""), `${folderPath}${name}/`);
          }
          if (!isFile) continue;

          const extension = name.includes(".") ? name.slice(name.lastIndexOf(".") + 1) : "";

          /** @type {Uint8Array | null} */
          let contentBytes = null;
          if (includeContent && typeof child?.id === "string" && child.id) {
            const contentUrl = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(child.id)}/content`;
            let response = await fetchResponseWithRetry(contentUrl, { method: "GET", headers, signal });
            if (response.status >= 300 && response.status < 400) {
              const redirect = response.headers.get("location");
              if (redirect) {
                response = await fetchResponseWithRetry(redirect, { method: "GET", headers, signal });
              }
            }
            if (!response.ok) throw new Error(`HTTP ${response.status} for ${contentUrl}`);
            const buf = await response.arrayBuffer();
            contentBytes = new Uint8Array(buf);
          }

          const lastModified = parseDateTime(child?.lastModifiedDateTime);
          const created = parseDateTime(child?.createdDateTime);

          const attributes = {
            // Mirror the shape callers often expect from SharePoint.Files().
            Size: typeof child?.size === "number" ? child.size : null,
            WebUrl: typeof child?.webUrl === "string" ? child.webUrl : null,
            DriveId: driveId,
          };

          outRows.push([
            includeContent ? contentBytes : null,
            name || null,
            extension || null,
            null,
            lastModified ?? null,
            created ?? null,
            attributes,
            folderPath,
          ]);
        }
      };

      for (const drive of drives) {
        const driveId = typeof drive?.id === "string" ? drive.id : null;
        if (!driveId) continue;
        await walkFolder(driveId, "root", `${siteUrl}/`);
      }

      const table = new DataTable(
        [
          { name: "Content", type: "any" },
          { name: "Name", type: "string" },
          { name: "Extension", type: "string" },
          { name: "Date accessed", type: "date" },
          { name: "Date modified", type: "date" },
          { name: "Date created", type: "date" },
          { name: "Attributes", type: "any" },
          { name: "Folder Path", type: "string" },
        ],
        outRows,
      );

      return {
        table,
        meta: {
          refreshedAt: new Date(now()),
          sourceTimestamp: siteLastModified,
          schema: { columns: table.columns, inferred: true },
          rowCount: table.rows.length,
          rowCountEstimate: table.rows.length,
          provenance: { kind: "sharepoint", siteUrl, mode: "files" },
        },
      };
    }

    /** @type {never} */
    const exhausted = request.mode;
    throw new Error(`Unsupported SharePoint mode '${exhausted}'`);
  }
}

