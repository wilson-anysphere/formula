export interface MarketplaceSearchResult<T = MarketplaceExtensionSummary> {
  total: number;
  results: T[];
  nextCursor: string | null;
}

export interface MarketplaceExtensionSummary {
  id: string;
  name: string;
  displayName: string;
  publisher: string;
  description: string;
  latestVersion: string | null;
  verified: boolean;
  featured: boolean;
  deprecated?: boolean;
  blocked?: boolean;
  malicious?: boolean;
  publisherRevoked?: boolean;
  categories: string[];
  tags: string[];
  screenshots: string[];
  downloadCount: number;
  updatedAt: string;
}

export interface MarketplaceExtensionVersion {
  version: string;
  sha256: string;
  uploadedAt: string;
  yanked: boolean;
  scanStatus?: string;
  signingKeyId?: string | null;
  formatVersion?: number;
}

export interface MarketplacePublisherKey {
  id: string;
  publicKeyPem: string;
  revoked: boolean;
}

export interface MarketplaceExtensionDetails extends MarketplaceExtensionSummary {
  versions: MarketplaceExtensionVersion[];
  readme: string;
  publisherPublicKeyPem: string | null;
  publisherKeys?: MarketplacePublisherKey[];
  createdAt: string;
  deprecated: boolean;
  blocked: boolean;
  malicious: boolean;
  publisherRevoked?: boolean;
  publisherRevokedAt?: string | null;
}

export interface MarketplaceDownloadResult {
  // Downloads are materialized via `Response.arrayBuffer()`, so bytes are always backed by an `ArrayBuffer`.
  bytes: Uint8Array<ArrayBuffer>;
  signatureBase64: string | null;
  sha256: string | null;
  formatVersion: number | null;
  publisher: string | null;
  publisherKeyId: string | null;
  scanStatus: string | null;
  filesSha256: string | null;
}

export interface MarketplaceClientOptions {
  /**
   * Base URL for marketplace API endpoints.
   *
   * Defaults to:
   * - `import.meta.env.VITE_FORMULA_MARKETPLACE_BASE_URL` when present (Vite)
   * - `process.env.VITE_FORMULA_MARKETPLACE_BASE_URL` when present (Node tests/tooling)
   * - `"/api"` otherwise
   *
   * The expected routes are:
   * - `${baseUrl}/search`
   * - `${baseUrl}/extensions/:id`
   * - `${baseUrl}/extensions/:id/download/:version`
   */
  baseUrl?: string;
}

export function normalizeMarketplaceBaseUrl(baseUrl: string): string {
  let raw = String(baseUrl ?? "").trim();
  if (!raw) return "/api";

  raw = raw.replace(/\\/g, "/");
  // Base URL should not carry query/hash. (These break `${baseUrl}/search` concatenation.)
  raw = raw.split("#", 1)[0]!.split("?", 1)[0]!;

  // Strip trailing slashes.
  //
  // Note: callers sometimes pass "/" to mean "the default API base". Treat that as "/api"
  // rather than returning "/" (which would produce "//search" when constructing endpoint URLs).
  raw = raw.replace(/\/+$/, "");
  if (raw === "") return "/api";

  const looksAbsolute = /^[a-zA-Z][a-zA-Z0-9+.-]*:\/\//.test(raw);
  if (looksAbsolute) {
    let url: URL;
    try {
      url = new URL(raw);
    } catch {
      return "/api";
    }

    // Base URL should not carry query/hash.
    url.search = "";
    url.hash = "";

    let pathname = url.pathname.replace(/\/+$/, "");
    // Treat an origin ("https://host") as the marketplace host and append the standard "/api" prefix.
    if (pathname === "" || pathname === "/") pathname = "/api";
    url.pathname = pathname;

    return `${url.origin}${url.pathname}`;
  }

  // Normalize relative paths (typically "/api").
  let out = raw;
  while (out.startsWith("./")) out = out.slice(2);
  if (!out.startsWith("/")) out = `/${out}`;
  // Collapse multiple leading slashes so protocol-relative forms like "//host/api" don't
  // accidentally get treated as absolute hosts by `new URL(...)`.
  out = out.replace(/^\/+/, "/");
  out = out.replace(/\/+$/, "");
  if (out === "") return "/api";
  return out;
}

function resolveDefaultMarketplaceBaseUrl(): string {
  const metaEnv = (import.meta as any)?.env as Record<string, unknown> | undefined;
  const viteValue = metaEnv?.VITE_FORMULA_MARKETPLACE_BASE_URL;
  if (typeof viteValue === "string" && viteValue.trim().length > 0) {
    return viteValue;
  }

  const nodeEnv = (globalThis as any)?.process?.env as Record<string, unknown> | undefined;
  const nodeValue = nodeEnv?.VITE_FORMULA_MARKETPLACE_BASE_URL;
  if (typeof nodeValue === "string" && nodeValue.trim().length > 0) {
    return nodeValue;
  }

  return "/api";
}

async function sha256Hex(bytes: Uint8Array): Promise<string> {
  const subtle = globalThis.crypto?.subtle;
  if (!subtle?.digest) {
    throw new Error("Marketplace client requires crypto.subtle.digest() to verify downloads");
  }

  // `crypto.subtle.digest` expects a BufferSource backed by an `ArrayBuffer`. TypeScript models
  // `Uint8Array` as potentially backed by a `SharedArrayBuffer` (`ArrayBufferLike`), so normalize
  // to an `ArrayBuffer`-backed view for type safety.
  const normalized: Uint8Array<ArrayBuffer> =
    bytes.buffer instanceof ArrayBuffer ? (bytes as Uint8Array<ArrayBuffer>) : new Uint8Array(bytes);

  const hash = new Uint8Array(await subtle.digest("SHA-256", normalized));
  let out = "";
  for (const b of hash) out += b.toString(16).padStart(2, "0");
  return out;
}

type TauriInvoke = (cmd: string, args?: Record<string, any>) => Promise<any>;

function getTauriInvoke(): TauriInvoke | null {
  const tauri = (globalThis as any)?.__TAURI__;
  const invoke: unknown = tauri?.core?.invoke ?? tauri?.invoke;
  if (typeof invoke !== "function") return null;
  return invoke.bind(tauri?.core ?? tauri);
}

function isAbsoluteHttpUrl(value: string): boolean {
  try {
    const url = new URL(value);
    return url.protocol === "http:" || url.protocol === "https:";
  } catch {
    return false;
  }
}

function base64ToBytes(base64: string): Uint8Array<ArrayBuffer> {
  const raw = String(base64 ?? "");
  if (raw.length === 0) {
    throw new Error("Invalid base64 payload (empty)");
  }

  if (typeof Buffer !== "undefined") {
    const buf = Buffer.from(raw, "base64");
    const sliced = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
    return new Uint8Array(sliced);
  }

  if (typeof atob === "function") {
    const bin = atob(raw);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return bytes;
  }

  throw new Error("Base64 decoding is not available in this runtime");
}

export class MarketplaceClient {
  readonly baseUrl: string;

  constructor(options: MarketplaceClientOptions = {}) {
    this.baseUrl = normalizeMarketplaceBaseUrl(options.baseUrl ?? resolveDefaultMarketplaceBaseUrl());
  }

  async search(params: {
    q?: string;
    category?: string;
    tag?: string;
    verified?: boolean;
    featured?: boolean;
    sort?: string;
    limit?: number;
    offset?: number;
    cursor?: string | null;
  }): Promise<MarketplaceSearchResult> {
    const invoke = getTauriInvoke();
    if (invoke && isAbsoluteHttpUrl(this.baseUrl)) {
      return invoke("marketplace_search", {
        baseUrl: this.baseUrl,
        q: params.q,
        category: params.category,
        tag: params.tag,
        verified: params.verified,
        featured: params.featured,
        sort: params.sort,
        limit: params.limit,
        offset: params.offset,
        cursor: params.cursor ?? undefined
      });
    }

    const url = new URL(`${this.baseUrl}/search`, globalThis.location?.href ?? "http://localhost/");
    if (params.q) url.searchParams.set("q", params.q);
    if (params.category) url.searchParams.set("category", params.category);
    if (params.tag) url.searchParams.set("tag", params.tag);
    if (params.verified !== undefined) url.searchParams.set("verified", params.verified ? "true" : "false");
    if (params.featured !== undefined) url.searchParams.set("featured", params.featured ? "true" : "false");
    if (params.sort) url.searchParams.set("sort", params.sort);
    if (params.limit !== undefined) url.searchParams.set("limit", String(params.limit));
    if (params.offset !== undefined) url.searchParams.set("offset", String(params.offset));
    if (params.cursor) url.searchParams.set("cursor", params.cursor);

    const res = await fetch(url.toString());
    if (!res.ok) {
      throw new Error(`Marketplace search failed (${res.status})`);
    }
    return res.json();
  }

  async getExtension(id: string): Promise<MarketplaceExtensionDetails | null> {
    const invoke = getTauriInvoke();
    if (invoke && isAbsoluteHttpUrl(this.baseUrl)) {
      const result = await invoke("marketplace_get_extension", { baseUrl: this.baseUrl, id: String(id) });
      if (!result) return null;
      return result as MarketplaceExtensionDetails;
    }

    const url = new URL(
      `${this.baseUrl}/extensions/${encodeURIComponent(id)}`,
      globalThis.location?.href ?? "http://localhost/"
    );
    const res = await fetch(url.toString());
    if (res.status === 404) return null;
    if (!res.ok) {
      throw new Error(`Marketplace getExtension failed (${res.status})`);
    }
    return res.json();
  }

  async downloadPackage(id: string, version: string): Promise<MarketplaceDownloadResult | null> {
    const invoke = getTauriInvoke();
    if (invoke && isAbsoluteHttpUrl(this.baseUrl)) {
      const payload = await invoke("marketplace_download_package", {
        baseUrl: this.baseUrl,
        id: String(id),
        version: String(version)
      });
      if (!payload) return null;

      const bytes = base64ToBytes(String(payload.bytesBase64 ?? ""));

      const signatureBase64 = payload.signatureBase64;
      const sha256 = payload.sha256;
      if (!sha256) {
        throw new Error("Marketplace download missing x-package-sha256 (mandatory)");
      }
      const expectedSha = String(sha256).trim().toLowerCase();
      if (!/^[0-9a-f]{64}$/i.test(expectedSha)) {
        throw new Error("Marketplace download has invalid x-package-sha256 (expected 64-char hex)");
      }
      const computedSha = await sha256Hex(bytes);
      if (computedSha !== expectedSha) {
        throw new Error(`Marketplace download sha256 mismatch: expected ${expectedSha} but got ${computedSha}`);
      }

      const formatVersion =
        typeof payload.formatVersion === "number" && Number.isFinite(payload.formatVersion)
          ? Number(payload.formatVersion)
          : null;

      return {
        bytes,
        signatureBase64: signatureBase64 ? String(signatureBase64) : null,
        sha256: expectedSha,
        formatVersion,
        publisher: payload.publisher ? String(payload.publisher) : null,
        publisherKeyId: payload.publisherKeyId ? String(payload.publisherKeyId) : null,
        scanStatus: payload.scanStatus ? String(payload.scanStatus) : null,
        filesSha256: payload.filesSha256 ? String(payload.filesSha256) : null,
      };
    }

    const url = new URL(
      `${this.baseUrl}/extensions/${encodeURIComponent(id)}/download/${encodeURIComponent(version)}`,
      globalThis.location?.href ?? "http://localhost/"
    );
    const res = await fetch(url.toString());
    if (res.status === 404) return null;
    if (!res.ok) {
      throw new Error(`Marketplace download failed (${res.status})`);
    }

    const buf = await res.arrayBuffer();
    const bytes: Uint8Array<ArrayBuffer> = new Uint8Array(buf);

    const signatureBase64 = res.headers.get("x-package-signature");
    const sha256 = res.headers.get("x-package-sha256");
    if (!sha256) {
      throw new Error("Marketplace download missing x-package-sha256 (mandatory)");
    }
    const expectedSha = String(sha256).trim().toLowerCase();
    if (!/^[0-9a-f]{64}$/i.test(expectedSha)) {
      throw new Error("Marketplace download has invalid x-package-sha256 (expected 64-char hex)");
    }
    const computedSha = await sha256Hex(bytes);
    if (computedSha !== expectedSha) {
      throw new Error(`Marketplace download sha256 mismatch: expected ${expectedSha} but got ${computedSha}`);
    }
    const formatHeader = res.headers.get("x-package-format-version");
    const formatVersion =
      formatHeader && Number.isFinite(Number(formatHeader)) ? Number.parseInt(formatHeader, 10) : null;
    const publisher = res.headers.get("x-publisher");
    const publisherKeyId = res.headers.get("x-publisher-key-id");
    const scanStatus = res.headers.get("x-package-scan-status");
    const filesSha256 = res.headers.get("x-package-files-sha256");

    return {
      bytes,
      signatureBase64: signatureBase64 ? String(signatureBase64) : null,
      sha256: expectedSha,
      formatVersion,
      publisher: publisher ? String(publisher) : null,
      publisherKeyId: publisherKeyId ? String(publisherKeyId) : null,
      scanStatus: scanStatus ? String(scanStatus) : null,
      filesSha256: filesSha256 ? String(filesSha256) : null,
    };
  }
}
