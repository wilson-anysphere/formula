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
}

export interface MarketplaceDownloadResult {
  bytes: Uint8Array;
  signatureBase64: string | null;
  sha256: string | null;
  formatVersion: number | null;
  publisher: string | null;
  publisherKeyId: string | null;
}

export interface MarketplaceClientOptions {
  /**
   * Base URL for marketplace API endpoints. Defaults to `"/api"`.
   *
   * The expected routes are:
   * - `${baseUrl}/search`
   * - `${baseUrl}/extensions/:id`
   * - `${baseUrl}/extensions/:id/download/:version`
   */
  baseUrl?: string;
}

function normalizeBaseUrl(baseUrl: string): string {
  const trimmed = String(baseUrl || "").trim();
  if (!trimmed) return "/api";
  return trimmed.endsWith("/") ? trimmed.slice(0, -1) : trimmed;
}

export class MarketplaceClient {
  readonly baseUrl: string;

  constructor(options: MarketplaceClientOptions = {}) {
    this.baseUrl = normalizeBaseUrl(options.baseUrl ?? "/api");
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
    const url = new URL(`${this.baseUrl}/extensions/${encodeURIComponent(id)}`, globalThis.location?.href ?? "http://localhost/");
    const res = await fetch(url.toString());
    if (res.status === 404) return null;
    if (!res.ok) {
      throw new Error(`Marketplace getExtension failed (${res.status})`);
    }
    return res.json();
  }

  async downloadPackage(id: string, version: string): Promise<MarketplaceDownloadResult | null> {
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
    const bytes = new Uint8Array(buf);

    const signatureBase64 = res.headers.get("x-package-signature");
    const sha256 = res.headers.get("x-package-sha256");
    const formatHeader = res.headers.get("x-package-format-version");
    const formatVersion =
      formatHeader && Number.isFinite(Number(formatHeader)) ? Number.parseInt(formatHeader, 10) : null;
    const publisher = res.headers.get("x-publisher");
    const publisherKeyId = res.headers.get("x-publisher-key-id");

    return {
      bytes,
      signatureBase64: signatureBase64 ? String(signatureBase64) : null,
      sha256: sha256 ? String(sha256) : null,
      formatVersion,
      publisher: publisher ? String(publisher) : null,
      publisherKeyId: publisherKeyId ? String(publisherKeyId) : null
    };
  }
}
