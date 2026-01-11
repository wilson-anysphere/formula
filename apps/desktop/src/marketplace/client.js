export class MarketplaceClient {
  constructor({ baseUrl }) {
    if (!baseUrl) throw new Error("baseUrl is required");
    this.baseUrl = baseUrl.replace(/\/$/, "");
  }

  async search({
    q = "",
    category = "",
    tag = "",
    verified = undefined,
    featured = undefined,
    sort = "",
    limit = 20,
    offset = 0,
    cursor = "",
  } = {}) {
    const params = new URLSearchParams();
    if (q) params.set("q", q);
    if (category) params.set("category", category);
    if (tag) params.set("tag", tag);
    if (verified !== undefined) params.set("verified", String(verified));
    if (featured !== undefined) params.set("featured", String(featured));
    if (sort) params.set("sort", sort);
    params.set("limit", String(limit));
    params.set("offset", String(offset));
    if (cursor) params.set("cursor", cursor);

    const response = await fetch(`${this.baseUrl}/api/search?${params}`);
    if (!response.ok) throw new Error(`Search failed (${response.status})`);
    return response.json();
  }

  async getExtension(id) {
    const response = await fetch(`${this.baseUrl}/api/extensions/${encodeURIComponent(id)}`);
    if (response.status === 404) return null;
    if (!response.ok) throw new Error(`Get extension failed (${response.status})`);
    return response.json();
  }

  async downloadPackage(id, version) {
    const url = `${this.baseUrl}/api/extensions/${encodeURIComponent(id)}/download/${encodeURIComponent(version)}`;
    const response = await fetch(url);
    if (response.status === 404) return null;
    if (!response.ok) throw new Error(`Download failed (${response.status})`);

    const signatureBase64 = response.headers.get("x-package-signature");
    const sha256 = response.headers.get("x-package-sha256");
    const formatVersion = Number(response.headers.get("x-package-format-version") || "1");
    const publisher = response.headers.get("x-publisher");
    const bytes = Buffer.from(await response.arrayBuffer());

    return { bytes, signatureBase64, sha256, formatVersion, publisher };
  }
}
