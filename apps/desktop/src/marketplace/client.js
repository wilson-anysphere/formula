import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";

function sha256Hex(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

function isSha256Hex(value) {
  return typeof value === "string" && /^[0-9a-f]{64}$/i.test(value.trim());
}

function safePathComponent(value) {
  return String(value || "").replace(/[^a-zA-Z0-9._-]/g, "_");
}

async function readJsonIfExists(filePath) {
  try {
    const raw = await fs.readFile(filePath, "utf8");
    return JSON.parse(raw);
  } catch (error) {
    if (error && (error.code === "ENOENT" || error.code === "ENOTDIR")) return null;
    if (error instanceof SyntaxError) return null;
    throw error;
  }
}

async function atomicWriteJson(filePath, value) {
  await fs.mkdir(path.dirname(filePath), { recursive: true });
  const tmp = `${filePath}.${process.pid}.${Date.now()}.tmp`;
  await fs.writeFile(tmp, JSON.stringify(value, null, 2));
  try {
    await fs.rename(tmp, filePath);
  } catch (error) {
    if (error?.code === "EEXIST" || error?.code === "EPERM") {
      try {
        await fs.rm(filePath, { force: true });
        await fs.rename(tmp, filePath);
        return;
      } catch (renameError) {
        await fs.rm(tmp, { force: true }).catch(() => {});
        throw renameError;
      }
    }
    await fs.rm(tmp, { force: true }).catch(() => {});
    throw error;
  }
}

async function atomicWriteFile(filePath, bytes) {
  await fs.mkdir(path.dirname(filePath), { recursive: true });
  const tmp = `${filePath}.${process.pid}.${Date.now()}.tmp`;
  await fs.writeFile(tmp, bytes);
  try {
    await fs.rename(tmp, filePath);
  } catch (error) {
    if (error?.code === "EEXIST" || error?.code === "EPERM") {
      try {
        await fs.rm(filePath, { force: true });
        await fs.rename(tmp, filePath);
        return;
      } catch (renameError) {
        await fs.rm(tmp, { force: true }).catch(() => {});
        throw renameError;
      }
    }
    await fs.rm(tmp, { force: true }).catch(() => {});
    throw error;
  }
}

export class MarketplaceClient {
  constructor({ baseUrl, cacheDir = null }) {
    if (!baseUrl) throw new Error("baseUrl is required");
    this.baseUrl = baseUrl.replace(/\/$/, "");
    this.cacheDir = cacheDir ? path.resolve(cacheDir) : null;
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
    const url = `${this.baseUrl}/api/extensions/${encodeURIComponent(id)}`;

    let cached = null;
    if (this.cacheDir) {
      const safeId = safePathComponent(id);
      cached = await readJsonIfExists(path.join(this.cacheDir, "extensions", safeId, "index.json"));
      if (!cached || typeof cached !== "object" || !cached.body || typeof cached.body !== "object") {
        cached = null;
      }
    }

    const headers = cached?.etag ? { "If-None-Match": String(cached.etag) } : undefined;
    let response = await fetch(url, { headers });
    if (response.status === 304) {
      if (cached?.body) return cached.body;
      response = await fetch(url);
    }

    if (response.status === 404) return null;
    if (!response.ok) throw new Error(`Get extension failed (${response.status})`);
    const body = await response.json();

    if (this.cacheDir) {
      const safeId = safePathComponent(id);
      const cacheBase = path.join(this.cacheDir, "extensions", safeId);
      const etag = response.headers.get("etag");
      try {
        await atomicWriteJson(path.join(cacheBase, "index.json"), {
          etag: etag || null,
          body,
        });
      } catch {
        // Best-effort cache; failures should not break marketplace functionality.
      }
    }

    return body;
  }

  async downloadPackage(id, version) {
    const url = `${this.baseUrl}/api/extensions/${encodeURIComponent(id)}/download/${encodeURIComponent(version)}`;

    let cached = null;
    let cachedBytes = null;
    if (this.cacheDir) {
      const safeId = safePathComponent(id);
      const safeVersion = safePathComponent(version);
      const cacheBase = path.join(this.cacheDir, "packages", safeId, safeVersion);
      const indexPath = path.join(cacheBase, "index.json");
      cached = await readJsonIfExists(indexPath);
      const cachedSha256 =
        cached && typeof cached === "object" && typeof cached.sha256 === "string" ? cached.sha256.trim().toLowerCase() : null;
      if (cachedSha256 && isSha256Hex(cachedSha256)) {
        const cachedPath = path.join(cacheBase, `${safePathComponent(cachedSha256)}.fextpkg`);
        try {
          cachedBytes = await fs.readFile(cachedPath);
          const cachedSha = sha256Hex(cachedBytes);
          if (cachedSha !== cachedSha256) {
            cached = null;
            cachedBytes = null;
          } else {
            cached.sha256 = cachedSha256;
          }
        } catch {
          cached = null;
          cachedBytes = null;
        }
      } else {
        cached = null;
        cachedBytes = null;
      }
    }

    const headers = cached?.etag ? { "If-None-Match": cached.etag } : undefined;
    let response = await fetch(url, { headers });
    if (response.status === 404) return null;
    if (response.status === 304) {
      if (cached && cachedBytes) {
        const headerSha = response.headers.get("x-package-sha256");
        const headerShaNorm = headerSha ? String(headerSha).trim().toLowerCase() : null;
        if (headerShaNorm && (!isSha256Hex(headerShaNorm) || headerShaNorm !== String(cached.sha256).toLowerCase())) {
          response = await fetch(url);
        } else {
          return {
            bytes: Buffer.from(cachedBytes),
            signatureBase64: cached.signatureBase64 || null,
            sha256: cached.sha256,
            formatVersion: Number(cached.formatVersion || 1),
            publisher: cached.publisher || null,
            publisherKeyId: cached.publisherKeyId || null,
          };
        }
      } else {
        response = await fetch(url);
      }
    }
    if (response.status === 404) return null;
    if (!response.ok) throw new Error(`Download failed (${response.status})`);

    const signatureBase64 = response.headers.get("x-package-signature");
    const sha256 = response.headers.get("x-package-sha256");
    if (!sha256) {
      throw new Error("Marketplace download missing x-package-sha256 (mandatory)");
    }
    const expectedSha = String(sha256).trim().toLowerCase();
    if (!isSha256Hex(expectedSha)) {
      throw new Error("Marketplace download has invalid x-package-sha256 (expected 64-char hex)");
    }
    const formatVersion = Number(response.headers.get("x-package-format-version") || "1");
    const publisher = response.headers.get("x-publisher");
    const publisherKeyId = response.headers.get("x-publisher-key-id");
    const bytes = Buffer.from(await response.arrayBuffer());

    const computedSha = sha256Hex(bytes);
    if (computedSha !== expectedSha) {
      throw new Error(`Marketplace download sha256 mismatch: expected ${sha256} but got ${computedSha}`);
    }

    if (this.cacheDir) {
      const safeId = safePathComponent(id);
      const safeVersion = safePathComponent(version);
      const cacheBase = path.join(this.cacheDir, "packages", safeId, safeVersion);
      const etag = response.headers.get("etag");
      const entry = {
        sha256: computedSha,
        etag: etag || null,
        signatureBase64: signatureBase64 || null,
        formatVersion,
        publisher: publisher || null,
        publisherKeyId: publisherKeyId || null,
      };
      const bytesPath = path.join(cacheBase, `${safePathComponent(computedSha)}.fextpkg`);
      try {
        await atomicWriteFile(bytesPath, bytes);
        await atomicWriteJson(path.join(cacheBase, "index.json"), entry);
      } catch {
        // Best-effort cache; failures should not break marketplace functionality.
      }
    }

    return { bytes, signatureBase64, sha256: expectedSha, formatVersion, publisher, publisherKeyId };
  }
}
