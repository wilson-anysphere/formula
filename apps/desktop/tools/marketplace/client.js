// Deprecated / Node-only marketplace client used by Node test harnesses and legacy tooling.
// The production Desktop (Tauri/WebView) runtime uses `packages/extension-marketplace/src/MarketplaceClient.ts`.
//
// NOTE: This client expects `baseUrl` to be the marketplace origin (e.g. `https://marketplace.formula.app`), and it
// appends `/api/...` internally.
//
// To match common usage in the Desktop/Tauri runtime (which typically configures an *API* base URL that already includes
// `/api`), we also accept a trailing `/api` and normalize it away.

import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";

function normalizeMarketplaceOrigin(baseUrl) {
  const raw = String(baseUrl ?? "").trim();
  if (!raw) throw new Error("baseUrl is required");

  // Defensive normalization: query/hash are never meaningful for an origin-style base URL,
  // and leaving them in place can generate confusing URLs like:
  // `https://host/api?x=y/api/extensions/...`
  const withoutQueryHash = raw.split("#", 1)[0].split("?", 1)[0];
  const withoutTrailingSlash = withoutQueryHash.replace(/\/+$/, "");
  if (!withoutTrailingSlash) {
    throw new Error("baseUrl must be a non-empty URL (e.g. https://marketplace.formula.app)");
  }

  if (withoutTrailingSlash.endsWith("/api")) {
    const stripped = withoutTrailingSlash.slice(0, -4);
    if (!stripped) {
      throw new Error("baseUrl must be an origin URL, not a bare '/api' path");
    }
    return stripped;
  }

  return withoutTrailingSlash;
}

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
    this.baseUrl = normalizeMarketplaceOrigin(baseUrl);
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
    let cacheBase = null;
    let indexPath = null;
    if (this.cacheDir) {
      const safeId = safePathComponent(id);
      const safeVersion = safePathComponent(version);
      cacheBase = path.join(this.cacheDir, "packages", safeId, safeVersion);
      indexPath = path.join(cacheBase, "index.json");
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
            const headerPublisherKeyId = response.headers.get("x-publisher-key-id");
            const resolvedPublisherKeyId =
              headerPublisherKeyId && headerPublisherKeyId.trim().length > 0
                ? String(headerPublisherKeyId).trim()
                : cached.publisherKeyId || null;

            // 304 responses still include metadata headers; refresh our on-disk cache index
            // opportunistically so older caches can learn about newly-added headers.
            const refreshed = { ...cached };
            const headerSignature = response.headers.get("x-package-signature");
            const headerFormatVersion = response.headers.get("x-package-format-version");
            const headerPublisher = response.headers.get("x-publisher");
            const headerScanStatus = response.headers.get("x-package-scan-status");
            const headerEtag = response.headers.get("etag");
            const headerFilesSha = response.headers.get("x-package-files-sha256");

            if (headerShaNorm && isSha256Hex(headerShaNorm)) {
              refreshed.sha256 = headerShaNorm;
            }
            if (headerEtag && headerEtag.trim().length > 0) {
              refreshed.etag = String(headerEtag).trim();
            }
            if (headerSignature && headerSignature.trim().length > 0) {
              refreshed.signatureBase64 = String(headerSignature).trim();
            }
            if (headerPublisher && headerPublisher.trim().length > 0) {
              refreshed.publisher = String(headerPublisher).trim();
            }
            if (headerScanStatus && headerScanStatus.trim().length > 0) {
              refreshed.scanStatus = String(headerScanStatus).trim();
            }
            if (headerFormatVersion && headerFormatVersion.trim().length > 0) {
              const parsed = Number(headerFormatVersion);
              if (Number.isFinite(parsed) && parsed > 0) {
                refreshed.formatVersion = parsed;
              }
            }
            if (resolvedPublisherKeyId) {
              refreshed.publisherKeyId = resolvedPublisherKeyId;
            }
            if (headerFilesSha && headerFilesSha.trim().length > 0) {
              const normalized = String(headerFilesSha).trim().toLowerCase();
              if (isSha256Hex(normalized)) {
                refreshed.filesSha256 = normalized;
              }
            }

            const cacheIndexChanged =
              refreshed.sha256 !== cached.sha256 ||
              refreshed.etag !== cached.etag ||
              refreshed.signatureBase64 !== cached.signatureBase64 ||
              refreshed.publisher !== cached.publisher ||
              refreshed.scanStatus !== cached.scanStatus ||
              refreshed.formatVersion !== cached.formatVersion ||
              refreshed.publisherKeyId !== cached.publisherKeyId ||
              refreshed.filesSha256 !== cached.filesSha256;

            if (this.cacheDir && cacheBase && indexPath && cacheIndexChanged) {
              try {
                await atomicWriteJson(indexPath, refreshed);
            } catch {
              // ignore cache update failures
            }
          }

          return {
            bytes: Buffer.from(cachedBytes),
            signatureBase64: refreshed.signatureBase64 || null,
            sha256: String(refreshed.sha256 || cached.sha256),
            formatVersion: Number(refreshed.formatVersion || 1),
            publisher: refreshed.publisher || null,
            publisherKeyId: resolvedPublisherKeyId,
            scanStatus: refreshed.scanStatus || null,
            filesSha256: refreshed.filesSha256 || null,
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
    const scanStatus = response.headers.get("x-package-scan-status");
    const filesShaHeader = response.headers.get("x-package-files-sha256");
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
        scanStatus: scanStatus ? String(scanStatus).trim() : null,
        filesSha256: filesShaHeader ? String(filesShaHeader).trim().toLowerCase() : null,
      };
      const bytesPath = path.join(cacheBase, `${safePathComponent(computedSha)}.fextpkg`);
      try {
        await atomicWriteFile(bytesPath, bytes);
        await atomicWriteJson(path.join(cacheBase, "index.json"), entry);
      } catch {
        // Best-effort cache; failures should not break marketplace functionality.
      }
    }

    return {
      bytes,
      signatureBase64,
      sha256: expectedSha,
      formatVersion,
      publisher,
      publisherKeyId,
      scanStatus: scanStatus ? String(scanStatus).trim() : null,
      filesSha256: filesShaHeader ? String(filesShaHeader).trim().toLowerCase() : null,
    };
  }
}
