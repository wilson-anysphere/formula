/**
 * @typedef {"snapshot" | "checkpoint" | "restore"} VersionKind
 *
 * @typedef {{
 *   id: string;
 *   kind: VersionKind;
 *   timestampMs: number;
 *   userId: string | null;
 *   userName: string | null;
 *   description: string | null;
 *   checkpointName: string | null;
 *   checkpointLocked: boolean | null;
 *   checkpointAnnotations: string | null;
 *   snapshot: Uint8Array;
 * }} VersionRecord
 */

/**
 * The cloud API schema (`services/api` -> `document_versions`) stores an opaque
 * `data` blob plus a `description` string. To preserve the full `VersionRecord`
 * metadata we store a JSON envelope in `data`:
 *
 * ```json
 * {
 *   "schemaVersion": 1,
 *   "meta": { ...VersionRecordWithoutSnapshot },
 *   "snapshotBase64": "..."
 * }
 * ```
 *
 * This keeps the API backwards compatible while allowing the client to round-trip
 * `kind`, checkpoint fields, and author metadata.
 *
 * @typedef {{
 *   schemaVersion: 1;
 *   meta: Omit<VersionRecord, "snapshot">;
 *   snapshotBase64: string;
 * }} VersionEnvelopeV1
 */

function bytesToBase64(bytes) {
  // Node / desktop (Buffer available)
  // eslint-disable-next-line no-undef
  if (typeof Buffer !== "undefined") return Buffer.from(bytes).toString("base64");
  // Browser fallback
  let binary = "";
  for (let i = 0; i < bytes.length; i++) binary += String.fromCharCode(bytes[i]);
  // eslint-disable-next-line no-undef
  return btoa(binary);
}

function base64ToBytes(base64) {
  // eslint-disable-next-line no-undef
  if (typeof Buffer !== "undefined") return new Uint8Array(Buffer.from(base64, "base64"));
  // eslint-disable-next-line no-undef
  const binary = atob(base64);
  const out = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) out[i] = binary.charCodeAt(i);
  return out;
}

function encodeEnvelopeToDataBase64(version) {
  /** @type {VersionEnvelopeV1} */
  const envelope = {
    schemaVersion: 1,
    meta: {
      id: version.id,
      kind: version.kind,
      timestampMs: version.timestampMs,
      userId: version.userId,
      userName: version.userName,
      description: version.description,
      checkpointName: version.checkpointName,
      checkpointLocked: version.checkpointLocked,
      checkpointAnnotations: version.checkpointAnnotations
    },
    snapshotBase64: bytesToBase64(version.snapshot)
  };

  const json = JSON.stringify(envelope);
  const bytes = new TextEncoder().encode(json);
  return bytesToBase64(bytes);
}

/**
 * @param {string} dataBase64
 * @returns {VersionRecord}
 */
function decodeEnvelopeFromDataBase64(dataBase64) {
  const jsonBytes = base64ToBytes(dataBase64);
  const json = new TextDecoder().decode(jsonBytes);
  /** @type {unknown} */
  let parsed;
  try {
    parsed = JSON.parse(json);
  } catch (err) {
    throw new Error("ApiVersionStore: stored version payload is not valid JSON envelope");
  }

  if (!parsed || typeof parsed !== "object") {
    throw new Error("ApiVersionStore: stored version envelope is malformed");
  }

  const envelope = /** @type {any} */ (parsed);
  if (envelope.schemaVersion !== 1) {
    throw new Error(`ApiVersionStore: unsupported version envelope schemaVersion: ${envelope.schemaVersion}`);
  }
  if (!envelope.meta || typeof envelope.meta !== "object") {
    throw new Error("ApiVersionStore: stored version envelope is missing meta");
  }
  if (typeof envelope.snapshotBase64 !== "string") {
    throw new Error("ApiVersionStore: stored version envelope is missing snapshotBase64");
  }

  const meta = envelope.meta;
  return {
    id: meta.id,
    kind: meta.kind,
    timestampMs: meta.timestampMs,
    userId: meta.userId ?? null,
    userName: meta.userName ?? null,
    description: meta.description ?? null,
    checkpointName: meta.checkpointName ?? null,
    checkpointLocked: meta.checkpointLocked ?? null,
    checkpointAnnotations: meta.checkpointAnnotations ?? null,
    snapshot: base64ToBytes(envelope.snapshotBase64)
  };
}

/**
 * Cloud-backed `VersionStore` implementation that persists versions to the
 * Formula API (`services/api`) via `document_versions`.
 */
export class ApiVersionStore {
  /**
   * @param {{
   *   baseUrl: string;
   *   docId: string;
   *   auth: { cookie: string } | { bearerToken: string };
   *   fetchImpl?: typeof fetch;
   * }} opts
   */
  constructor(opts) {
    if (!opts?.baseUrl) throw new Error("ApiVersionStore: baseUrl is required");
    if (!opts?.docId) throw new Error("ApiVersionStore: docId is required");
    if (!opts?.auth) throw new Error("ApiVersionStore: auth is required");
    const hasCookie = "cookie" in opts.auth;
    const hasBearer = "bearerToken" in opts.auth;
    if (hasCookie === hasBearer) {
      throw new Error("ApiVersionStore: auth must provide exactly one of { cookie } or { bearerToken }");
    }

    this.baseUrl = opts.baseUrl;
    this.docId = opts.docId;
    this.auth = opts.auth;
    this.fetchImpl = opts.fetchImpl ?? fetch;
  }

  _headers(extra = {}) {
    /** @type {Record<string, string>} */
    const headers = { ...extra };
    if ("cookie" in this.auth) headers.cookie = this.auth.cookie;
    if ("bearerToken" in this.auth) headers.authorization = `Bearer ${this.auth.bearerToken}`;
    return headers;
  }

  _url(path) {
    return new URL(path, this.baseUrl).toString();
  }

  /**
   * @param {string} path
   * @param {{ method?: string, body?: any }} [opts]
   * @returns {Promise<any>}
   */
  async _requestJson(path, opts = {}) {
    const method = opts.method ?? "GET";
    const url = this._url(path);
    const res = await this.fetchImpl(url, {
      method,
      headers: this._headers({
        accept: "application/json",
        ...(opts.body ? { "content-type": "application/json" } : {})
      }),
      body: opts.body ? JSON.stringify(opts.body) : undefined
    });

    if (!res.ok) {
      let details = "";
      try {
        details = await res.text();
      } catch {
        details = "";
      }
      const suffix = details ? ` - ${details}` : "";
      throw new Error(`ApiVersionStore: ${method} ${url} failed: ${res.status}${suffix}`);
    }

    if (res.status === 204) return null;
    const text = await res.text();
    if (!text) return null;
    return JSON.parse(text);
  }

  /**
   * @param {VersionRecord} version
   */
  async saveVersion(version) {
    const dataBase64 = encodeEnvelopeToDataBase64(version);
    await this._requestJson(`/docs/${this.docId}/versions`, {
      method: "POST",
      body: {
        id: version.id,
        description: version.description,
        dataBase64
      }
    });
  }

  /**
   * @param {string} versionId
   * @returns {Promise<VersionRecord | null>}
   */
  async getVersion(versionId) {
    try {
      const res = await this._requestJson(`/docs/${this.docId}/versions/${versionId}`);
      const version = res?.version ?? res;
      if (!version) return null;
      if (typeof version.dataBase64 !== "string") {
        throw new Error("ApiVersionStore: GET version response is missing dataBase64");
      }
      return decodeEnvelopeFromDataBase64(version.dataBase64);
    } catch (err) {
      // Convert 404s into null, matching the store contract.
      if (err instanceof Error && err.message.includes(" failed: 404")) return null;
      throw err;
    }
  }

  /**
   * @returns {Promise<VersionRecord[]>}
   */
  async listVersions() {
    const res = await this._requestJson(`/docs/${this.docId}/versions`);
    const versions = res?.versions ?? [];
    if (!Array.isArray(versions)) {
      throw new Error("ApiVersionStore: list versions response is malformed");
    }

    /** @type {VersionRecord[]} */
    const out = [];
    for (const row of versions) {
      if (row && typeof row === "object" && typeof row.dataBase64 === "string") {
        out.push(decodeEnvelopeFromDataBase64(row.dataBase64));
        continue;
      }
      const id = row?.id;
      if (typeof id !== "string") {
        throw new Error("ApiVersionStore: list versions response is missing version id");
      }
      const full = await this.getVersion(id);
      if (full) out.push(full);
    }

    out.sort((a, b) => b.timestampMs - a.timestampMs);
    return out;
  }

  /**
   * @param {string} versionId
   * @param {{ checkpointLocked?: boolean }} patch
   */
  async updateVersion(versionId, patch) {
    if (patch.checkpointLocked === undefined) return;
    await this._requestJson(`/docs/${this.docId}/versions/${versionId}`, {
      method: "PATCH",
      body: { checkpointLocked: patch.checkpointLocked }
    });
  }

  /**
   * @param {string} versionId
   */
  async deleteVersion(versionId) {
    await this._requestJson(`/docs/${this.docId}/versions/${versionId}`, {
      method: "DELETE"
    });
  }
}

