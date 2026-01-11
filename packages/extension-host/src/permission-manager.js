const fs = require("node:fs/promises");
const path = require("node:path");

class PermissionError extends Error {
  constructor(message) {
    super(message);
    this.name = "PermissionError";
  }
}

function normalizeStringArray(value) {
  if (!Array.isArray(value)) return [];
  return value
    .filter((entry) => typeof entry === "string")
    .map((entry) => entry.trim())
    .filter((entry) => entry.length > 0);
}

function normalizeNetworkPolicy(value) {
  if (!value) return null;
  if (value === true) return { mode: "full" };
  if (typeof value !== "object" || Array.isArray(value)) return null;

  const rawMode = value.mode;
  const mode =
    rawMode === "full" || rawMode === "deny" || rawMode === "allowlist" ? rawMode : undefined;

  const hosts = normalizeStringArray(value.hosts);
  if (mode === "full") return { mode: "full" };
  if (mode === "deny") return { mode: "deny" };
  if (mode === "allowlist") return hosts.length > 0 ? { mode: "allowlist", hosts } : { mode: "allowlist" };

  // If a future schema writes `hosts` without `mode`, treat it as allowlist.
  if (hosts.length > 0) return { mode: "allowlist", hosts };
  return null;
}

function normalizePermissionRecord(value) {
  // v1 format: ["cells.write", "network"]
  if (Array.isArray(value)) {
    const out = {};
    for (const perm of value) {
      if (typeof perm !== "string") continue;
      const name = perm.trim();
      if (!name) continue;
      if (name === "network") {
        // Backcompat: existing network grants become full access.
        out.network = { mode: "full" };
      } else {
        out[name] = true;
      }
    }
    return out;
  }

  if (!value || typeof value !== "object" || Array.isArray(value)) return {};

  const out = {};
  for (const [key, entry] of Object.entries(value)) {
    if (key === "network") {
      const normalized = normalizeNetworkPolicy(entry);
      if (normalized) out.network = normalized;
      continue;
    }

    if (entry === true) out[key] = true;
  }
  return out;
}

function normalizePermissionsStore(data) {
  if (!data || typeof data !== "object" || Array.isArray(data)) return { store: {}, migrated: false };
  const out = {};
  let migrated = false;

  for (const [extensionId, record] of Object.entries(data)) {
    const normalized = normalizePermissionRecord(record);
    out[extensionId] = normalized;
    if (record !== normalized) migrated = true;
  }

  return { store: out, migrated };
}

function collectDeclaredPermissions(declaredPermissions) {
  const list = Array.isArray(declaredPermissions) ? declaredPermissions : [];
  const out = new Set();
  for (const entry of list) {
    if (typeof entry === "string") {
      const trimmed = entry.trim();
      if (trimmed) out.add(trimmed);
      continue;
    }
    if (entry && typeof entry === "object" && !Array.isArray(entry)) {
      for (const key of Object.keys(entry)) {
        const trimmed = String(key).trim();
        if (trimmed) out.add(trimmed);
      }
    }
  }
  return out;
}

function safeParseUrl(url) {
  try {
    return new URL(String(url));
  } catch {
    return null;
  }
}

function isUrlAllowedByHosts(urlString, hosts) {
  const parsed = safeParseUrl(urlString);
  if (!parsed) return false;
  const origin = parsed.origin;
  const host = parsed.hostname;

  for (const rawEntry of normalizeStringArray(hosts)) {
    const entry = rawEntry.trim();
    if (!entry) continue;

    // Origin match, e.g. "https://api.example.com"
    if (entry.includes("://")) {
      if (origin === entry) return true;
      continue;
    }

    // Host wildcard match, e.g. "*.example.com"
    if (entry.startsWith("*.")) {
      const suffix = entry.slice(2);
      if (host === suffix) return true;
      if (host.endsWith(`.${suffix}`)) return true;
      continue;
    }

    if (host === entry) return true;
  }

  return false;
}

class PermissionManager {
  constructor({ storagePath, prompt }) {
    if (!storagePath) throw new Error("PermissionManager requires storagePath");
    this._storagePath = storagePath;
    this._prompt = typeof prompt === "function" ? prompt : async () => false;
    this._loaded = false;
    this._data = {};
    this._needsSave = false;
  }

  async _ensureLoaded() {
    if (this._loaded) return;
    try {
      const raw = await fs.readFile(this._storagePath, "utf8");
      const parsed = JSON.parse(raw);
      const { store, migrated } = normalizePermissionsStore(parsed);
      this._data = store;
      this._needsSave = migrated;
    } catch {
      this._data = {};
    }
    this._loaded = true;
    if (this._needsSave) {
      this._needsSave = false;
      try {
        await this._save();
      } catch {
        // ignore migration write failures
      }
    }
  }

  async _save() {
    await fs.mkdir(path.dirname(this._storagePath), { recursive: true });
    await fs.writeFile(this._storagePath, JSON.stringify(this._data, null, 2), "utf8");
  }

  async getGrantedPermissions(extensionId) {
    await this._ensureLoaded();
    const record = normalizePermissionRecord(this._data[extensionId]);
    this._data[extensionId] = record;
    return JSON.parse(JSON.stringify(record));
  }

  async revokePermissions(extensionId, permissions) {
    await this._ensureLoaded();
    const id = String(extensionId);
    const current = normalizePermissionRecord(this._data[id]);
    if (!this._data[id]) return;

    if (!Array.isArray(permissions) || permissions.length === 0) {
      delete this._data[id];
      await this._save();
      return;
    }

    const toRevoke = new Set(normalizeStringArray(permissions));
    let changed = false;
    for (const perm of toRevoke) {
      if (Object.prototype.hasOwnProperty.call(current, perm)) {
        delete current[perm];
        changed = true;
      }
    }

    if (!changed) return;
    this._data[id] = current;
    await this._save();
  }

  async resetAllPermissions() {
    await this._ensureLoaded();
    this._data = {};
    await this._save();
  }

  async ensurePermissions({ extensionId, displayName, declaredPermissions }, permissions, context = {}) {
    await this._ensureLoaded();
    const requested = Array.isArray(permissions) ? permissions : [];
    if (requested.length === 0) return true;

    const declared = collectDeclaredPermissions(declaredPermissions);
    for (const perm of requested) {
      if (!declared.has(perm)) {
        throw new PermissionError(`Permission not declared in manifest: ${perm}`);
      }
    }

    const record = normalizePermissionRecord(this._data[extensionId]);
    const apiKey = typeof context.apiKey === "string" ? context.apiKey : null;

    const networkUrl = typeof context?.network?.url === "string" ? context.network.url : null;
    const parsedUrl = networkUrl ? safeParseUrl(networkUrl) : null;
    const requestedHost = parsedUrl?.hostname ?? null;

    const needed = [];
    for (const perm of requested) {
      if (perm === "network") {
        const policy = normalizeNetworkPolicy(record.network);
        if (!policy) {
          needed.push(perm);
          continue;
        }

        if (policy.mode === "full") continue;
        if (policy.mode === "deny") {
          needed.push(perm);
          continue;
        }

        if (!networkUrl) {
          // Without a URL we can only assert that some network permission exists.
          continue;
        }

        if (!isUrlAllowedByHosts(networkUrl, policy.hosts ?? [])) {
          needed.push(perm);
        }
        continue;
      }

      if (record[perm] === true) continue;
      needed.push(perm);
    }

    if (needed.length === 0) return true;

    const accepted = await this._prompt({
      extensionId,
      displayName,
      permissions: needed,
      apiKey: apiKey ?? undefined,
      request: {
        apiKey: apiKey ?? undefined,
        permissions: needed,
        ...(needed.includes("network")
          ? {
              network: {
                url: networkUrl ?? undefined,
                host: requestedHost ?? undefined,
                mode:
                  normalizeNetworkPolicy(record.network)?.mode ??
                  (networkUrl ? "allowlist" : "full")
              }
            }
          : {})
      }
    });

    if (!accepted) {
      const detail =
        needed.length === 1 && needed[0] === "network" && requestedHost
          ? `network (${requestedHost})`
          : needed.join(", ");
      throw new PermissionError(`Permission denied: ${detail}`);
    }

    for (const perm of needed) {
      if (perm === "network") {
        if (!networkUrl) {
          record.network = { mode: "full" };
          continue;
        }

        const policy = normalizeNetworkPolicy(record.network);
        const nextHosts = new Set(normalizeStringArray(policy?.hosts));
        if (requestedHost) nextHosts.add(requestedHost);
        record.network = {
          mode: "allowlist",
          hosts: [...nextHosts].sort()
        };
        continue;
      }
      record[perm] = true;
    }

    this._data[extensionId] = record;
    await this._save();

    return true;
  }
}

module.exports = {
  PermissionError,
  PermissionManager
};
