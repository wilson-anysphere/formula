class PermissionError extends Error {
  constructor(message) {
    super(message);
    this.name = "PermissionError";
  }
}

function deepEqual(a, b) {
  if (a === b) return true;
  if (a == null || b == null) return false;
  if (typeof a !== typeof b) return false;

  if (Array.isArray(a) || Array.isArray(b)) {
    if (!Array.isArray(a) || !Array.isArray(b)) return false;
    if (a.length !== b.length) return false;
    for (let i = 0; i < a.length; i += 1) {
      if (!deepEqual(a[i], b[i])) return false;
    }
    return true;
  }

  if (typeof a === "object") {
    const aKeys = Object.keys(a);
    const bKeys = Object.keys(b);
    if (aKeys.length !== bKeys.length) return false;
    aKeys.sort();
    bKeys.sort();
    for (let i = 0; i < aKeys.length; i += 1) {
      if (aKeys[i] !== bKeys[i]) return false;
    }
    for (const key of aKeys) {
      if (!deepEqual(a[key], b[key])) return false;
    }
    return true;
  }

  return false;
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

  if (hosts.length > 0) return { mode: "allowlist", hosts };
  return null;
}

function normalizePermissionRecord(value) {
  if (Array.isArray(value)) {
    const out = {};
    for (const perm of value) {
      if (typeof perm !== "string") continue;
      const name = perm.trim();
      if (!name) continue;
      if (name === "network") {
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
  if (!data || typeof data !== "object" || Array.isArray(data)) {
    return { store: Object.create(null), migrated: false };
  }
  const out = Object.create(null);
  let migrated = false;

  try {
    if (Object.getPrototypeOf(data) !== Object.prototype) {
      migrated = true;
    }
  } catch {
    migrated = true;
  }

  for (const [extensionId, record] of Object.entries(data)) {
    const normalized = normalizePermissionRecord(record);
    if (Object.keys(normalized).length === 0) {
      migrated = true;
      continue;
    }
    out[extensionId] = normalized;
    if (!deepEqual(record, normalized)) migrated = true;
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

    if (entry.includes("://")) {
      if (origin === entry) return true;
      continue;
    }

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

function getDefaultLocalStorage() {
  try {
    if (typeof globalThis === "undefined") return null;
    return globalThis.localStorage ?? null;
  } catch {
    return null;
  }
}

class PermissionManager {
  constructor({ prompt, storage, storageKey = "formula.extensionHost.permissions" } = {}) {
    this._prompt = typeof prompt === "function" ? prompt : async () => false;
    this._storage = storage ?? getDefaultLocalStorage();
    this._storageKey = String(storageKey);
    this._loaded = false;
    this._data = Object.create(null);
    this._needsSave = false;
  }

  _ensureLoaded() {
    if (this._loaded) return;
    this._loaded = true;

    if (!this._storage) {
      this._data = {};
      return;
    }

    try {
      const raw = this._storage.getItem(this._storageKey);
      const hadRaw = raw != null;
      const parsed = hadRaw ? JSON.parse(raw) : {};
      const { store, migrated } = normalizePermissionsStore(parsed);
      this._data = store;
      // Treat any persisted empty/invalid store as a migration so we can rewrite storage to a clean
      // slate (and remove the key entirely when the store is empty).
      this._needsSave = migrated || (hadRaw && Object.keys(store).length === 0);
    } catch {
      this._data = Object.create(null);
      // If the stored value is corrupted (invalid JSON), treat it as a migration so we rewrite
      // storage to a clean slate (and remove the key entirely when the store is empty).
      this._needsSave = true;
    }
    if (this._needsSave) {
      this._needsSave = false;
      this._save();
    }
  }

  _save() {
    if (!this._storage) return;
    try {
      // When the store becomes empty, remove the key entirely so uninstall/reset flows leave a
      // clean slate in storage (avoids persisting a noisy "{}" record).
      if (Object.keys(this._data).length === 0) {
        if (typeof this._storage.removeItem === "function") {
          this._storage.removeItem(this._storageKey);
        } else {
          this._storage.setItem(this._storageKey, JSON.stringify(this._data));
        }
        return;
      }

      this._storage.setItem(this._storageKey, JSON.stringify(this._data));
    } catch {
      // Ignore storage failures (quota, disabled localStorage, etc.). Permissions
      // still live in-memory for the lifetime of this host.
    }
  }

  /**
   * @param {string} extensionId
   */
  async getGrantedPermissions(extensionId) {
    this._ensureLoaded();
    const id = String(extensionId);
    const hadEntry = Object.prototype.hasOwnProperty.call(this._data, id);
    const record = normalizePermissionRecord(this._data[id]);
    if (hadEntry) {
      this._data[id] = record;
    }
    return JSON.parse(JSON.stringify(record));
  }

  async revokePermissions(extensionId, permissions) {
    this._ensureLoaded();
    const id = String(extensionId);
    if (!this._data[id]) return;
    const current = normalizePermissionRecord(this._data[id]);

    if (!Array.isArray(permissions) || permissions.length === 0) {
      delete this._data[id];
      this._save();
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
    if (Object.keys(current).length === 0) {
      delete this._data[id];
    } else {
      this._data[id] = current;
    }
    this._save();
  }

  /**
   * Reset (clear) all stored permissions for a single extension.
   *
   * @param {string} extensionId
   */
  async resetPermissions(extensionId) {
    return this.revokePermissions(extensionId, []);
  }

  async resetAllPermissions() {
    this._ensureLoaded();
    this._data = {};
    this._save();
  }

  /**
   * @param {{ extensionId: string, displayName?: string, declaredPermissions?: string[] }} meta
   * @param {string[]} permissions
   */
  async ensurePermissions({ extensionId, displayName, declaredPermissions }, permissions, context = {}) {
    this._ensureLoaded();
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
        if (!networkUrl) continue;
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
        record.network = { mode: "allowlist", hosts: [...nextHosts].sort() };
        continue;
      }
      record[perm] = true;
    }
    this._data[extensionId] = record;
    this._save();
    return true;
  }
}

export { PermissionError, PermissionManager };
