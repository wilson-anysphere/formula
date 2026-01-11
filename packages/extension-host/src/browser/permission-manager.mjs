class PermissionError extends Error {
  constructor(message) {
    super(message);
    this.name = "PermissionError";
  }
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
    this._data = {};
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
      const parsed = raw ? JSON.parse(raw) : {};
      this._data = parsed && typeof parsed === "object" ? parsed : {};
    } catch {
      this._data = {};
    }
  }

  _save() {
    if (!this._storage) return;
    try {
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
    const list = Array.isArray(this._data[extensionId]) ? this._data[extensionId] : [];
    return new Set(list);
  }

  /**
   * @param {{ extensionId: string, displayName?: string, declaredPermissions?: string[] }} meta
   * @param {string[]} permissions
   */
  async ensurePermissions({ extensionId, displayName, declaredPermissions }, permissions) {
    this._ensureLoaded();
    const requested = Array.isArray(permissions) ? permissions : [];
    if (requested.length === 0) return true;

    const declared = new Set(declaredPermissions ?? []);
    for (const perm of requested) {
      if (!declared.has(perm)) {
        throw new PermissionError(`Permission not declared in manifest: ${perm}`);
      }
    }

    const granted = await this.getGrantedPermissions(extensionId);
    const needed = requested.filter((p) => !granted.has(p));
    if (needed.length === 0) return true;

    const accepted = await this._prompt({
      extensionId,
      displayName,
      permissions: needed
    });

    if (!accepted) {
      throw new PermissionError(`Permission denied: ${needed.join(", ")}`);
    }

    for (const perm of needed) granted.add(perm);
    this._data[extensionId] = [...granted].sort();
    this._save();
    return true;
  }
}

export { PermissionError, PermissionManager };
