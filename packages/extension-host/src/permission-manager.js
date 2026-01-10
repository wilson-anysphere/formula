const fs = require("node:fs/promises");
const path = require("node:path");

class PermissionError extends Error {
  constructor(message) {
    super(message);
    this.name = "PermissionError";
  }
}

class PermissionManager {
  constructor({ storagePath, prompt }) {
    if (!storagePath) throw new Error("PermissionManager requires storagePath");
    this._storagePath = storagePath;
    this._prompt = typeof prompt === "function" ? prompt : async () => false;
    this._loaded = false;
    this._data = {};
  }

  async _ensureLoaded() {
    if (this._loaded) return;
    try {
      const raw = await fs.readFile(this._storagePath, "utf8");
      this._data = JSON.parse(raw);
    } catch {
      this._data = {};
    }
    this._loaded = true;
  }

  async _save() {
    await fs.mkdir(path.dirname(this._storagePath), { recursive: true });
    await fs.writeFile(this._storagePath, JSON.stringify(this._data, null, 2), "utf8");
  }

  async getGrantedPermissions(extensionId) {
    await this._ensureLoaded();
    const list = Array.isArray(this._data[extensionId]) ? this._data[extensionId] : [];
    return new Set(list);
  }

  async ensurePermissions({ extensionId, displayName, declaredPermissions }, permissions) {
    await this._ensureLoaded();
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
    await this._save();

    return true;
  }
}

module.exports = {
  PermissionError,
  PermissionManager
};

