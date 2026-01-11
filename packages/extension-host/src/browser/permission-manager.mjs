class PermissionError extends Error {
  constructor(message) {
    super(message);
    this.name = "PermissionError";
  }
}

class PermissionManager {
  constructor({ prompt } = {}) {
    this._prompt = typeof prompt === "function" ? prompt : async () => false;
    /** @type {Map<string, Set<string>>} */
    this._granted = new Map();
  }

  /**
   * @param {string} extensionId
   */
  async getGrantedPermissions(extensionId) {
    const granted = this._granted.get(extensionId);
    return new Set(granted ? [...granted] : []);
  }

  /**
   * @param {{ extensionId: string, displayName?: string, declaredPermissions?: string[] }} meta
   * @param {string[]} permissions
   */
  async ensurePermissions({ extensionId, displayName, declaredPermissions }, permissions) {
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
    this._granted.set(extensionId, granted);
    return true;
  }
}

export { PermissionError, PermissionManager };

