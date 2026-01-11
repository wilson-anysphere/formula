import fs from "node:fs/promises";
import path from "node:path";

import extensionHostPkg from "../../../../packages/extension-host/src/index.js";

const { ExtensionHost } = extensionHostPkg;

async function readJsonIfExists(filePath, fallback) {
  try {
    const raw = await fs.readFile(filePath, "utf8");
    return JSON.parse(raw);
  } catch (error) {
    if (error && (error.code === "ENOENT" || error.code === "ENOTDIR")) return fallback;
    throw error;
  }
}

/**
 * Desktop runtime shim that keeps a single `ExtensionHost` instance alive and
 * loads marketplace-installed extensions into it.
 *
 * - Extensions are installed/extracted by `apps/desktop/src/marketplace/extensionManager.js`
 * - The installed list is stored in `ExtensionManager.statePath`
 * - This manager reads that state file at startup and calls `ExtensionHost.loadExtension(...)`
 *
 * This module is intentionally Node-only (uses `worker_threads` via `@formula/extension-host`).
 */
export class ExtensionHostManager {
  /**
   * @param {{
   *   extensionsDir: string,
   *   statePath: string,
   *   engineVersion?: string,
   *   permissionPrompt?: (req: { extensionId: string, displayName: string, permissions: string[] }) => Promise<boolean>,
   *   permissionsStoragePath?: string,
   *   extensionStoragePath?: string,
   *   auditDbPath?: string | null,
   *   activationTimeoutMs?: number,
   *   commandTimeoutMs?: number,
   *   customFunctionTimeoutMs?: number,
   *   memoryMb?: number,
   *   spreadsheet?: any,
   * }} params
   */
  constructor({
    extensionsDir,
    statePath,
    engineVersion = "1.0.0",
    permissionPrompt,
    permissionsStoragePath,
    extensionStoragePath,
    auditDbPath = null,
    activationTimeoutMs,
    commandTimeoutMs,
    customFunctionTimeoutMs,
    memoryMb,
    spreadsheet,
  }) {
    if (!extensionsDir) throw new Error("extensionsDir is required");
    if (!statePath) throw new Error("statePath is required");

    this.extensionsDir = extensionsDir;
    this.statePath = statePath;

    const baseDir = path.dirname(path.resolve(statePath));
    const resolvedPermissionsStoragePath = permissionsStoragePath ?? path.join(baseDir, "permissions.json");
    const resolvedExtensionStoragePath = extensionStoragePath ?? path.join(baseDir, "storage.json");

    this._host = new ExtensionHost({
      engineVersion,
      permissionPrompt,
      permissionsStoragePath: resolvedPermissionsStoragePath,
      extensionStoragePath: resolvedExtensionStoragePath,
      auditDbPath,
      activationTimeoutMs,
      commandTimeoutMs,
      customFunctionTimeoutMs,
      memoryMb,
      spreadsheet,
    });

    this._started = false;
  }

  get spreadsheet() {
    return this._host.spreadsheet;
  }

  async _loadInstalledState() {
    const state = await readJsonIfExists(this.statePath, { installed: {} });
    return state && typeof state === "object" ? state : { installed: {} };
  }

  /**
   * Loads all installed extensions (per `ExtensionManager.statePath`) into the runtime.
   * Safe to call multiple times (already-loaded extensions are skipped).
   */
  async startup() {
    const state = await this._loadInstalledState();
    const installedIds = Object.keys(state.installed ?? {}).sort();
    const loaded = new Set(this._host.listExtensions().map((ext) => ext.id));

    for (const extensionId of installedIds) {
      if (loaded.has(extensionId)) continue;
      const extensionPath = path.join(this.extensionsDir, extensionId);
      await this._host.loadExtension(extensionPath);
    }

    if (!this._started) {
      await this._host.startup();
      this._started = true;
    }
  }

  async dispose() {
    await this._host.dispose();
    this._started = false;
  }

  async executeCommand(commandId, ...args) {
    return this._host.executeCommand(String(commandId), ...args);
  }

  async invokeCustomFunction(name, ...args) {
    return this._host.invokeCustomFunction(String(name), ...args);
  }

  listContributions() {
    return {
      commands: this._host.getContributedCommands(),
      keybindings: this._host.getContributedKeybindings(),
      menus: this._host.getContributedMenus(),
      panels: this._host.getContributedPanels(),
      customFunctions: this._host.getContributedCustomFunctions(),
      dataConnectors: this._host.getContributedDataConnectors(),
    };
  }

  /**
   * Reload an installed extension:
   * - If already loaded: unload (removes contributions + terminates worker)
   * - Load fresh from disk (re-reads manifest, registers new contributions, spawns new worker)
   */
  async reloadExtension(extensionId) {
    const id = String(extensionId);
    const loaded = this._host.listExtensions().some((ext) => ext.id === id);
    if (loaded) {
      await this._host.unloadExtension(id);
    }

    const extensionPath = path.join(this.extensionsDir, id);
    await this._host.loadExtension(extensionPath);
    if (this._started) {
      // Ensure extensions that activate on `onStartupFinished` get a chance to run
      // when they are installed/updated after the runtime has already started.
      await this._host.startup();
    }
  }

  /**
   * Unload an extension from the runtime (does not delete files on disk).
   */
  async unloadExtension(extensionId) {
    const id = String(extensionId);
    const loaded = this._host.listExtensions().some((ext) => ext.id === id);
    if (!loaded) return;
    await this._host.unloadExtension(id);
  }
}
