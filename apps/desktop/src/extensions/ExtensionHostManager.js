import * as fsSync from "node:fs";
import fs from "node:fs/promises";
import path from "node:path";

import extensionHostPkg from "../../../../packages/extension-host/src/index.js";
import extensionPackagePkg from "../../../../shared/extension-package/index.js";

const { ExtensionHost } = extensionHostPkg;
const { verifyExtractedExtensionDir } = extensionPackagePkg;

async function readJsonIfExists(filePath, fallback) {
  try {
    const raw = await fs.readFile(filePath, "utf8");
    return JSON.parse(raw);
  } catch (error) {
    if (error && (error.code === "ENOENT" || error.code === "ENOTDIR")) return fallback;
    if (error instanceof SyntaxError) return fallback;
    throw error;
  }
}

async function atomicWriteJson(filePath, data) {
  await fs.mkdir(path.dirname(filePath), { recursive: true });
  const tmp = `${filePath}.${process.pid}.${Date.now()}.tmp`;
  await fs.writeFile(tmp, JSON.stringify(data, null, 2));
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
 *   dataConnectorTimeoutMs?: number,
 *   memoryMb?: number,
 *   spreadsheet?: any,
 *   extensionManager?: any,
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
    dataConnectorTimeoutMs,
    memoryMb,
    spreadsheet,
    extensionManager = null,
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
      dataConnectorTimeoutMs,
      memoryMb,
      spreadsheet,
    });

    this._extensionManagerSubscription = null;
    this._started = false;

    if (extensionManager) {
      this.bindToExtensionManager(extensionManager);
    }

    // Serialize all operations that touch the underlying ExtensionHost. Many host APIs mutate
    // shared runtime state (commands, panels, registrations) and are not safe to call concurrently
    // with unload/reload/sync.
    this._hostQueue = Promise.resolve();

    // `syncInstalledExtensions()` can be triggered from multiple sources (e.g. installer events,
    // explicit UI refresh calls). Serialize sync runs and coalesce concurrent requests so we don't
    // unload/reload the same extension simultaneously.
    this._syncRunner = null;
    this._syncRequested = false;

    this._stateWatcher = null;
    this._stateWatchTimer = null;
  }

  _runHostOperation(fn) {
    const run = () => Promise.resolve().then(fn);
    const task = this._hostQueue.then(run, run);
    // Keep the queue alive even if a task fails.
    this._hostQueue = task.catch(() => {});
    return task;
  }

  get spreadsheet() {
    return this._host.spreadsheet;
  }

  async _loadInstalledState() {
    const state = await readJsonIfExists(this.statePath, { installed: {} });
    if (!state || typeof state !== "object") return { installed: {} };
    if (!state.installed || typeof state.installed !== "object") state.installed = {};
    return state;
  }

  async _saveInstalledState(state) {
    await atomicWriteJson(this.statePath, state);
  }

  async _markCorrupted(state, extensionId, reason) {
    const record = state.installed?.[extensionId];
    if (!record || typeof record !== "object") return;
    record.corrupted = true;
    record.corruptedAt = new Date().toISOString();
    record.corruptedReason = reason;
    await this._saveInstalledState(state);
  }

  async _verifyInstalledExtension(state, extensionId) {
    const record = state.installed?.[extensionId];
    if (!record || typeof record !== "object") {
      return { ok: false, reason: "Missing installed extension metadata" };
    }

    if (record.corrupted) {
      return { ok: false, reason: record.corruptedReason || "Extension is quarantined" };
    }

    if (!Array.isArray(record.files) || record.files.length === 0) {
      const reason =
        "Missing integrity metadata (installed with an older version). Repair (reinstall) is required.";
      await this._markCorrupted(state, extensionId, reason);
      return { ok: false, reason };
    }

    const extensionPath = path.join(this.extensionsDir, extensionId);
    let result;
    try {
      result = await verifyExtractedExtensionDir(extensionPath, record.files, {
        ignoreExtraPaths: [".DS_Store", "Thumbs.db", "desktop.ini"],
      });
    } catch (error) {
      result = { ok: false, reason: error?.message ?? String(error) };
    }
    if (!result.ok) {
      await this._markCorrupted(state, extensionId, result.reason || "Extension integrity check failed");
    }
    return result;
  }

  /**
   * Loads all installed extensions (per `ExtensionManager.statePath`) into the runtime.
   * Safe to call multiple times (already-loaded extensions are skipped).
   */
  async startup() {
    return this._runHostOperation(async () => {
      const state = await this._loadInstalledState();
      const installedIds = Object.keys(state.installed ?? {}).sort();
      const loaded = new Set(this._host.listExtensions().map((ext) => ext.id));

      for (const extensionId of installedIds) {
        if (loaded.has(extensionId)) continue;

        const verification = await this._verifyInstalledExtension(state, extensionId);
        if (!verification.ok) {
          // Do not prevent the rest of the runtime from starting if one install is corrupted.
          // The extension is marked as corrupted in state and can be repaired by reinstalling.
          // eslint-disable-next-line no-console
          console.warn(
            `Skipping extension ${extensionId}: integrity check failed: ${verification.reason || "unknown reason"}`
          );
          continue;
        }

        const extensionPath = path.join(this.extensionsDir, extensionId);
        await this._host.loadExtension(extensionPath);
      }

      if (!this._started) {
        await this._host.startup();
        this._started = true;
      }
    });
  }

  async dispose() {
    if (this._extensionManagerSubscription) {
      try {
        this._extensionManagerSubscription.dispose();
      } catch {
        // ignore
      }
      this._extensionManagerSubscription = null;
    }
    this.unwatchStateFile();
    await this._runHostOperation(async () => {
      await this._host.dispose();
      this._started = false;
    });
  }

  async watchStateFile({ debounceMs = 200 } = {}) {
    if (this._stateWatcher) return;

    const resolvedStatePath = path.resolve(this.statePath);
    const dir = path.dirname(resolvedStatePath);
    const base = path.basename(resolvedStatePath);

    await fs.mkdir(dir, { recursive: true });

    const debounce = Number.isFinite(debounceMs) ? Math.max(0, debounceMs) : 200;

    const scheduleSync = () => {
      if (this._stateWatchTimer) clearTimeout(this._stateWatchTimer);
      this._stateWatchTimer = setTimeout(() => {
        this._stateWatchTimer = null;
        void this.syncInstalledExtensions().catch((error) => {
          // eslint-disable-next-line no-console
          console.warn(`Failed to sync installed extensions: ${String(error?.message ?? error)}`);
        });
      }, debounce);
    };

    try {
      this._stateWatcher = fsSync.watch(dir, { persistent: false }, (_eventType, filename) => {
        if (!filename) {
          scheduleSync();
          return;
        }
        const changed = filename instanceof Buffer ? filename.toString("utf8") : String(filename);
        if (!changed) {
          scheduleSync();
          return;
        }
        // Some platforms return a full path (or other non-basename strings) here. Be permissive
        // and sync whenever the reported filename appears to refer to the state file.
        if (changed === base || changed.endsWith(`/${base}`) || changed.endsWith(`\\${base}`) || changed.endsWith(base)) {
          scheduleSync();
        }
      });
    } catch (error) {
      throw new Error(`Failed to watch extension state file: ${error?.message ?? String(error)}`);
    }

    this._stateWatcher.on("error", (error) => {
      // eslint-disable-next-line no-console
      console.warn(`Extension state file watcher error: ${String(error?.message ?? error)}`);
    });
  }

  unwatchStateFile() {
    if (this._stateWatchTimer) {
      clearTimeout(this._stateWatchTimer);
      this._stateWatchTimer = null;
    }
    if (this._stateWatcher) {
      try {
        this._stateWatcher.close();
      } catch {
        // ignore
      }
      this._stateWatcher = null;
    }
  }

  bindToExtensionManager(extensionManager) {
    if (!extensionManager || typeof extensionManager.onDidChange !== "function") {
      throw new Error("bindToExtensionManager requires an ExtensionManager with onDidChange()");
    }

    if (this._extensionManagerSubscription) {
      try {
        this._extensionManagerSubscription.dispose();
      } catch {
        // ignore
      }
      this._extensionManagerSubscription = null;
    }

    this._extensionManagerSubscription = extensionManager.onDidChange(() => {
      void this.syncInstalledExtensions().catch((error) => {
        // eslint-disable-next-line no-console
        console.warn(`Failed to sync installed extensions: ${String(error?.message ?? error)}`);
      });
    });
  }

  async _syncInstalledExtensionsOnce() {
    const state = await this._loadInstalledState();
    const installed = state.installed ?? {};
    const installedIds = new Set(Object.keys(installed));

    // Always re-verify integrity for installed extensions when syncing. This allows the
    // runtime to detect on-disk tampering that happens after the initial install/startup
    // and quarantine the extension (mark corrupted + unload) before it executes again.
    for (const id of installedIds) {
      if (installed[id]?.corrupted) continue;
      try {
        const verification = await this._verifyInstalledExtension(state, id);
        if (!verification.ok) {
          // eslint-disable-next-line no-console
          console.warn(
            `Quarantining extension ${id}: integrity check failed: ${verification.reason || "unknown reason"}`
          );
        }
      } catch (error) {
        const reason = error?.message ?? String(error);
        await this._markCorrupted(state, id, reason);
        // eslint-disable-next-line no-console
        console.warn(`Quarantining extension ${id}: integrity check failed: ${reason}`);
      }
    }

    const loadedExtensions = this._host.listExtensions();
    const loadedById = new Map(loadedExtensions.map((ext) => [ext.id, ext]));

    const toUnload = [];
    const toReload = [];
    const toLoad = [];

    for (const ext of loadedExtensions) {
      if (!installedIds.has(ext.id)) {
        toUnload.push(ext.id);
        continue;
      }

      if (installed[ext.id]?.corrupted) {
        toUnload.push(ext.id);
        continue;
      }

      const expectedVersion = installed[ext.id]?.version;
      const loadedVersion = ext.manifest?.version;
      if (expectedVersion && loadedVersion && expectedVersion !== loadedVersion) {
        toReload.push(ext.id);
      }
    }

    for (const id of installedIds) {
      if (!loadedById.has(id)) {
        toLoad.push(id);
      }
    }

    for (const id of toUnload) {
      await this._host.unloadExtension(id);
    }

    for (const id of toReload) {
      if (installed[id]?.corrupted) continue;
      try {
        await this._reloadExtensionUnsafe(id);
      } catch (error) {
        // eslint-disable-next-line no-console
        console.warn(`Failed to reload extension ${id}: ${String(error?.message ?? error)}`);
      }
    }

    for (const id of toLoad) {
      // Reuse reload semantics so we always run integrity checks before loading
      // new installed extensions into the runtime.
      if (installed[id]?.corrupted) continue;
      try {
        await this._reloadExtensionUnsafe(id);
      } catch (error) {
        // eslint-disable-next-line no-console
        console.warn(`Failed to load extension ${id}: ${String(error?.message ?? error)}`);
      }
    }
  }

  async syncInstalledExtensions() {
    this._syncRequested = true;

    if (this._syncRunner) {
      return this._syncRunner;
    }

    this._syncRunner = this._runHostOperation(async () => {
      try {
        while (this._syncRequested) {
          this._syncRequested = false;
          await this._syncInstalledExtensionsOnce();
        }
      } finally {
        this._syncRequested = false;
        this._syncRunner = null;
      }
    });

    return this._syncRunner;
  }

  async executeCommand(commandId, ...args) {
    return this._runHostOperation(() => this._host.executeCommand(String(commandId), ...args));
  }

  async invokeCustomFunction(name, ...args) {
    return this._runHostOperation(() => this._host.invokeCustomFunction(String(name), ...args));
  }

  async invokeDataConnector(connectorId, method, ...args) {
    return this._runHostOperation(() =>
      this._host.invokeDataConnector(String(connectorId), String(method), ...args)
    );
  }

  async activateView(viewId) {
    return this._runHostOperation(() => this._host.activateView(String(viewId)));
  }

  getPanel(panelId) {
    return this._host.getPanel(String(panelId));
  }

  getPanelOutgoingMessages(panelId) {
    return this._host.getPanelOutgoingMessages(String(panelId));
  }

  dispatchPanelMessage(panelId, message) {
    return this._host.dispatchPanelMessage(String(panelId), message);
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
  async _reloadExtensionUnsafe(extensionId) {
    const id = String(extensionId);
    const loaded = this._host.listExtensions().some((ext) => ext.id === id);
    if (loaded) {
      await this._host.unloadExtension(id);
    }

    const state = await this._loadInstalledState();
    const verification = await this._verifyInstalledExtension(state, id);
    if (!verification.ok) {
      const reason = verification.reason || "unknown reason";
      throw new Error(
        `Extension integrity check failed for ${id}: ${reason}. Repair (reinstall) the extension to continue.`
      );
    }

    const extensionPath = path.join(this.extensionsDir, id);
    await this._host.loadExtension(extensionPath);
    if (this._started) {
      await this._host.startupExtension?.(id);
    }
  }

  async reloadExtension(extensionId) {
    return this._runHostOperation(() => this._reloadExtensionUnsafe(extensionId));
  }

  /**
   * Unload an extension from the runtime (does not delete files on disk).
   */
  async _unloadExtensionUnsafe(extensionId) {
    const id = String(extensionId);
    const loaded = this._host.listExtensions().some((ext) => ext.id === id);
    if (!loaded) return;
    await this._host.unloadExtension(id);
  }

  async unloadExtension(extensionId) {
    return this._runHostOperation(() => this._unloadExtensionUnsafe(extensionId));
  }
}
