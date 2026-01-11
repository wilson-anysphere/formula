import { PermissionDeniedError } from "./errors.js";
import {
  createLockedDownGrant,
  principalKey,
  normalizeScopePath,
  checkPermissionGrant
} from "./permissions.js";

function mergeUniqueStrings(into, values) {
  for (const value of values) {
    if (typeof value !== "string") continue;
    if (!into.includes(value)) into.push(value);
  }
}

function mergePermissionGrant(current, update) {
  const merged = {
    filesystem: {
      read: [...current.filesystem.read],
      readwrite: [...current.filesystem.readwrite]
    },
    network: {
      mode: current.network.mode,
      allowlist: [...current.network.allowlist]
    },
    clipboard: current.clipboard,
    notifications: current.notifications,
    automation: current.automation
  };

  if (update?.filesystem) {
    if (Array.isArray(update.filesystem.read)) {
      mergeUniqueStrings(merged.filesystem.read, update.filesystem.read.map(normalizeScopePath));
    }
    if (Array.isArray(update.filesystem.readwrite)) {
      mergeUniqueStrings(
        merged.filesystem.readwrite,
        update.filesystem.readwrite.map(normalizeScopePath)
      );
    }
  }

  if (update?.network) {
    if (update.network.mode === "full") {
      merged.network.mode = "full";
    } else if (update.network.mode === "allowlist") {
      if (merged.network.mode !== "full") merged.network.mode = "allowlist";
      if (Array.isArray(update.network.allowlist)) {
        mergeUniqueStrings(merged.network.allowlist, update.network.allowlist);
      }
    } else if (update.network.mode === "none") {
      // Never escalate "none" into a broader grant.
      // "none" updates are ignored unless the grant is uninitialized.
      if (!current) merged.network.mode = "none";
    } else if (Array.isArray(update.network.allowlist)) {
      if (merged.network.mode !== "full") merged.network.mode = "allowlist";
      mergeUniqueStrings(merged.network.allowlist, update.network.allowlist);
    }
  }

  if (update?.clipboard === true) merged.clipboard = true;
  if (update?.notifications === true) merged.notifications = true;
  if (update?.automation === true) merged.automation = true;

  return merged;
}

function buildGrantForRequest(request) {
  switch (request.kind) {
    case "filesystem": {
      const access = request.access;
      const scope = normalizeScopePath(request.path);
      if (access === "readwrite") {
        return { filesystem: { readwrite: [scope] } };
      }
      return { filesystem: { read: [scope] } };
    }
    case "network": {
      const origin = new URL(request.url).origin;
      return { network: { mode: "allowlist", allowlist: [origin] } };
    }
    case "clipboard":
      return { clipboard: true };
    case "notifications":
      return { notifications: true };
    case "automation":
      return { automation: true };
    default:
      return null;
  }
}

export class PermissionManager {
  /**
   * @param {object} options
   * @param {{ log: (event: any) => string } | null} [options.auditLogger]
   * @param {(prompt: {principal: any, request: any}) => Promise<boolean>} [options.onPrompt]
   */
  constructor({ auditLogger = null, onPrompt = null } = {}) {
    this.auditLogger = auditLogger;
    this.onPrompt = onPrompt;
    this.grants = new Map();
  }

  getGrant(principal) {
    const key = principalKey(principal);
    return this.grants.get(key) ?? createLockedDownGrant();
  }

  /**
   * Returns a JSON-serializable snapshot of a principal's permissions.
   * Intended to be passed into sandboxes (Worker/vm) without sharing mutable state.
   */
  getSnapshot(principal) {
    const grant = this.getGrant(principal);
    return JSON.parse(JSON.stringify(grant));
  }

  grant(principal, update, metadata = {}) {
    const key = principalKey(principal);
    const current = this.grants.get(key) ?? createLockedDownGrant();
    const merged = mergePermissionGrant(current, update);
    this.grants.set(key, merged);

    this.auditLogger?.log({
      eventType: "security.permission.granted",
      actor: principal,
      success: true,
      metadata: { update, ...metadata }
    });

    return merged;
  }

  check(principal, request) {
    const grant = this.getGrant(principal);
    return checkPermissionGrant(grant, request);
  }

  /**
   * @param {any} principal
   * @param {any} request
   * @param {object} [options]
   * @param {boolean} [options.promptIfDenied]
   */
  async ensure(principal, request, { promptIfDenied = false } = {}) {
    const decision = this.check(principal, request);
    this.auditLogger?.log({
      eventType: "security.permission.checked",
      actor: principal,
      success: decision.allowed,
      metadata: {
        request,
        ...(decision.allowed ? {} : { reason: decision.reason })
      }
    });
    if (decision.allowed) return;

    if (promptIfDenied && this.onPrompt) {
      this.auditLogger?.log({
        eventType: "security.permission.prompted",
        actor: principal,
        success: true,
        metadata: { request }
      });

      const granted = await this.onPrompt({ principal, request });
      if (granted) {
        const update = buildGrantForRequest(request);
        if (update) this.grant(principal, update, { reason: "prompt" });

        const after = this.check(principal, request);
        if (after.allowed) return;
      }
    }

    this.auditLogger?.log({
      eventType: "security.permission.denied",
      actor: principal,
      success: false,
      metadata: { request, reason: decision.reason }
    });

    throw new PermissionDeniedError({ principal, request, reason: decision.reason });
  }
}
