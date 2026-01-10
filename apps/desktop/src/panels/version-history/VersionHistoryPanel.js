/**
 * This file is a minimal placeholder for the Version History UI panel.
 *
 * The real app UI (React) is out of scope for this repo, but we keep the module
 * boundary so the integration points are explicit.
 */

import { t } from "../../i18n/index.js";

/**
 * @typedef {import("../../versioning/index.js").VersionManager} VersionManager
 */

/**
 * @param {import("../../versioning/index.js").VersionRecord[]} versions
 */
export function buildVersionHistoryItems(versions) {
  return versions.map((v) => ({
    id: v.id,
    kind: v.kind,
    timestampMs: v.timestampMs,
    title:
      v.kind === "checkpoint"
        ? v.checkpointName ?? t("versionHistory.checkpoint")
        : v.description ??
          (v.kind === "snapshot"
            ? t("versionHistory.autoSave")
            : v.kind === "restore"
              ? t("versionHistory.restore")
              : v.kind),
    locked: v.kind === "checkpoint" ? Boolean(v.checkpointLocked) : false,
  }));
}
