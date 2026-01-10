import { createHash } from "node:crypto";

/**
 * @param {string} text
 * @returns {string} lowercase hex sha256 digest
 */
export function sha256Hex(text) {
  return createHash("sha256").update(text).digest("hex");
}
