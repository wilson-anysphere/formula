/**
 * Canonical action identifiers used by the DLP policy engine.
 *
 * These are intentionally stable strings so policy documents can be stored in
 * local storage and the cloud backend without migrations caused by code
 * refactors.
 */
import dlpCore from "./core.js";

export const { DLP_ACTION } = dlpCore;
