import { describe, expect, it } from "vitest";

import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const SRC_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

describe("desktop quit wiring", () => {
  it("flushes collab local persistence before hard-exit", async () => {
    const absMainPath = path.join(SRC_ROOT, "main.ts");
    const content = await readFile(absMainPath, "utf8");

    // Guardrail: the main desktop entrypoint must include the helper.
    expect(content).toContain("flushCollabLocalPersistenceBestEffort");

    // Ensure the `registerAppQuitHandlers` quit/restart paths flush before quitting.
    const quitAppIdx = content.indexOf("quitApp: async () => {");
    const restartAppIdx = content.indexOf("restartApp: async () => {", quitAppIdx);
    expect(quitAppIdx).toBeGreaterThanOrEqual(0);
    expect(restartAppIdx).toBeGreaterThan(quitAppIdx);

    const quitAppBody = content.slice(quitAppIdx, restartAppIdx);
    const quitAppFlushIdx = quitAppBody.indexOf("flushCollabLocalPersistenceBestEffort");
    const quitAppQuitIdx = quitAppBody.indexOf('invoke("quit_app")');
    expect(quitAppFlushIdx).toBeGreaterThanOrEqual(0);
    expect(quitAppQuitIdx).toBeGreaterThan(quitAppFlushIdx);

    const restartMarker = "// OAuth PKCE redirect capture:";
    const restartEndIdx = content.indexOf(restartMarker, restartAppIdx);
    expect(restartEndIdx).toBeGreaterThan(restartAppIdx);

    const restartAppBody = content.slice(restartAppIdx, restartEndIdx);
    const restartFlushIdx = restartAppBody.indexOf("flushCollabLocalPersistenceBestEffort");
    const restartInvokeIdx = restartAppBody.indexOf('invoke("restart_app")');
    expect(restartFlushIdx).toBeGreaterThanOrEqual(0);
    expect(restartInvokeIdx).toBeGreaterThan(restartFlushIdx);

    // Ensure the ribbon/native close flow (`handleCloseRequest({ quit: true })`) flushes too.
    const handleCloseStart = content.indexOf("async function handleCloseRequest");
    const handleCloseEnd = content.indexOf("handleCloseRequestForRibbon = handleCloseRequest;", handleCloseStart);
    expect(handleCloseStart).toBeGreaterThanOrEqual(0);
    expect(handleCloseEnd).toBeGreaterThan(handleCloseStart);

    const handleCloseBody = content.slice(handleCloseStart, handleCloseEnd);
    const handleFlushIdx = handleCloseBody.indexOf("flushCollabLocalPersistenceBestEffort");
    const handleQuitIdx = handleCloseBody.indexOf('invoke("quit_app")');
    expect(handleFlushIdx).toBeGreaterThanOrEqual(0);
    expect(handleQuitIdx).toBeGreaterThan(handleFlushIdx);
  });
});

