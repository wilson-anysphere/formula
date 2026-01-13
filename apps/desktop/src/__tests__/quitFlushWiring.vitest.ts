import { describe, expect, it } from "vitest";

import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const SRC_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

function findIndex(content: string, re: RegExp, fromIndex: number = 0): number {
  const slice = content.slice(fromIndex);
  const match = slice.match(re);
  return match ? fromIndex + (match.index ?? 0) : -1;
}

describe("desktop quit wiring", () => {
  it("flushes collab local persistence before hard-exit", async () => {
    const absMainPath = path.join(SRC_ROOT, "main.ts");
    const content = await readFile(absMainPath, "utf8");

    // Guardrail: the main desktop entrypoint must include the helper.
    expect(content).toContain("flushCollabLocalPersistenceBestEffort");

    // Ensure the `registerAppQuitHandlers` quit/restart paths flush before quitting.
    const quitAppIdx = findIndex(content, /\bquitApp\s*:\s*async\s*\(\s*\)\s*=>\s*\{/, 0);
    const restartAppIdx = findIndex(content, /\brestartApp\s*:\s*async\s*\(\s*\)\s*=>\s*\{/, Math.max(0, quitAppIdx));
    expect(quitAppIdx).toBeGreaterThanOrEqual(0);
    expect(restartAppIdx).toBeGreaterThan(quitAppIdx);

    const quitAppBody = content.slice(quitAppIdx, restartAppIdx);
    const quitAppFlushIdx = quitAppBody.indexOf("flushCollabLocalPersistenceBestEffort");
    const quitAppBinderIdleIdx = quitAppBody.indexOf("whenCollabBinderIdle");
    const quitAppQuitIdx = findIndex(quitAppBody, /\binvoke\s*\(\s*["']quit_app["']/, 0);
    expect(quitAppFlushIdx).toBeGreaterThanOrEqual(0);
    expect(quitAppBinderIdleIdx).toBeGreaterThanOrEqual(0);
    expect(quitAppQuitIdx).toBeGreaterThan(quitAppFlushIdx);

    const restartMarker = "// OAuth PKCE redirect capture:";
    const restartEndIdx = content.indexOf(restartMarker, restartAppIdx);
    expect(restartEndIdx).toBeGreaterThan(restartAppIdx);

    const restartAppBody = content.slice(restartAppIdx, restartEndIdx);
    const restartFlushIdx = restartAppBody.indexOf("flushCollabLocalPersistenceBestEffort");
    const restartBinderIdleIdx = restartAppBody.indexOf("whenCollabBinderIdle");
    const restartInvokeIdx = findIndex(restartAppBody, /\binvoke\s*\(\s*["']restart_app["']/, 0);
    expect(restartFlushIdx).toBeGreaterThanOrEqual(0);
    expect(restartBinderIdleIdx).toBeGreaterThanOrEqual(0);
    expect(restartInvokeIdx).toBeGreaterThan(restartFlushIdx);

    // Ensure the ribbon/native close flow (`handleCloseRequest({ quit: true })`) flushes too.
    const handleCloseStart = findIndex(content, /\basync function handleCloseRequest\s*\(/, 0);
    const handleCloseEnd = content.indexOf("handleCloseRequestForRibbon = handleCloseRequest;", Math.max(0, handleCloseStart));
    expect(handleCloseStart).toBeGreaterThanOrEqual(0);
    expect(handleCloseEnd).toBeGreaterThan(handleCloseStart);

    const handleCloseBody = content.slice(handleCloseStart, handleCloseEnd);
    const handleFlushIdx = handleCloseBody.indexOf("flushCollabLocalPersistenceBestEffort");
    const handleBinderIdleIdx = handleCloseBody.indexOf("whenCollabBinderIdle");
    const handleQuitIdx = findIndex(handleCloseBody, /\binvoke\s*\(\s*["']quit_app["']/, 0);
    expect(handleFlushIdx).toBeGreaterThanOrEqual(0);
    expect(handleBinderIdleIdx).toBeGreaterThanOrEqual(0);
    expect(handleQuitIdx).toBeGreaterThan(handleFlushIdx);
  });
});
