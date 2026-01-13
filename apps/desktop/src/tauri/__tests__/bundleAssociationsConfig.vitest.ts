import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";

function loadTauriConfig(): any {
  const tauriConfUrl = new URL("../../../src-tauri/tauri.conf.json", import.meta.url);
  return JSON.parse(readFileSync(tauriConfUrl, "utf8")) as any;
}

function collectDeepLinkSchemes(config: any): string[] {
  const deepLink = config?.plugins?.["deep-link"];
  const desktop = deepLink?.desktop;

  const schemes: string[] = [];
  const addSchemesFromProtocol = (protocol: any) => {
    const raw = protocol?.schemes;
    if (typeof raw === "string") {
      const v = raw.trim().toLowerCase();
      if (v) schemes.push(v);
      return;
    }
    if (Array.isArray(raw)) {
      for (const item of raw) {
        if (typeof item !== "string") continue;
        const v = item.trim().toLowerCase();
        if (v) schemes.push(v);
      }
    }
  };

  if (Array.isArray(desktop)) {
    for (const protocol of desktop) addSchemesFromProtocol(protocol);
  } else if (desktop != null) {
    addSchemesFromProtocol(desktop);
  }

  return Array.from(new Set(schemes)).sort();
}

describe("tauri.conf.json bundle association guardrails", () => {
  it("does not contain duplicate file association extensions", () => {
    const config = loadTauriConfig();
    const associations = config?.bundle?.fileAssociations;
    expect(Array.isArray(associations), "expected bundle.fileAssociations to be an array").toBe(true);

    const seen = new Set<string>();
    const duplicates: string[] = [];

    for (const assoc of associations as any[]) {
      const exts = Array.isArray(assoc?.ext) ? assoc.ext : [];
      for (const ext of exts) {
        if (typeof ext !== "string") continue;
        const normalized = ext.trim().replace(/^\./, "").toLowerCase();
        if (!normalized) continue;
        if (seen.has(normalized)) duplicates.push(normalized);
        seen.add(normalized);
      }
    }

    expect(duplicates, `duplicate extensions found in bundle.fileAssociations: ${duplicates.join(", ")}`).toEqual([]);
  });

  it("declares MIME types for all file associations (required for Linux .desktop integration)", () => {
    const config = loadTauriConfig();
    const associations = config?.bundle?.fileAssociations;
    expect(Array.isArray(associations), "expected bundle.fileAssociations to be an array").toBe(true);

    const missing: string[] = [];
    for (const assoc of associations as any[]) {
      const exts = Array.isArray(assoc?.ext) ? assoc.ext.filter((v: any) => typeof v === "string") : [];
      const label = exts.length ? exts.join(",") : "<unknown>";

      expect(exts.length, `expected file association entry ${label} to declare at least one extension`).toBeGreaterThan(0);

      const mimeType = assoc?.mimeType;
      if (typeof mimeType !== "string" || mimeType.trim() === "") missing.push(label);
    }

    expect(missing, `missing mimeType for file association entries: ${missing.join("; ")}`).toEqual([]);
  });

  it("configures deep-link desktop schemes so Linux bundles include x-scheme-handler/*", () => {
    const config = loadTauriConfig();
    const schemes = collectDeepLinkSchemes(config);
    expect(schemes.length, "expected plugins[\"deep-link\"].desktop.schemes to be configured").toBeGreaterThan(0);
    expect(schemes).toContain("formula");
  });
});
