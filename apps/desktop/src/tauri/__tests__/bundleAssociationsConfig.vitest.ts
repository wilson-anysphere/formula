import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";

import { stripHtmlComments } from "../../__tests__/sourceTextUtils";

function loadTauriConfig(): any {
  const tauriConfUrl = new URL("../../../src-tauri/tauri.conf.json", import.meta.url);
  return JSON.parse(readFileSync(tauriConfUrl, "utf8")) as any;
}

function hasExtension(config: any, ext: string): boolean {
  const associations = Array.isArray(config?.bundle?.fileAssociations) ? config.bundle.fileAssociations : [];
  const needle = ext.trim().replace(/^\./, "").toLowerCase();
  if (!needle) return false;
  return associations.some((assoc: any) =>
    Array.isArray(assoc?.ext)
      ? assoc.ext.some((v: any) => typeof v === "string" && v.trim().replace(/^\./, "").toLowerCase() === needle)
      : false,
  );
}

function collectDeepLinkSchemes(config: any): string[] {
  const deepLink = config?.plugins?.["deep-link"];
  const desktop = deepLink?.desktop;

  const schemes: string[] = [];
  const normalizeScheme = (raw: string): string => raw.trim().toLowerCase().replace(/[:/]+$/, "");
  const addSchemesFromProtocol = (protocol: any) => {
    const raw = protocol?.schemes;
    if (typeof raw === "string") {
      const v = normalizeScheme(raw);
      if (v) schemes.push(v);
      return;
    }
    if (Array.isArray(raw)) {
      for (const item of raw) {
        if (typeof item !== "string") continue;
        const v = normalizeScheme(item);
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
    const invalid = schemes.filter((scheme) => /[:/]/.test(scheme));
    expect(invalid, `invalid deep-link scheme(s) (must be scheme names, no ':' or '/'): ${invalid.join(", ")}`).toEqual(
      [],
    );
  });

  it("ships a shared-mime-info Parquet definition when .parquet file association is configured", () => {
    const config = loadTauriConfig();
    if (!hasExtension(config, "parquet")) {
      // If Parquet support is removed, this check is irrelevant.
      return;
    }

    const linux = config?.bundle?.linux;
    expect(linux && typeof linux === "object", "expected bundle.linux to be configured").toBeTruthy();

    const identifier = typeof config?.identifier === "string" ? config.identifier.trim() : "";
    expect(identifier, "expected tauri.conf.json identifier to be configured").toBeTruthy();

    const expectedDest = `usr/share/mime/packages/${identifier}.xml`;
    const expectedSrc = `mime/${identifier}.xml`;

    for (const target of ["deb", "rpm", "appimage"] as const) {
      const files = linux?.[target]?.files;
      expect(files && typeof files === "object", `expected bundle.linux.${target}.files to be configured`).toBeTruthy();
      expect(files[expectedDest]).toBe(expectedSrc);
    }

    // Ensure `update-mime-database` triggers exist at install time.
    expect(Array.isArray(linux?.deb?.depends), "expected bundle.linux.deb.depends to be an array").toBe(true);
    expect(linux.deb.depends).toContain("shared-mime-info");
    expect(Array.isArray(linux?.rpm?.depends), "expected bundle.linux.rpm.depends to be an array").toBe(true);
    expect(linux.rpm.depends).toContain("shared-mime-info");
  });

  it("includes a Parquet glob in the shared-mime-info definition file", () => {
    const config = loadTauriConfig();
    if (!hasExtension(config, "parquet")) {
      return;
    }

    const identifier = typeof config?.identifier === "string" ? config.identifier.trim() : "";
    expect(identifier, "expected tauri.conf.json identifier to be configured").toBeTruthy();

    const xmlUrl = new URL(`../../../src-tauri/mime/${identifier}.xml`, import.meta.url);
    const xml = stripHtmlComments(readFileSync(xmlUrl, "utf8"));
    expect(xml).toContain('mime-type type="application/vnd.apache.parquet"');
    expect(xml).toContain('glob pattern="*.parquet"');
  });
});
