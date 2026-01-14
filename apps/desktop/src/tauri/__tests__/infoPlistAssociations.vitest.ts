import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";

import { stripHtmlComments } from "../../__tests__/sourceTextUtils";

function loadTauriConfig(): any {
  const tauriConfUrl = new URL("../../../src-tauri/tauri.conf.json", import.meta.url);
  return JSON.parse(readFileSync(tauriConfUrl, "utf8")) as any;
}

function loadInfoPlistText(): string {
  const infoPlistUrl = new URL("../../../src-tauri/Info.plist", import.meta.url);
  return stripHtmlComments(readFileSync(infoPlistUrl, "utf8"));
}

function normalizeExt(ext: string): string {
  return ext.trim().replace(/^\./, "").toLowerCase();
}

function collectFileAssociationExtensions(config: any): string[] {
  const associations = Array.isArray(config?.bundle?.fileAssociations) ? config.bundle.fileAssociations : [];
  const exts: string[] = [];
  for (const assoc of associations) {
    const raw = Array.isArray(assoc?.ext) ? assoc.ext : [];
    for (const ext of raw) {
      if (typeof ext !== "string") continue;
      const normalized = normalizeExt(ext);
      if (normalized) exts.push(normalized);
    }
  }
  return Array.from(new Set(exts)).sort();
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

describe("Info.plist desktop integration guardrails", () => {
  it("includes CFBundleDocumentTypes entries for the file associations configured in tauri.conf.json", () => {
    const config = loadTauriConfig();
    const infoPlist = loadInfoPlistText();

    const expectedExts = collectFileAssociationExtensions(config);
    expect(expectedExts.length).toBeGreaterThan(0);

    // This is intentionally a lightweight (string-based) check. The release workflow validates the
    // *built* bundle's binary plist via `plutil` + `plistlib` to catch packaging regressions.
    for (const ext of expectedExts) {
      expect(
        infoPlist.includes(`<string>${ext}</string>`),
        `apps/desktop/src-tauri/Info.plist missing <string>${ext}</string> (keep CFBundleDocumentTypes in sync with bundle.fileAssociations)`,
      ).toBe(true);
    }
  });

  it("includes CFBundleURLSchemes entries for the deep-link schemes configured in tauri.conf.json", () => {
    const config = loadTauriConfig();
    const infoPlist = loadInfoPlistText();

    const expectedSchemes = collectDeepLinkSchemes(config);
    expect(expectedSchemes.length).toBeGreaterThan(0);

    for (const scheme of expectedSchemes) {
      expect(
        infoPlist.includes(`<string>${scheme}</string>`),
        `apps/desktop/src-tauri/Info.plist missing <string>${scheme}</string> (keep CFBundleURLSchemes in sync with plugins[\"deep-link\"].desktop.schemes)`,
      ).toBe(true);
    }
  });
});
