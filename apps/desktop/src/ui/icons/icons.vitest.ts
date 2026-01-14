import { readdirSync, readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { stripComments } from "../../__tests__/sourceTextUtils";

function listIconSourceFiles(dir: string) {
  return readdirSync(dir)
    .filter((name) => name.endsWith(".tsx"))
    .filter((name) => !["Icon.tsx"].includes(name))
    .sort();
}

describe("ui/icons", () => {
  it("does not introduce hardcoded colors and uses Icon base component", () => {
    const iconsDir = dirname(fileURLToPath(import.meta.url));
    const files = listIconSourceFiles(iconsDir);
    const indexSrc = stripComments(readFileSync(join(iconsDir, "index.ts"), "utf8"));

    for (const file of files) {
      const fullPath = join(iconsDir, file);
      const src = stripComments(readFileSync(fullPath, "utf8"));
      const exportName = file.replace(/\.tsx$/, "");

      // All icon components should go through the shared base component (so sizing,
      // stroke styles, and currentColor defaults are consistent).
      expect(src, `${file} should use <Icon>`).toContain("<Icon");

      // Ensure the icon is re-exported from the barrel so consumers can import
      // consistently from `ui/icons`.
      expect(indexSrc, `index.ts should export ${exportName}`).toContain(`"./${exportName}"`);

      // Prevent hard-coded color values. Icons must inherit `currentColor`.
      expect(src, `${file} should not contain hex colors`).not.toMatch(/#[0-9a-fA-F]{3,8}\b/);
      expect(src, `${file} should not contain rgb()/hsl() colors`).not.toMatch(/\b(?:rgb|hsl)a?\(/);
      expect(src, `${file} should not contain CSS var() colors`).not.toMatch(/\bvar\(--/);

      // If fill or stroke are specified explicitly, they must still respect currentColor.
      for (const match of src.matchAll(/\b(fill|stroke)\s*=\s*"([^"]+)"/g)) {
        const [, attr, value] = match;
        if (attr === "fill") {
          expect(["none", "currentColor"]).toContain(value);
        } else if (attr === "stroke") {
          // `stroke="none"` is allowed for filled shapes. Otherwise omit stroke and
          // inherit from Icon (currentColor).
          expect(["none", "currentColor"]).toContain(value);
        }
      }
    }
  });
});
