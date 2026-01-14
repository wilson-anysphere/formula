import { readdirSync, readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { stripComments } from "../../__tests__/sourceTextUtils";

function listRibbonIconSourceFiles(dir: string) {
  return readdirSync(dir)
    .filter((name) => name.endsWith(".tsx"))
    .filter((name) => !name.endsWith(".vitest.tsx"))
    .sort();
}

describe("ribbon/icons", () => {
  it("does not introduce hardcoded colors (icons must use currentColor)", () => {
    const iconsDir = join(dirname(fileURLToPath(import.meta.url)), "../icons");
    const files = listRibbonIconSourceFiles(iconsDir);

    for (const file of files) {
      const src = stripComments(readFileSync(join(iconsDir, file), "utf8"));

      expect(src, `${file} should not contain hex colors`).not.toMatch(/#[0-9a-fA-F]{3,8}\b/);
      expect(src, `${file} should not contain rgb()/hsl() colors`).not.toMatch(/\b(?:rgb|hsl)a?\(/);
      expect(src, `${file} should not contain CSS var() colors`).not.toMatch(/\bvar\(--/);

      // If fill or stroke are specified explicitly, they must still respect currentColor.
      for (const match of src.matchAll(/\b(fill|stroke)\s*=\s*"([^"]+)"/g)) {
        const [, attr, value] = match;
        if (attr === "fill") {
          expect(["none", "currentColor"]).toContain(value);
        } else if (attr === "stroke") {
          expect(["none", "currentColor"]).toContain(value);
        }
      }
    }
  });

  it("uses the shared <Icon> base component for custom ribbon icons", () => {
    const iconsDir = join(dirname(fileURLToPath(import.meta.url)), "../icons");
    const commonIconsSrc = stripComments(readFileSync(join(iconsDir, "commonIcons.tsx"), "utf8"));
    expect(commonIconsSrc).toContain("<Icon");
  });
});
