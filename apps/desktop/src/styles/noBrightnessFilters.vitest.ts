import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

describe("accent hover CSS", () => {
  it("does not use filter: brightness(...) (use tokens instead)", () => {
    const stylesRoot = path.dirname(fileURLToPath(import.meta.url));
    const desktopSrcRoot = path.resolve(stylesRoot, "..");

    const targets = [
      path.join(desktopSrcRoot, "styles", "ribbon.css"),
      path.join(desktopSrcRoot, "titlebar", "titlebar.css"),
    ];

    for (const target of targets) {
      const css = fs.readFileSync(target, "utf8");
      expect(/filter\s*:\s*brightness\(/i.test(css)).toBe(false);
    }
  });
});

