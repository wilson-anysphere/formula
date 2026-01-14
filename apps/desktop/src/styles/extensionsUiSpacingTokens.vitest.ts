import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

describe("extensions-ui.css", () => {
  it("does not hardcode spacing in padding/margin/gap declarations (use --space-* tokens)", () => {
    const stylesRoot = path.dirname(fileURLToPath(import.meta.url));
    const target = path.join(stylesRoot, "extensions-ui.css");
    const css = fs.readFileSync(target, "utf8");
    const stripped = css.replace(/\/\*[\s\S]*?\*\//g, " ");

    // We intentionally keep this lightweight (regex-based) to avoid pulling in a full CSS parser.
    // The goal is specifically to prevent regressions where px values creep back into spacing
    // declarations in the extensions UI overlays/dialogs.
    const spacingDeclRe =
      /\b(?<property>padding(?:-[a-z-]+)?|margin(?:-[a-z-]+)?|(?:row|column)-gap|gap)\s*:\s*(?<value>[^;}]*)/gi;
    const pxRe = /(?<num>-?(?:\d+(?:\.\d+)?|\.\d+))px\b/g;

    const violations: string[] = [];
    let match;
    while ((match = spacingDeclRe.exec(stripped))) {
      const property = match.groups?.property ?? "";
      const value = (match.groups?.value ?? "").trim();
      const line = stripped.slice(0, match.index).split("\n").length;

      for (const pxMatch of value.matchAll(pxRe)) {
        const num = Number.parseFloat(pxMatch.groups?.num ?? "");
        if (Number.isFinite(num) && num !== 0) {
          violations.push(`${path.relative(stylesRoot, target)}:${line} ${property}: ${value}`);
          break;
        }
      }
    }

    expect(violations).toEqual([]);
  });
});
