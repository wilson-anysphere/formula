import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

import { stripComments } from "../../__tests__/sourceTextUtils";

describe("TableHeaderCell styles", () => {
  it("does not use inline React style objects (CSS should own layout/styling)", () => {
    const testsRoot = path.dirname(fileURLToPath(import.meta.url));
    const desktopSrcRoot = path.resolve(testsRoot, "..", "..");
    const target = path.join(desktopSrcRoot, "table", "TableHeaderCell.tsx");
    const source = stripComments(fs.readFileSync(target, "utf8"));

    // A regression guard: header styling should live in CSS (ui.css/table.css), not JSX.
    expect(/style\s*=\s*\{\{/u.test(source)).toBe(false);
  });
});
