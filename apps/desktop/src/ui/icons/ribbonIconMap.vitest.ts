import { describe, expect, it } from "vitest";

import { getRibbonIcon } from "./ribbonIconMap";

describe("ribbonIconMap", () => {
  it("covers Home-tab essentials used by the ribbon", () => {
    const requiredCommandIds = [
      "home.font.bold",
      "home.font.italic",
      "home.font.underline",
      "home.font.borders",
      "home.font.fillColor",
      "home.font.fontColor",
      "home.alignment.alignLeft",
      "home.alignment.center",
      "home.alignment.alignRight",
      "home.alignment.wrapText",
      "home.alignment.mergeCenter",
      "home.number.accounting",
      "home.number.percent",
      "home.number.comma",
      "home.number.increaseDecimal",
      "home.number.decreaseDecimal",
      "home.editing.sortFilter",
      "home.editing.findSelect",
    ] as const;

    for (const id of requiredCommandIds) {
      expect(getRibbonIcon(id), `Expected icon for ${id}`).toBeDefined();
    }
  });
});

