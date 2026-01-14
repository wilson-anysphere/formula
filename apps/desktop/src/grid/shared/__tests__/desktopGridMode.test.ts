import { describe, expect, it } from "vitest";

import { resolveDesktopGridMode } from "../desktopGridMode";

describe("resolveDesktopGridMode", () => {
  it("defaults to shared when there are no query/env overrides", () => {
    expect(resolveDesktopGridMode("", null)).toBe("shared");
  });

  it("honors env overrides", () => {
    expect(resolveDesktopGridMode("", "legacy")).toBe("legacy");
    expect(resolveDesktopGridMode("", "shared")).toBe("shared");
    expect(resolveDesktopGridMode("", "old")).toBe("legacy");
    expect(resolveDesktopGridMode("", "new")).toBe("shared");
    expect(resolveDesktopGridMode("", "1")).toBe("shared");
    expect(resolveDesktopGridMode("", "0")).toBe("legacy");
    expect(resolveDesktopGridMode("", "on")).toBe("shared");
    expect(resolveDesktopGridMode("", "off")).toBe("legacy");
    expect(resolveDesktopGridMode("", "yes")).toBe("shared");
    expect(resolveDesktopGridMode("", "no")).toBe("legacy");
    expect(resolveDesktopGridMode("", true)).toBe("shared");
    expect(resolveDesktopGridMode("", false)).toBe("legacy");
  });

  it("honors query string overrides over env", () => {
    expect(resolveDesktopGridMode("?grid=legacy", "shared")).toBe("legacy");
    expect(resolveDesktopGridMode("?grid=shared", "legacy")).toBe("shared");
  });
});
