import { describe, expect, it } from "vitest";
import {
  assertOutboundRegionAllowed,
  assertRegionAllowedInSet,
  getAllowedRegions,
  resolveAiProcessingRegion,
  DataResidencyViolationError
} from "../policies/dataResidency";

describe("dataResidency policy", () => {
  it("enforces custom allowedRegions (unique, non-empty)", () => {
    expect(
      getAllowedRegions({
        region: "custom",
        allowedRegions: ["us", "eu", "us", "", "  "]
      })
    ).toEqual(["us", "eu"]);

    expect(() => getAllowedRegions({ region: "custom", allowedRegions: [] })).toThrow(
      "custom residency requires allowedRegions"
    );
  });

  it("validates aiProcessingRegion against allowCrossRegionProcessing", () => {
    expect(
      resolveAiProcessingRegion({
        region: "eu",
        allowCrossRegionProcessing: true,
        aiProcessingRegion: "us"
      })
    ).toBe("us");

    expect(() =>
      resolveAiProcessingRegion({
        region: "eu",
        allowCrossRegionProcessing: false,
        aiProcessingRegion: "us"
      })
    ).toThrow(/allowCrossRegionProcessing=false/);
  });

  it("allows/blocks outbound regions based on allowCrossRegionProcessing", () => {
    // Cross-region allowed => any region is permitted.
    expect(() =>
      assertOutboundRegionAllowed({
        orgId: "org_1",
        operation: "test.outbound",
        requestedRegion: "us",
        region: "eu",
        allowCrossRegionProcessing: true
      })
    ).not.toThrow();

    // Cross-region disabled => region must be in allowedRegions.
    expect(() =>
      assertOutboundRegionAllowed({
        orgId: "org_1",
        operation: "test.outbound",
        requestedRegion: "us",
        region: "eu",
        allowCrossRegionProcessing: false
      })
    ).toThrow(DataResidencyViolationError);

    const allowed = getAllowedRegions({ region: "eu" });
    expect(() =>
      assertRegionAllowedInSet({
        orgId: "org_1",
        operation: "test.assert",
        requestedRegion: "us",
        allowedRegions: allowed
      })
    ).toThrow(DataResidencyViolationError);
  });
});
