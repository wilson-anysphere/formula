class DataResidencyViolationError extends Error {
  constructor(message, { orgId, requestedRegion, allowedRegions, operation } = {}) {
    super(message);
    this.name = "DataResidencyViolationError";
    this.orgId = orgId;
    this.requestedRegion = requestedRegion;
    this.allowedRegions = allowedRegions;
    this.operation = operation;
  }
}

function uniq(list) {
  return Array.from(new Set(list));
}

function getAllowedRegions(dataResidency) {
  if (!dataResidency || typeof dataResidency !== "object") {
    throw new TypeError("dataResidency must be an object");
  }

  const region = dataResidency.region;
  if (region === "custom") {
    if (!Array.isArray(dataResidency.allowedRegions) || dataResidency.allowedRegions.length === 0) {
      throw new Error("custom residency requires allowedRegions");
    }
    return uniq(dataResidency.allowedRegions);
  }

  if (region === "us" || region === "eu" || region === "apac") {
    return [region];
  }

  throw new Error(`Unsupported residency region: ${region}`);
}

function resolvePrimaryStorageRegion(dataResidency) {
  const allowed = getAllowedRegions(dataResidency);
  const preferred = dataResidency.primaryStorageRegion ?? allowed[0];
  if (!allowed.includes(preferred)) {
    throw new Error(
      `primaryStorageRegion ${preferred} must be included in allowedRegions (${allowed.join(", ")})`
    );
  }
  return preferred;
}

function resolveAiProcessingRegion(dataResidency) {
  const allowed = getAllowedRegions(dataResidency);
  const region = dataResidency.aiProcessingRegion ?? resolvePrimaryStorageRegion(dataResidency);
  if (!allowed.includes(region) && !dataResidency.allowCrossRegionProcessing) {
    throw new Error(
      `aiProcessingRegion ${region} violates allowCrossRegionProcessing=false (allowed: ${allowed.join(
        ", "
      )})`
    );
  }
  return region;
}

function assertRegionAllowed({ orgId, dataResidency, operation, requestedRegion }) {
  const allowed = getAllowedRegions(dataResidency);
  if (!allowed.includes(requestedRegion)) {
    throw new DataResidencyViolationError(
      `Region ${requestedRegion} is not allowed for ${operation} (allowed: ${allowed.join(", ")})`,
      { orgId, requestedRegion, allowedRegions: allowed, operation }
    );
  }
}

module.exports = {
  DataResidencyViolationError,
  getAllowedRegions,
  resolveAiProcessingRegion,
  resolvePrimaryStorageRegion,
  assertRegionAllowed
};
