export class DataResidencyViolationError extends Error {
  readonly orgId?: string;
  readonly requestedRegion: string;
  readonly allowedRegions: string[];
  readonly operation: string;

  constructor(
    message: string,
    options: { orgId?: string; requestedRegion: string; allowedRegions: string[]; operation: string }
  ) {
    super(message);
    this.name = "DataResidencyViolationError";
    this.orgId = options.orgId;
    this.requestedRegion = options.requestedRegion;
    this.allowedRegions = options.allowedRegions;
    this.operation = options.operation;
    Object.setPrototypeOf(this, new.target.prototype);
  }
}

function uniq(values: string[]): string[] {
  return Array.from(new Set(values));
}

function toNonEmptyStrings(value: unknown): string[] {
  if (!Array.isArray(value)) return [];
  return value
    .filter((item): item is string => typeof item === "string")
    .map((item) => item.trim())
    .filter((item) => item.length > 0);
}

export function getAllowedRegions(options: { region: string; allowedRegions?: unknown }): string[] {
  const region = options.region;
  if (region === "custom") {
    const allowed = uniq(toNonEmptyStrings(options.allowedRegions));
    if (allowed.length === 0) {
      throw new Error("custom residency requires allowedRegions");
    }
    return allowed;
  }

  if (region === "us" || region === "eu" || region === "apac") {
    return [region];
  }

  throw new Error(`Unsupported residency region: ${region}`);
}

export function resolvePrimaryStorageRegion(options: {
  region: string;
  allowedRegions?: unknown;
  primaryStorageRegion?: string | null | undefined;
}): string {
  const allowed = getAllowedRegions({ region: options.region, allowedRegions: options.allowedRegions });
  const preferred = options.primaryStorageRegion ?? allowed[0];
  if (!allowed.includes(preferred)) {
    throw new Error(
      `primaryStorageRegion ${preferred} must be included in allowedRegions (${allowed.join(", ")})`
    );
  }
  return preferred;
}

export function resolveAiProcessingRegion(options: {
  region: string;
  allowedRegions?: unknown;
  aiProcessingRegion?: string | null | undefined;
  allowCrossRegionProcessing: boolean;
  primaryStorageRegion?: string | null | undefined;
}): string {
  const allowed = getAllowedRegions({ region: options.region, allowedRegions: options.allowedRegions });
  const region =
    options.aiProcessingRegion ??
    resolvePrimaryStorageRegion({
      region: options.region,
      allowedRegions: options.allowedRegions,
      primaryStorageRegion: options.primaryStorageRegion
    });

  if (!allowed.includes(region) && !options.allowCrossRegionProcessing) {
    throw new Error(
      `aiProcessingRegion ${region} violates allowCrossRegionProcessing=false (allowed: ${allowed.join(", ")})`
    );
  }
  return region;
}

export function assertRegionAllowed(options: {
  orgId?: string;
  requestedRegion: string;
  operation: string;
  allowedRegions: string[];
}): void {
  if (!options.allowedRegions.includes(options.requestedRegion)) {
    throw new DataResidencyViolationError(
      `Region ${options.requestedRegion} is not allowed for ${options.operation} (allowed: ${options.allowedRegions.join(
        ", "
      )})`,
      {
        orgId: options.orgId,
        requestedRegion: options.requestedRegion,
        allowedRegions: options.allowedRegions,
        operation: options.operation
      }
    );
  }
}

export function assertOutboundRegionAllowed(options: {
  orgId?: string;
  requestedRegion: string;
  operation: string;
  region: string;
  allowedRegions?: unknown;
  allowCrossRegionProcessing: boolean;
}): void {
  if (options.allowCrossRegionProcessing) return;
  const allowed = getAllowedRegions({ region: options.region, allowedRegions: options.allowedRegions });
  assertRegionAllowed({
    orgId: options.orgId,
    requestedRegion: options.requestedRegion,
    operation: options.operation,
    allowedRegions: allowed
  });
}

