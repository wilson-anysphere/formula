import crypto from "node:crypto";
import tls from "node:tls";
import { Agent, type Dispatcher } from "undici";

export type CheckServerIdentity = (hostname: string, cert: tls.PeerCertificate) => Error | undefined;

export function normalizeFingerprintHex(value: string): string {
  return value.replaceAll(":", "").toLowerCase();
}

export function sha256FingerprintHexFromCertRaw(raw: Buffer): string {
  return crypto.createHash("sha256").update(raw).digest("hex");
}

function isSha256FingerprintHex(value: string): boolean {
  const normalized = normalizeFingerprintHex(value);
  return /^[0-9a-f]{64}$/.test(normalized);
}

function normalizePins(rawPins: unknown): string[] {
  if (!rawPins) return [];

  if (typeof rawPins === "string") {
    try {
      return normalizePins(JSON.parse(rawPins));
    } catch {
      return [];
    }
  }

  if (!Array.isArray(rawPins)) return [];

  const normalized: string[] = [];
  for (const pin of rawPins) {
    if (typeof pin !== "string" || pin.length === 0) continue;
    normalized.push(normalizeFingerprintHex(pin));
  }

  return Array.from(new Set(normalized)).sort();
}

function pinsHash(pins: string[]): string {
  return crypto.createHash("sha256").update(pins.join(","), "utf8").digest("hex");
}

export function createPinnedCheckServerIdentity({ pins }: { pins: string[] }): CheckServerIdentity {
  if (!Array.isArray(pins) || pins.length === 0) {
    throw new TypeError("pins must be a non-empty array");
  }

  const normalizedPins = pins.map((pin) => {
    if (typeof pin !== "string" || pin.length === 0) {
      throw new TypeError("pin must be a non-empty string");
    }
    return normalizeFingerprintHex(pin);
  });

  return function checkServerIdentity(hostname, cert) {
    const defaultError = tls.checkServerIdentity(hostname, cert);
    if (defaultError) {
      (defaultError as { retriable?: boolean }).retriable = false;
      return defaultError;
    }

    const fingerprint = cert?.raw
      ? sha256FingerprintHexFromCertRaw(cert.raw)
      : typeof cert?.fingerprint256 === "string"
        ? normalizeFingerprintHex(cert.fingerprint256)
        : null;

    if (!fingerprint) {
      const err = new Error("Certificate pinning failed: certificate fingerprint not available");
      (err as { retriable?: boolean }).retriable = false;
      return err;
    }

    if (!normalizedPins.includes(normalizeFingerprintHex(fingerprint))) {
      const err = new Error("Certificate pinning failed: server certificate fingerprint mismatch");
      (err as { retriable?: boolean }).retriable = false;
      return err;
    }

    return undefined;
  };
}

export type OrgTlsPolicy = {
  certificatePinningEnabled: boolean;
  certificatePins: unknown;
  /**
   * Optional CA bundle for outbound TLS connections.
   *
   * Not currently configured via org_settings, but supported for tests / future
   * integrations.
   */
  ca?: tls.ConnectionOptions["ca"];
};

export function createTlsConnectOptions({
  certificatePinningEnabled,
  certificatePins,
  ca
}: OrgTlsPolicy): tls.ConnectionOptions {
  const options: tls.ConnectionOptions = { minVersion: "TLSv1.3" };
  if (ca) options.ca = ca;

  if (certificatePinningEnabled) {
    const pins = normalizePins(certificatePins);
    if (pins.length === 0) {
      const err = new Error("certificatePins must be non-empty when certificate pinning is enabled");
      (err as { retriable?: boolean }).retriable = false;
      throw err;
    }

    for (const pin of pins) {
      if (!isSha256FingerprintHex(pin)) {
        const err = new Error("certificatePins must be SHA-256 fingerprints (hex, optionally colon-separated)");
        (err as { retriable?: boolean }).retriable = false;
        throw err;
      }
    }

    options.checkServerIdentity = createPinnedCheckServerIdentity({ pins });
  }

  return options;
}

const agentCache = new Map<string, Agent>();

function createAgentCacheKey(policy: OrgTlsPolicy): string {
  const enabled = Boolean(policy.certificatePinningEnabled);
  const pins = enabled ? normalizePins(policy.certificatePins) : [];
  const baseKey = `${enabled ? "pinning" : "nopinning"}:${pinsHash(pins)}`;

  if (!policy.ca) return baseKey;

  const caHashInput = Array.isArray(policy.ca)
    ? policy.ca.map((part) => (typeof part === "string" ? part : part.toString("base64"))).join("|")
    : typeof policy.ca === "string"
      ? policy.ca
      : policy.ca.toString("base64");

  return `${baseKey}:ca:${crypto.createHash("sha256").update(caHashInput, "utf8").digest("hex")}`;
}

function getOrCreateAgent(policy: OrgTlsPolicy): Agent {
  const key = createAgentCacheKey(policy);
  const cached = agentCache.get(key);
  if (cached) return cached;

  const connect = createTlsConnectOptions(policy);
  const agent = new Agent({ connect });
  agentCache.set(key, agent);
  return agent;
}

export async function fetchWithOrgTls(
  url: string | URL,
  init: RequestInit = {},
  { tls: policy }: { tls?: OrgTlsPolicy } = {}
): Promise<Response> {
  if (!policy) return fetch(url, init);

  const parsed = typeof url === "string" ? new URL(url) : url;
  if (parsed.protocol !== "https:") {
    if (policy.certificatePinningEnabled) {
      const err = new Error("certificate pinning requires an https URL");
      (err as { retriable?: boolean }).retriable = false;
      throw err;
    }
    return fetch(url, init);
  }

  const dispatcher: Dispatcher = getOrCreateAgent(policy);
  return fetch(url, { ...init, dispatcher } as RequestInit);
}

export async function closeCachedOrgTlsAgents(): Promise<void> {
  const agents = Array.from(agentCache.values());
  agentCache.clear();
  await Promise.allSettled(agents.map((agent) => agent.close()));
}

// Backwards-compatible/test-friendly alias.
export const closeCachedOrgTlsAgentsForTests = closeCachedOrgTlsAgents;
