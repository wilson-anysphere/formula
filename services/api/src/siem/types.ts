export type MaybeEncryptedSecret =
  | string
  | {
      /**
       * Reference to a value stored in the database-backed encrypted secret store
       * (`secrets` table).
       */
      secretRef: string;
    }
  | {
      // Backwards-compatible placeholder for configs written before the API
      // secret store existed. New configs should use `secretRef`.
      encrypted: string;
    }
  | {
      ciphertext: string;
    };

export type SiemAuthConfig =
  | { type: "none" }
  | { type: "bearer"; token: MaybeEncryptedSecret }
  | { type: "basic"; username: MaybeEncryptedSecret; password: MaybeEncryptedSecret }
  | { type: "header"; name: string; value: MaybeEncryptedSecret };

export type SiemRetryConfig = {
  maxAttempts?: number;
  baseDelayMs?: number;
  maxDelayMs?: number;
  jitter?: boolean;
};

export type SiemRedactionOptions = {
  redactionText?: string;
  sensitiveKeyPatterns?: RegExp[];
};

export type SiemBatchFormat = "json" | "cef" | "leef";

export interface SiemEndpointConfig {
  endpointUrl: string;
  /**
   * Region where the SIEM endpoint / collector processes data. Used for
   * org-level data residency enforcement on outbound exports.
   *
   * If omitted, defaults to the org's primary residency region at runtime.
   */
  dataRegion?: string;
  format?: SiemBatchFormat;
  timeoutMs?: number;
  /**
   * Header name for an idempotency key, computed deterministically from the
   * batch's event ids. If null/undefined, idempotency headers are not sent.
   */
  idempotencyKeyHeader?: string | null;
  headers?: Record<string, string>;
  auth?: SiemAuthConfig;
  retry?: SiemRetryConfig;
  redactionOptions?: SiemRedactionOptions;
  /**
   * Preferred batch size for exports. The exporter may clamp it.
   */
  batchSize?: number;
}
