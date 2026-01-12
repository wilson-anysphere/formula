import { describe, expect, it } from "vitest";
import { loadConfig } from "../config";

describe("loadConfig production validation", () => {
  it("rejects default dev secrets in production", () => {
    expect(() =>
      loadConfig({
        NODE_ENV: "production",
        DATABASE_URL: "postgres://user:pass@db:5432/formula",
        COOKIE_SECURE: "true",
        PUBLIC_BASE_URL: "https://api.example.com"
      })
    ).toThrow(/SYNC_TOKEN_SECRET|SECRET_STORE_KEY|LOCAL_KMS_MASTER_KEY/);
  });

  it("accepts explicit secure settings in production", () => {
    const cfg = loadConfig({
      NODE_ENV: "production",
      PORT: "4000",
      DATABASE_URL: "postgres://user:pass@db:5432/formula",
      COOKIE_SECURE: "true",
      PUBLIC_BASE_URL: "https://api.example.com",
      SYNC_TOKEN_SECRET: "prod-sync-token-secret",
      SECRET_STORE_KEY: "prod-secret-store-key",
      LOCAL_KMS_MASTER_KEY: "prod-local-kms-master-key",
      CORS_ALLOWED_ORIGINS: "https://app.example.com"
    });

    expect(cfg.port).toBe(4000);
    expect(cfg.cookieSecure).toBe(true);
    expect(cfg.publicBaseUrl).toBe("https://api.example.com");
    expect(cfg.corsAllowedOrigins).toEqual(["https://app.example.com"]);
  });
});

