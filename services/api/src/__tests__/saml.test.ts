import crypto from "node:crypto";
import path from "node:path";
import { fileURLToPath } from "node:url";
import zlib from "node:zlib";
import { describe, expect, it } from "vitest";
import { newDb } from "pg-mem";
import type { Pool } from "pg";
import { DOMParser } from "@xmldom/xmldom";
import { SignedXml } from "xml-crypto";
import { buildApp } from "../app";
import type { AppConfig } from "../config";
import { runMigrations } from "../db/migrations";
import { deriveSecretStoreKey } from "../secrets/secretStore";

function getMigrationsDir(): string {
  const here = path.dirname(fileURLToPath(import.meta.url));
  // services/api/src/__tests__ -> services/api/migrations
  return path.resolve(here, "../../migrations");
}

function extractCookie(setCookieHeader: string | string[] | undefined, cookieName?: string): string {
  if (!setCookieHeader) throw new Error("missing set-cookie header");
  const entries = Array.isArray(setCookieHeader) ? setCookieHeader : [setCookieHeader];
  const raw = cookieName ? entries.find((value) => value.startsWith(`${cookieName}=`)) : entries[0];
  if (!raw) throw new Error(`missing set-cookie for ${cookieName ?? "cookie"}`);
  return raw.split(";")[0];
}

function parseJsonValue(value: unknown): any {
  if (!value) return null;
  if (typeof value === "object") return value;
  if (typeof value === "string") return JSON.parse(value);
  return null;
}

const TEST_PRIVATE_KEY_PEM = `-----BEGIN PRIVATE KEY-----
MIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQDfL2/s0ZPeg7Az
0t8qqXpPm23GcBFesmTt+fDgFuI7XSqmaMEysSWZS7THleN49n34HbrjXFBUwDmG
SVVruk84+tiZNYu3vPUM8gF9CFxODxlfw/vjuIw477X1AHjVXQyaFlgkHYYsbLy5
HXnlb1COq02TWMSC3aODLWeHN45BJim9JhdaQIbVir6bYqH55dEviFCacU/hUYEy
RiSxLbi6xODCZSfKjl4fCX7UOY8zrBJpm9WxqiZO8s8twd6x8fRaEZca4cuk1eQb
2h342XeR4FXIrLomM229rTGWyKUvEHMscOL9T566z1hQdMM1rwBvSnaI82tbgsTP
xz4Pi4LFAgMBAAECggEACzQ8cC0NOUxvGgrp/SBI7Zol5qJVnOVjv7aeawF7FfPV
Ykk7+al+87UjPCnAI6BsLnp/mU5XEgocWStxSFkwBPJC0V4ox26K9r1nablzuM91
PKOAD6yCDZGrFsORTVTAfzPD9Pwucih7SOe76NKvvpnG6TC5nMA3pywuWFFnqMmD
nvMD31qm5p3RWpB6qyeig+6B3qFQnfT+zbnjBgvVwJxX/7MicJpyPJXjQrZ7i2fV
yud4CTQB6//NU6foDfexz5mY36sa4xLFUCpwmmKAACcmuApi4hvLlgxiiC7llHOc
JbWykc5BWC9gCuV6cBfbfltbfDKRKHwQs6CGeODoeQKBgQD+hncd1fOde1LFeLjq
xUNXiFvuCLN9k6P3hzXwwMZRi5bp0jPxWmRl8ltLbkMZqz3Bwe21BH+eqi0Ig687
zw1TLiZ3Fsi9XWSfBTZ1gJQ5L6BTXc3CDM0yqq69WUKikDHshd+N9zkfRK94q0te
UBUzlbYdB+4WhV1PNFuLLnKvNwKBgQDgenxVTkInWOrxKzw3d69DECOvg6SrTHaL
KxZelcO0BwV0hdA/o+rgcIXAMMJ2Ai0r3tQGymt+Jm2XFfqxXnozW0YgshYOQ+PH
RagcgB8WBs9FPh8Dk6qSs8E/wEqUpVAVkMrkFH8kixPfd+mj9SPeutsDylCz/vrh
ygMcv9eD4wKBgEKEE4cZjcPfIb93kCPSj4nFmfi4D2hG+DfM/xy+1FUlPLg4ddii
PdCiqJcq5qBDryz+qEeBOHTXllM+TsI7lwjg6659ptJOIOip7RSCGLplJuoCfq4y
uEGAd5AGTrK7KkDcr5KjRCtWwOCxK04ncZL8kg4+L0t1aPA1B6N07QW7AoGBALk/
28N6ZdWa56hHCdasipJJi2mYthg3bczDrh44cdzrvnC+zXD3kSnPMlG8633/pc+C
gG9qNPNSOzZoCQ6+7RHczS7GSLVVCXC151o90WmYDQ0TivykrCuM9Hnr8qBhHInv
h8BZueMqcygECWgpMYTppzylhZxFXD1hPNhI6U4JAoGBAJLSMB7MsoDmjHkgLaFY
UdxhwfDC/PUAzWnLBpww7BXngZCUxVPJczNiTs1rcs9Ubn7x2Awt6HACQEdJ5PhP
0N1IbKXZCtjGA2fTl+VraJQSrkjXdu7TOjIAV3qd2Bhy91ICilw4/ZuFMvWwzyj9
ygA2BaYoUus+KIHQunak7yxM
-----END PRIVATE KEY-----`;

const TEST_CERT_PEM = `-----BEGIN CERTIFICATE-----
MIIDGTCCAgGgAwIBAgIUHaLrAum1ak+fBrgKbuZaTOpB7hAwDQYJKoZIhvcNAQEL
BQAwHDEaMBgGA1UEAwwRRm9ybXVsYSBTQU1MIFRlc3QwHhcNMjYwMTExMTczNzIx
WhcNMzYwMTA5MTczNzIxWjAcMRowGAYDVQQDDBFGb3JtdWxhIFNBTUwgVGVzdDCC
ASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAN8vb+zRk96DsDPS3yqpek+b
bcZwEV6yZO358OAW4jtdKqZowTKxJZlLtMeV43j2ffgduuNcUFTAOYZJVWu6Tzj6
2Jk1i7e89QzyAX0IXE4PGV/D++O4jDjvtfUAeNVdDJoWWCQdhixsvLkdeeVvUI6r
TZNYxILdo4MtZ4c3jkEmKb0mF1pAhtWKvptiofnl0S+IUJpxT+FRgTJGJLEtuLrE
4MJlJ8qOXh8JftQ5jzOsEmmb1bGqJk7yzy3B3rHx9FoRlxrhy6TV5BvaHfjZd5Hg
VcisuiYzbb2tMZbIpS8Qcyxw4v1PnrrPWFB0wzWvAG9Kdojza1uCxM/HPg+LgsUC
AwEAAaNTMFEwHQYDVR0OBBYEFDnj2MpmRc2HrK98n+mSLJ24l+GSMB8GA1UdIwQY
MBaAFDnj2MpmRc2HrK98n+mSLJ24l+GSMA8GA1UdEwEB/wQFMAMBAf8wDQYJKoZI
hvcNAQELBQADggEBADhYShmtRLRgQKC4Smt0CuiqyRfg3UDuvSJk+mId152qapnf
zM0g8+SrLO6iNZTSeRG1guMpOYyuneu0GdFabvKRMq6gG5OVTVx/ylPZYrKsprM7
6OODKC6lvhzL9RsWjVumJIra/dqQ+hlLfR3L6tclKJd9q9t9iYI3ogGpBA8c4PLw
sdD/Jjuoti/O+G+sxjbmtTksB8s3jSWeChmw22j9kfIgS56gLhOawO/72Du/FbUH
L4FBRr1qKZoYhjE1mlyKksgq71HdGecVOkIz+FAmNNieysZtjq7xz2MwQxg3yUik
/1PstcL2hIo5tNGciUwhiGjLFTfgwAU5RTqFaGM=
-----END CERTIFICATE-----`;

function signAssertion(xml: string, assertionId: string): string {
  const sig: any = new (SignedXml as any)({
    privateKey: TEST_PRIVATE_KEY_PEM,
    publicCert: TEST_CERT_PEM,
    signatureAlgorithm: "http://www.w3.org/2001/04/xmldsig-more#rsa-sha256",
    canonicalizationAlgorithm: "http://www.w3.org/2001/10/xml-exc-c14n#"
  });
  sig.addReference({
    xpath: `//*[local-name()='Assertion' and @ID='${assertionId}']`,
    transforms: [
      "http://www.w3.org/2000/09/xmldsig#enveloped-signature",
      "http://www.w3.org/2001/10/xml-exc-c14n#"
    ],
    digestAlgorithm: "http://www.w3.org/2001/04/xmlenc#sha256"
  });
  sig.computeSignature(xml, {
    prefix: "ds",
    location: {
      reference: `//*[local-name()='Assertion' and @ID='${assertionId}']/*[local-name()='Issuer']`,
      action: "after"
    }
  });
  return sig.getSignedXml();
}

function buildSignedSamlResponse(options: {
  callbackUrl: string;
  destinationUrl?: string;
  audience: string;
  inResponseTo?: string;
  nameId: string;
  email: string;
  name: string;
}): string {
  const responseId = `_${crypto.randomUUID()}`;
  const assertionId = `_${crypto.randomUUID()}`;
  const now = new Date();
  const issueInstant = now.toISOString();
  const notBefore = new Date(now.getTime() - 5_000).toISOString();
  const notOnOrAfter = new Date(now.getTime() + 5 * 60_000).toISOString();

  const inResponseTo = options.inResponseTo ? ` InResponseTo="${options.inResponseTo}"` : "";
  const destination = options.destinationUrl ?? options.callbackUrl;
  const xml = `<?xml version="1.0"?>
<samlp:Response xmlns:samlp="urn:oasis:names:tc:SAML:2.0:protocol"
                xmlns:saml="urn:oasis:names:tc:SAML:2.0:assertion"
                ID="${responseId}"
                Version="2.0"
                IssueInstant="${issueInstant}"${inResponseTo}
                Destination="${destination}">
  <saml:Issuer>https://idp.example.test/metadata</saml:Issuer>
  <samlp:Status>
    <samlp:StatusCode Value="urn:oasis:names:tc:SAML:2.0:status:Success"/>
  </samlp:Status>
  <saml:Assertion ID="${assertionId}" Version="2.0" IssueInstant="${issueInstant}">
    <saml:Issuer>https://idp.example.test/metadata</saml:Issuer>
    <saml:Subject>
      <saml:NameID Format="urn:oasis:names:tc:SAML:1.1:nameid-format:unspecified">${options.nameId}</saml:NameID>
      <saml:SubjectConfirmation Method="urn:oasis:names:tc:SAML:2.0:cm:bearer">
        <saml:SubjectConfirmationData NotOnOrAfter="${notOnOrAfter}" Recipient="${options.callbackUrl}"${
    options.inResponseTo ? ` InResponseTo="${options.inResponseTo}"` : ""
  }/>
      </saml:SubjectConfirmation>
    </saml:Subject>
    <saml:Conditions NotBefore="${notBefore}" NotOnOrAfter="${notOnOrAfter}">
      <saml:AudienceRestriction>
        <saml:Audience>${options.audience}</saml:Audience>
      </saml:AudienceRestriction>
    </saml:Conditions>
    <saml:AuthnStatement AuthnInstant="${issueInstant}">
      <saml:AuthnContext>
        <saml:AuthnContextClassRef>urn:oasis:names:tc:SAML:2.0:ac:classes:PasswordProtectedTransport</saml:AuthnContextClassRef>
      </saml:AuthnContext>
    </saml:AuthnStatement>
    <saml:AttributeStatement>
      <saml:Attribute Name="email">
        <saml:AttributeValue xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="xs:string" xmlns:xs="http://www.w3.org/2001/XMLSchema">${options.email}</saml:AttributeValue>
      </saml:Attribute>
      <saml:Attribute Name="name">
        <saml:AttributeValue xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="xs:string" xmlns:xs="http://www.w3.org/2001/XMLSchema">${options.name}</saml:AttributeValue>
      </saml:Attribute>
    </saml:AttributeStatement>
  </saml:Assertion>
</samlp:Response>`;

  // Sanity check that the element we sign exists before computing signature.
  const parsed = new DOMParser().parseFromString(xml, "text/xml");
  const assertion = parsed.getElementsByTagName("saml:Assertion")[0];
  if (!assertion) throw new Error("missing assertion in fixture");

  const signed = signAssertion(xml, assertionId);
  return Buffer.from(signed, "utf8").toString("base64");
}

function extractAuthnRequestId(samlRequest: string): string {
  const inflated = zlib.inflateRawSync(Buffer.from(samlRequest, "base64")).toString("utf8");
  const doc = new DOMParser().parseFromString(inflated, "text/xml");
  const root = doc.documentElement;
  const id = root.getAttribute("ID");
  if (!id) throw new Error("missing AuthnRequest ID");
  return id;
}

async function createTestApp(): Promise<{
  db: Pool;
  config: AppConfig;
  app: ReturnType<typeof buildApp>;
}> {
  const mem = newDb({ autoCreateForeignKeyIndices: true });
  const pgAdapter = mem.adapters.createPg();
  const db = new pgAdapter.Pool();
  await runMigrations(db, { migrationsDir: getMigrationsDir() });

  const config: AppConfig = {
    port: 0,
    databaseUrl: "postgres://unused",
    sessionCookieName: "formula_session",
    sessionTtlSeconds: 60 * 60,
    cookieSecure: false,
    publicBaseUrl: "http://localhost",
    publicBaseUrlHostAllowlist: ["localhost"],
    trustProxy: false,
    corsAllowedOrigins: [],
    syncTokenSecret: "test-sync-secret",
    syncTokenTtlSeconds: 60,
    secretStoreKeys: {
      currentKeyId: "legacy",
      keys: { legacy: deriveSecretStoreKey("test-secret-store-key") }
    },
    localKmsMasterKey: "test-local-kms-master-key",
    awsKmsEnabled: false,
    retentionSweepIntervalMs: null,
    oidcAuthStateCleanupIntervalMs: null
  };

  const app = buildApp({ db, config });
  await app.ready();

  return { db, config, app };
}

describe("SAML provider admin APIs", () => {
  it("supports CRUD and emits audit events", async () => {
    const { db, app } = await createTestApp();
    try {
      const registerRes = await app.inject({
        method: "POST",
        url: "/auth/register",
        payload: { email: "saml-admin@example.com", password: "password1234", name: "Admin", orgName: "SAML Org" }
      });
      expect(registerRes.statusCode).toBe(200);
      const orgId = (registerRes.json() as any).organization.id as string;
      const cookie = extractCookie(registerRes.headers["set-cookie"], "formula_session");

      const listEmpty = await app.inject({
        method: "GET",
        url: `/orgs/${orgId}/saml/providers`,
        headers: { cookie }
      });
      expect(listEmpty.statusCode).toBe(200);
      expect((listEmpty.json() as any).providers).toEqual([]);

      const invalidCert = await app.inject({
        method: "PUT",
        url: `/orgs/${orgId}/saml/providers/badcert`,
        headers: { cookie },
        payload: {
          entryPoint: "http://idp.example.test/sso",
          issuer: "http://sp.example.test/metadata",
          idpIssuer: "https://idp.example.test/metadata",
          idpCertPem: "not a cert",
          wantAssertionsSigned: true,
          wantResponseSigned: false,
          attributeMapping: { email: "email", name: "name" },
          enabled: true
        }
      });
      expect(invalidCert.statusCode).toBe(400);
      expect((invalidCert.json() as any).error).toBe("invalid_certificate");

      const putRes = await app.inject({
        method: "PUT",
        url: `/orgs/${orgId}/saml/providers/okta`,
        headers: { cookie },
        payload: {
          entryPoint: "http://idp.example.test/sso",
          issuer: "http://sp.example.test/metadata",
          idpIssuer: "https://idp.example.test/metadata",
          idpCertPem: TEST_CERT_PEM,
          wantAssertionsSigned: true,
          wantResponseSigned: false,
          attributeMapping: { email: "email", name: "name" },
          enabled: true
        }
      });
      expect(putRes.statusCode).toBe(200);
      expect((putRes.json() as any).provider.providerId).toBe("okta");

      const getProvider = await app.inject({
        method: "GET",
        url: `/orgs/${orgId}/saml/providers/okta`,
        headers: { cookie }
      });
      expect(getProvider.statusCode).toBe(200);
      expect((getProvider.json() as any).provider).toMatchObject({
        providerId: "okta",
        entryPoint: "http://idp.example.test/sso",
        issuer: "http://sp.example.test/metadata",
        idpIssuer: "https://idp.example.test/metadata",
        enabled: true
      });

      const putRes2 = await app.inject({
        method: "PUT",
        url: `/orgs/${orgId}/saml/providers/okta`,
        headers: { cookie },
        payload: {
          entryPoint: "http://idp.example.test/sso2",
          issuer: "http://sp.example.test/metadata",
          idpIssuer: "https://idp.example.test/metadata",
          idpCertPem: TEST_CERT_PEM,
          wantAssertionsSigned: true,
          wantResponseSigned: false,
          attributeMapping: { email: "email", name: "name" },
          enabled: false
        }
      });
      expect(putRes2.statusCode).toBe(200);
      expect((putRes2.json() as any).provider.enabled).toBe(false);

      const list = await app.inject({ method: "GET", url: `/orgs/${orgId}/saml/providers`, headers: { cookie } });
      expect(list.statusCode).toBe(200);
      expect((list.json() as any).providers).toHaveLength(1);
      expect((list.json() as any).providers[0]).toMatchObject({ providerId: "okta", enabled: false });

      const del = await app.inject({ method: "DELETE", url: `/orgs/${orgId}/saml/providers/okta`, headers: { cookie } });
      expect(del.statusCode).toBe(200);
      expect((del.json() as any).ok).toBe(true);

      const audit = await db.query(
        "SELECT event_type, details FROM audit_log WHERE event_type LIKE 'admin.integration_%' ORDER BY created_at ASC"
      );
      const events = audit.rows.map((row) => ({ event_type: row.event_type, details: parseJsonValue(row.details) }));
      expect(events.map((event) => event.event_type)).toEqual([
        "admin.integration_added",
        "admin.integration_updated",
        "admin.integration_removed"
      ]);
      for (const event of events) {
        expect(event.details).toMatchObject({ type: "saml", providerId: "okta" });
      }
    } finally {
      await app.close();
      await db.end();
    }
  });

  it("accepts URN entity IDs for issuer fields", async () => {
    const { db, app } = await createTestApp();
    try {
      const registerRes = await app.inject({
        method: "POST",
        url: "/auth/register",
        payload: { email: "saml-urn-admin@example.com", password: "password1234", name: "Admin", orgName: "SAML URN Org" }
      });
      expect(registerRes.statusCode).toBe(200);
      const orgId = (registerRes.json() as any).organization.id as string;
      const cookie = extractCookie(registerRes.headers["set-cookie"], "formula_session");

      const putRes = await app.inject({
        method: "PUT",
        url: `/orgs/${orgId}/saml/providers/urn`,
        headers: { cookie },
        payload: {
          entryPoint: "http://idp.example.test/sso",
          issuer: "urn:formula:test:sp",
          idpIssuer: "urn:formula:test:idp",
          idpCertPem: TEST_CERT_PEM,
          wantAssertionsSigned: true,
          wantResponseSigned: false,
          attributeMapping: { email: "email", name: "name" },
          enabled: true
        }
      });
      expect(putRes.statusCode).toBe(200);
      expect((putRes.json() as any).provider).toMatchObject({
        providerId: "urn",
        issuer: "urn:formula:test:sp",
        idpIssuer: "urn:formula:test:idp"
      });

      const getProvider = await app.inject({
        method: "GET",
        url: `/orgs/${orgId}/saml/providers/urn`,
        headers: { cookie }
      });
      expect(getProvider.statusCode).toBe(200);
      expect((getProvider.json() as any).provider).toMatchObject({
        providerId: "urn",
        issuer: "urn:formula:test:sp",
        idpIssuer: "urn:formula:test:idp"
      });
    } finally {
      await app.close();
      await db.end();
    }
  });
});

describe("SAML SSO", () => {
  it("successful login provisions user + membership and issues a session", async () => {
    const { db, config, app } = await createTestApp();
    try {
      const ownerRegister = await app.inject({
        method: "POST",
        url: "/auth/register",
        payload: { email: "saml-owner@example.com", password: "password1234", name: "Owner", orgName: "SSO Org" }
      });
      expect(ownerRegister.statusCode).toBe(200);
      const orgId = (ownerRegister.json() as any).organization.id as string;

      await db.query("UPDATE org_settings SET allowed_auth_methods = $2::jsonb WHERE org_id = $1", [
        orgId,
        JSON.stringify(["password", "saml"])
      ]);

      const cookie = extractCookie(ownerRegister.headers["set-cookie"]);
      const putProvider = await app.inject({
        method: "PUT",
        url: `/orgs/${orgId}/saml/providers/test`,
        headers: { cookie },
        payload: {
          entryPoint: "http://idp.example.test/sso",
          issuer: "http://sp.example.test/metadata",
          idpIssuer: "https://idp.example.test/metadata",
          // Duplicate cert to exercise certificate-bundle handling (rollover support).
          idpCertPem: `${TEST_CERT_PEM}\n${TEST_CERT_PEM}`,
          wantAssertionsSigned: true,
          wantResponseSigned: false,
          attributeMapping: { email: "email", name: "name" },
          enabled: true
        }
      });
      expect(putProvider.statusCode).toBe(200);

      const metadataRes = await app.inject({ method: "GET", url: `/auth/saml/${orgId}/test/metadata` });
      expect(metadataRes.statusCode).toBe(200);
      expect(String(metadataRes.headers["content-type"] ?? "")).toContain("application/xml");
      const metadataXml = metadataRes.body;
      expect(metadataXml).toContain(`entityID="http://sp.example.test/metadata"`);
      expect(metadataXml).toContain(`Location="${config.publicBaseUrl}/auth/saml/${orgId}/test/callback"`);

      const startRes = await app.inject({ method: "GET", url: `/auth/saml/${orgId}/test/start` });
      expect(startRes.statusCode).toBe(302);
      const startUrl = new URL(startRes.headers.location as string);
      expect(`${startUrl.origin}${startUrl.pathname}`).toBe("http://idp.example.test/sso");
      const samlRequest = startUrl.searchParams.get("SAMLRequest");
      expect(samlRequest).toBeTruthy();
      const requestId = extractAuthnRequestId(samlRequest!);
      const relayState = startUrl.searchParams.get("RelayState");
      expect(relayState).toBeTruthy();

      const callbackUrl = `${config.publicBaseUrl}/auth/saml/${orgId}/test/callback`;
      const samlResponse = buildSignedSamlResponse({
        callbackUrl,
        audience: "http://sp.example.test/metadata",
        inResponseTo: requestId,
        nameId: "saml-subject-123",
        email: "saml-user@example.com",
        name: "SAML User"
      });

      const callbackRes = await app.inject({
        method: "POST",
        url: `/auth/saml/${orgId}/test/callback`,
        headers: { "content-type": "application/x-www-form-urlencoded" },
        payload: new URLSearchParams({ SAMLResponse: samlResponse, RelayState: relayState! }).toString()
      });
      expect(callbackRes.statusCode).toBe(200);

      const sessionCookie = extractCookie(callbackRes.headers["set-cookie"], config.sessionCookieName);
      expect(sessionCookie.startsWith(`${config.sessionCookieName}=`)).toBe(true);

      const me = await app.inject({ method: "GET", url: "/me", headers: { cookie: sessionCookie } });
      expect(me.statusCode).toBe(200);
      const meBody = me.json() as any;
      expect(meBody.user.email).toBe("saml-user@example.com");
      expect(meBody.organizations.some((o: any) => o.id === orgId)).toBe(true);

      const identities = await db.query(
        "SELECT provider, subject, email, org_id FROM user_identities WHERE org_id = $1",
        [orgId]
      );
      expect(identities.rowCount).toBe(1);
      expect(identities.rows[0]).toMatchObject({
        provider: "test",
        subject: "saml-subject-123",
        email: "saml-user@example.com",
        org_id: orgId
      });

      const audit = await db.query(
        "SELECT event_type, org_id, details FROM audit_log WHERE event_type = 'auth.login' ORDER BY created_at DESC LIMIT 1"
      );
      expect(audit.rowCount).toBe(1);
      expect(audit.rows[0].org_id).toBe(orgId);
      expect(parseJsonValue(audit.rows[0].details)).toMatchObject({ method: "saml", provider: "test" });

      // Replay the same response against a fresh RelayState should fail because the
      // original AuthnRequest ID (InResponseTo) is consumed.
      const replayStart = await app.inject({ method: "GET", url: `/auth/saml/${orgId}/test/start` });
      expect(replayStart.statusCode).toBe(302);
      const replayUrl = new URL(replayStart.headers.location as string);
      const replayState = replayUrl.searchParams.get("RelayState");
      expect(replayState).toBeTruthy();

      const replayRes = await app.inject({
        method: "POST",
        url: `/auth/saml/${orgId}/test/callback`,
        headers: { "content-type": "application/x-www-form-urlencoded" },
        payload: new URLSearchParams({ SAMLResponse: samlResponse, RelayState: replayState! }).toString()
      });
      expect(replayRes.statusCode).toBe(401);
      expect((replayRes.json() as any).error).toBe("invalid_saml_response");

      // If an IdP issuer is configured, responses from a different issuer should be rejected.
      const updateProvider = await app.inject({
        method: "PUT",
        url: `/orgs/${orgId}/saml/providers/test`,
        headers: { cookie },
        payload: {
          entryPoint: "http://idp.example.test/sso",
          issuer: "http://sp.example.test/metadata",
          idpIssuer: "https://idp.other.example.test/metadata",
          idpCertPem: TEST_CERT_PEM,
          wantAssertionsSigned: true,
          wantResponseSigned: false,
          attributeMapping: { email: "email", name: "name" },
          enabled: true
        }
      });
      expect(updateProvider.statusCode).toBe(200);

      const issuerStart = await app.inject({ method: "GET", url: `/auth/saml/${orgId}/test/start` });
      expect(issuerStart.statusCode).toBe(302);
      const issuerUrl = new URL(issuerStart.headers.location as string);
      const issuerRequestId = extractAuthnRequestId(issuerUrl.searchParams.get("SAMLRequest")!);
      const issuerRelayState = issuerUrl.searchParams.get("RelayState");
      expect(issuerRelayState).toBeTruthy();

      const issuerResponse = buildSignedSamlResponse({
        callbackUrl,
        audience: "http://sp.example.test/metadata",
        inResponseTo: issuerRequestId,
        nameId: "saml-subject-123",
        email: "saml-user@example.com",
        name: "SAML User"
      });

      const issuerCallback = await app.inject({
        method: "POST",
        url: `/auth/saml/${orgId}/test/callback`,
        headers: { "content-type": "application/x-www-form-urlencoded" },
        payload: new URLSearchParams({ SAMLResponse: issuerResponse, RelayState: issuerRelayState! }).toString()
      });
      expect(issuerCallback.statusCode).toBe(401);
      expect((issuerCallback.json() as any).error).toBe("invalid_saml_response");

      // Destination is outside the signed assertion for many IdPs. If present, it must match our ACS URL.
      const destinationStart = await app.inject({ method: "GET", url: `/auth/saml/${orgId}/test/start` });
      expect(destinationStart.statusCode).toBe(302);
      const destinationUrl = new URL(destinationStart.headers.location as string);
      const destinationRequestId = extractAuthnRequestId(destinationUrl.searchParams.get("SAMLRequest")!);
      const destinationRelayState = destinationUrl.searchParams.get("RelayState");
      expect(destinationRelayState).toBeTruthy();

      const destinationResponse = buildSignedSamlResponse({
        callbackUrl,
        destinationUrl: "http://evil.example.test/callback",
        audience: "http://sp.example.test/metadata",
        inResponseTo: destinationRequestId,
        nameId: "saml-subject-123",
        email: "saml-user@example.com",
        name: "SAML User"
      });

      const destinationCallback = await app.inject({
        method: "POST",
        url: `/auth/saml/${orgId}/test/callback`,
        headers: { "content-type": "application/x-www-form-urlencoded" },
        payload: new URLSearchParams({ SAMLResponse: destinationResponse, RelayState: destinationRelayState! }).toString()
      });
      expect(destinationCallback.statusCode).toBe(401);
      expect((destinationCallback.json() as any).error).toBe("invalid_saml_response");
    } finally {
      await app.close();
      await db.end();
    }
  });
});
