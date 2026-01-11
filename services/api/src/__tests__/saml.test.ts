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

function buildSignedSamlResponse(options: {
  assertionId: string;
  responseId: string;
  destination: string;
  inResponseTo: string;
  issuer: string;
  audience: string;
  nameId: string;
  email: string;
  name: string;
  notBefore: Date;
  notOnOrAfter: Date;
  subjectNotOnOrAfter: Date;
  authnContextClassRef?: string;
  privateKeyPem: string;
}): string {
  const issueInstant = new Date().toISOString();

  const xml = `
    <samlp:Response
      xmlns:samlp="urn:oasis:names:tc:SAML:2.0:protocol"
      xmlns:saml="urn:oasis:names:tc:SAML:2.0:assertion"
      ID="${options.responseId}"
      Version="2.0"
      IssueInstant="${issueInstant}"
      Destination="${options.destination}"
      InResponseTo="${options.inResponseTo}"
    >
      <saml:Issuer>${options.issuer}</saml:Issuer>
      <samlp:Status>
        <samlp:StatusCode Value="urn:oasis:names:tc:SAML:2.0:status:Success" />
      </samlp:Status>
      <saml:Assertion
        ID="${options.assertionId}"
        Version="2.0"
        IssueInstant="${issueInstant}"
      >
        <saml:Issuer>${options.issuer}</saml:Issuer>
        <saml:Subject>
          <saml:NameID>${options.nameId}</saml:NameID>
          <saml:SubjectConfirmation Method="urn:oasis:names:tc:SAML:2.0:cm:bearer">
            <saml:SubjectConfirmationData
              InResponseTo="${options.inResponseTo}"
              Recipient="${options.destination}"
              NotOnOrAfter="${options.subjectNotOnOrAfter.toISOString()}"
            />
          </saml:SubjectConfirmation>
        </saml:Subject>
        <saml:Conditions
          NotBefore="${options.notBefore.toISOString()}"
          NotOnOrAfter="${options.notOnOrAfter.toISOString()}"
        >
          <saml:AudienceRestriction>
            <saml:Audience>${options.audience}</saml:Audience>
          </saml:AudienceRestriction>
        </saml:Conditions>
        <saml:AuthnStatement AuthnInstant="${issueInstant}">
          <saml:AuthnContext>
            <saml:AuthnContextClassRef>${
              options.authnContextClassRef ?? "urn:oasis:names:tc:SAML:2.0:ac:classes:PasswordProtectedTransport"
            }</saml:AuthnContextClassRef>
          </saml:AuthnContext>
        </saml:AuthnStatement>
        <saml:AttributeStatement>
          <saml:Attribute Name="email">
            <saml:AttributeValue>${options.email}</saml:AttributeValue>
          </saml:Attribute>
          <saml:Attribute Name="name">
            <saml:AttributeValue>${options.name}</saml:AttributeValue>
          </saml:Attribute>
        </saml:AttributeStatement>
      </saml:Assertion>
    </samlp:Response>
  `.trim();

  const sig = new SignedXml();
  sig.signatureAlgorithm = "http://www.w3.org/2001/04/xmldsig-more#rsa-sha256";
  sig.canonicalizationAlgorithm = "http://www.w3.org/2001/10/xml-exc-c14n#";
  sig.privateKey = options.privateKeyPem;

  sig.addReference({
    xpath: `//*[local-name()='Assertion' and @ID='${options.assertionId}']`,
    transforms: [
      "http://www.w3.org/2000/09/xmldsig#enveloped-signature",
      "http://www.w3.org/2001/10/xml-exc-c14n#"
    ],
    digestAlgorithm: "http://www.w3.org/2001/04/xmlenc#sha256",
    uri: `#${options.assertionId}`
  });

  sig.computeSignature(xml, {
    location: {
      reference: `//*[local-name()='Assertion' and @ID='${options.assertionId}']/*[local-name()='Issuer']`,
      action: "after"
    }
  });

  return Buffer.from(sig.getSignedXml(), "utf8").toString("base64");
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
    publicBaseUrl: "http://localhost",
    publicBaseUrlHostAllowlist: ["localhost"],
    trustProxy: false,
    sessionCookieName: "formula_session",
    sessionTtlSeconds: 60 * 60,
    cookieSecure: false,
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

describe("SAML provider admin APIs (task spec)", () => {
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
        url: `/orgs/${orgId}/saml-providers`,
        headers: { cookie }
      });
      expect(listEmpty.statusCode).toBe(200);
      expect((listEmpty.json() as any).providers).toEqual([]);

      const invalidCert = await app.inject({
        method: "PUT",
        url: `/orgs/${orgId}/saml-providers/badcert`,
        headers: { cookie },
        payload: {
          idpEntryPoint: "http://idp.example.test/sso",
          spEntityId: "http://sp.example.test/metadata",
          idpIssuer: "https://idp.example.test/metadata",
          idpCertPem: "not a cert",
          attributeMapping: { email: "email", name: "name" },
          enabled: true
        }
      });
      expect(invalidCert.statusCode).toBe(400);
      expect((invalidCert.json() as any).error).toBe("invalid_certificate");

      const putRes = await app.inject({
        method: "PUT",
        url: `/orgs/${orgId}/saml-providers/okta`,
        headers: { cookie },
        payload: {
          idpEntryPoint: "http://idp.example.test/sso",
          spEntityId: "http://sp.example.test/metadata",
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

      const putRes2 = await app.inject({
        method: "PUT",
        url: `/orgs/${orgId}/saml-providers/okta`,
        headers: { cookie },
        payload: {
          idpEntryPoint: "http://idp.example.test/sso2",
          spEntityId: "http://sp.example.test/metadata",
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

      const list = await app.inject({ method: "GET", url: `/orgs/${orgId}/saml-providers`, headers: { cookie } });
      expect(list.statusCode).toBe(200);
      expect((list.json() as any).providers).toHaveLength(1);
      expect((list.json() as any).providers[0]).toMatchObject({ providerId: "okta", enabled: false });

      const del = await app.inject({
        method: "DELETE",
        url: `/orgs/${orgId}/saml-providers/okta`,
        headers: { cookie }
      });
      expect(del.statusCode).toBe(200);
      expect((del.json() as any).ok).toBe(true);

      const audit = await db.query(
        "SELECT event_type, details FROM audit_log WHERE event_type LIKE 'org.saml_provider.%' ORDER BY created_at ASC"
      );
      expect(audit.rows.map((row) => row.event_type)).toEqual([
        "org.saml_provider.created",
        "org.saml_provider.updated",
        "org.saml_provider.deleted"
      ]);

      const details = audit.rows.map((row) => parseJsonValue(row.details));
      expect(details[0].after.providerId).toBe("okta");
      expect(details[1].after.enabled).toBe(false);
      expect(details[2].before.providerId).toBe("okta");
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
        url: `/orgs/${orgId}/saml-providers/test`,
        headers: { cookie },
        payload: {
          idpEntryPoint: "http://idp.example.test/sso",
          spEntityId: "http://sp.example.test/metadata",
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
      const now = new Date();
      const samlResponse = buildSignedSamlResponse({
        assertionId: "_assertion-success",
        responseId: "_response-success",
        destination: callbackUrl,
        inResponseTo: requestId,
        issuer: "https://idp.example.test/metadata",
        audience: "http://sp.example.test/metadata",
        nameId: "saml-subject-123",
        email: "saml-user@example.com",
        name: "SAML User",
        notBefore: new Date(now.getTime() - 1000),
        notOnOrAfter: new Date(now.getTime() + 5 * 60 * 1000),
        subjectNotOnOrAfter: new Date(now.getTime() + 5 * 60 * 1000),
        privateKeyPem: TEST_PRIVATE_KEY_PEM
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
        provider: "saml:test",
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
    } finally {
      await app.close();
      await db.end();
    }
  });

  it(
    "rejects invalid signatures, wrong audience, expired assertions, and replayed assertion IDs",
    async () => {
      const { db, config, app } = await createTestApp();
      try {
        const ownerRegister = await app.inject({
          method: "POST",
          url: "/auth/register",
          payload: { email: "saml-owner2@example.com", password: "password1234", name: "Owner", orgName: "SSO Org 2" }
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
          url: `/orgs/${orgId}/saml-providers/test`,
          headers: { cookie },
          payload: {
            idpEntryPoint: "http://idp.example.test/sso",
            spEntityId: "http://sp.example.test/metadata",
            idpIssuer: "https://idp.example.test/metadata",
            idpCertPem: TEST_CERT_PEM,
            wantAssertionsSigned: true,
            wantResponseSigned: false,
            attributeMapping: { email: "email", name: "name" },
            enabled: true
          }
        });
        expect(putProvider.statusCode).toBe(200);

        const callbackUrl = `${config.publicBaseUrl}/auth/saml/${orgId}/test/callback`;

        // invalid signature
        {
          const start = await app.inject({ method: "GET", url: `/auth/saml/${orgId}/test/start` });
          const startUrl = new URL(start.headers.location as string);
          const requestId = extractAuthnRequestId(startUrl.searchParams.get("SAMLRequest")!);
          const relayState = startUrl.searchParams.get("RelayState")!;

          const badKey = crypto.generateKeyPairSync("rsa", { modulusLength: 2048 }).privateKey.export({
            type: "pkcs8",
            format: "pem"
          }) as string;

          const now = new Date();
          const samlResponse = buildSignedSamlResponse({
            assertionId: "_assertion-bad-sig",
            responseId: "_response-bad-sig",
            destination: callbackUrl,
            inResponseTo: requestId,
            issuer: "https://idp.example.test/metadata",
            audience: "http://sp.example.test/metadata",
            nameId: "subject-1",
            email: "user@example.com",
            name: "User",
            notBefore: new Date(now.getTime() - 1000),
            notOnOrAfter: new Date(now.getTime() + 5 * 60 * 1000),
            subjectNotOnOrAfter: new Date(now.getTime() + 5 * 60 * 1000),
            privateKeyPem: badKey
          });

          const callback = await app.inject({
            method: "POST",
            url: `/auth/saml/${orgId}/test/callback`,
            headers: { "content-type": "application/x-www-form-urlencoded" },
            payload: new URLSearchParams({ SAMLResponse: samlResponse, RelayState: relayState }).toString()
          });
          expect(callback.statusCode).toBe(401);
          expect((callback.json() as any).error).toBe("invalid_signature");
        }

        // wrong audience
        {
          const start = await app.inject({ method: "GET", url: `/auth/saml/${orgId}/test/start` });
          const startUrl = new URL(start.headers.location as string);
          const requestId = extractAuthnRequestId(startUrl.searchParams.get("SAMLRequest")!);
          const relayState = startUrl.searchParams.get("RelayState")!;

          const now = new Date();
          const samlResponse = buildSignedSamlResponse({
            assertionId: "_assertion-bad-aud",
            responseId: "_response-bad-aud",
            destination: callbackUrl,
            inResponseTo: requestId,
            issuer: "https://idp.example.test/metadata",
            audience: "wrong-audience",
            nameId: "subject-1",
            email: "user@example.com",
            name: "User",
            notBefore: new Date(now.getTime() - 1000),
            notOnOrAfter: new Date(now.getTime() + 5 * 60 * 1000),
            subjectNotOnOrAfter: new Date(now.getTime() + 5 * 60 * 1000),
            privateKeyPem: TEST_PRIVATE_KEY_PEM
          });

          const callback = await app.inject({
            method: "POST",
            url: `/auth/saml/${orgId}/test/callback`,
            headers: { "content-type": "application/x-www-form-urlencoded" },
            payload: new URLSearchParams({ SAMLResponse: samlResponse, RelayState: relayState }).toString()
          });
          expect(callback.statusCode).toBe(401);
          expect((callback.json() as any).error).toBe("invalid_audience");
        }

        // expired assertion
        {
          const start = await app.inject({ method: "GET", url: `/auth/saml/${orgId}/test/start` });
          const startUrl = new URL(start.headers.location as string);
          const requestId = extractAuthnRequestId(startUrl.searchParams.get("SAMLRequest")!);
          const relayState = startUrl.searchParams.get("RelayState")!;

          const now = new Date();
          const samlResponse = buildSignedSamlResponse({
            assertionId: "_assertion-expired",
            responseId: "_response-expired",
            destination: callbackUrl,
            inResponseTo: requestId,
            issuer: "https://idp.example.test/metadata",
            audience: "http://sp.example.test/metadata",
            nameId: "subject-1",
            email: "user@example.com",
            name: "User",
            notBefore: new Date(now.getTime() - 10 * 60 * 1000),
            notOnOrAfter: new Date(now.getTime() - 5 * 60 * 1000),
            subjectNotOnOrAfter: new Date(now.getTime() - 5 * 60 * 1000),
            privateKeyPem: TEST_PRIVATE_KEY_PEM
          });

          const callback = await app.inject({
            method: "POST",
            url: `/auth/saml/${orgId}/test/callback`,
            headers: { "content-type": "application/x-www-form-urlencoded" },
            payload: new URLSearchParams({ SAMLResponse: samlResponse, RelayState: relayState }).toString()
          });
          expect(callback.statusCode).toBe(401);
          expect((callback.json() as any).error).toBe("assertion_expired");
        }

        // replayed assertion id (distinct InResponseTo values, same assertion ID)
        {
          const assertionId = "_assertion-replay";
          const now = new Date();

          const start1 = await app.inject({ method: "GET", url: `/auth/saml/${orgId}/test/start` });
          const startUrl1 = new URL(start1.headers.location as string);
          const requestId1 = extractAuthnRequestId(startUrl1.searchParams.get("SAMLRequest")!);
          const relayState1 = startUrl1.searchParams.get("RelayState")!;

          const samlResponse1 = buildSignedSamlResponse({
            assertionId,
            responseId: "_response-replay-1",
            destination: callbackUrl,
            inResponseTo: requestId1,
            issuer: "https://idp.example.test/metadata",
            audience: "http://sp.example.test/metadata",
            nameId: "subject-1",
            email: "user@example.com",
            name: "User",
            notBefore: new Date(now.getTime() - 1000),
            notOnOrAfter: new Date(now.getTime() + 5 * 60 * 1000),
            subjectNotOnOrAfter: new Date(now.getTime() + 5 * 60 * 1000),
            privateKeyPem: TEST_PRIVATE_KEY_PEM
          });

          const ok = await app.inject({
            method: "POST",
            url: `/auth/saml/${orgId}/test/callback`,
            headers: { "content-type": "application/x-www-form-urlencoded" },
            payload: new URLSearchParams({ SAMLResponse: samlResponse1, RelayState: relayState1 }).toString()
          });
          expect(ok.statusCode).toBe(200);

          const start2 = await app.inject({ method: "GET", url: `/auth/saml/${orgId}/test/start` });
          const startUrl2 = new URL(start2.headers.location as string);
          const requestId2 = extractAuthnRequestId(startUrl2.searchParams.get("SAMLRequest")!);
          const relayState2 = startUrl2.searchParams.get("RelayState")!;

          const samlResponse2 = buildSignedSamlResponse({
            assertionId,
            responseId: "_response-replay-2",
            destination: callbackUrl,
            inResponseTo: requestId2,
            issuer: "https://idp.example.test/metadata",
            audience: "http://sp.example.test/metadata",
            nameId: "subject-1",
            email: "user@example.com",
            name: "User",
            notBefore: new Date(now.getTime() - 1000),
            notOnOrAfter: new Date(now.getTime() + 5 * 60 * 1000),
            subjectNotOnOrAfter: new Date(now.getTime() + 5 * 60 * 1000),
            privateKeyPem: TEST_PRIVATE_KEY_PEM
          });

          const replay = await app.inject({
            method: "POST",
            url: `/auth/saml/${orgId}/test/callback`,
            headers: { "content-type": "application/x-www-form-urlencoded" },
            payload: new URLSearchParams({ SAMLResponse: samlResponse2, RelayState: relayState2 }).toString()
          });
          expect(replay.statusCode).toBe(401);
          expect((replay.json() as any).error).toBe("replay_detected");
        }
      } finally {
        await app.close();
        await db.end();
      }
    },
    20_000
  );

  it("enforces org require_mfa via AuthnContextClassRef", async () => {
    const { db, config, app } = await createTestApp();
    try {
      const ownerRegister = await app.inject({
        method: "POST",
        url: "/auth/register",
        payload: { email: "saml-mfa-owner@example.com", password: "password1234", name: "Owner", orgName: "SSO Org 3" }
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
        url: `/orgs/${orgId}/saml-providers/test`,
        headers: { cookie },
        payload: {
          idpEntryPoint: "http://idp.example.test/sso",
          spEntityId: "http://sp.example.test/metadata",
          idpIssuer: "https://idp.example.test/metadata",
          idpCertPem: TEST_CERT_PEM,
          wantAssertionsSigned: true,
          wantResponseSigned: false,
          attributeMapping: { email: "email", name: "name" },
          enabled: true
        }
      });
      expect(putProvider.statusCode).toBe(200);

      await db.query("UPDATE org_settings SET require_mfa = true WHERE org_id = $1", [orgId]);

      const callbackUrl = `${config.publicBaseUrl}/auth/saml/${orgId}/test/callback`;
      const now = new Date();

      // Without MFA context.
      {
        const start = await app.inject({ method: "GET", url: `/auth/saml/${orgId}/test/start` });
        const startUrl = new URL(start.headers.location as string);
        const requestId = extractAuthnRequestId(startUrl.searchParams.get("SAMLRequest")!);
        const relayState = startUrl.searchParams.get("RelayState")!;

        const samlResponse = buildSignedSamlResponse({
          assertionId: "_assertion-mfa-missing",
          responseId: "_response-mfa-missing",
          destination: callbackUrl,
          inResponseTo: requestId,
          issuer: "https://idp.example.test/metadata",
          audience: "http://sp.example.test/metadata",
          nameId: "subject-mfa",
          email: "user@example.com",
          name: "User",
          notBefore: new Date(now.getTime() - 1000),
          notOnOrAfter: new Date(now.getTime() + 5 * 60 * 1000),
          subjectNotOnOrAfter: new Date(now.getTime() + 5 * 60 * 1000),
          privateKeyPem: TEST_PRIVATE_KEY_PEM
        });

        const callback = await app.inject({
          method: "POST",
          url: `/auth/saml/${orgId}/test/callback`,
          headers: { "content-type": "application/x-www-form-urlencoded" },
          payload: new URLSearchParams({ SAMLResponse: samlResponse, RelayState: relayState }).toString()
        });
        expect(callback.statusCode).toBe(401);
        expect((callback.json() as any).error).toBe("mfa_required");
      }

      // With MFA context.
      {
        const start = await app.inject({ method: "GET", url: `/auth/saml/${orgId}/test/start` });
        const startUrl = new URL(start.headers.location as string);
        const requestId = extractAuthnRequestId(startUrl.searchParams.get("SAMLRequest")!);
        const relayState = startUrl.searchParams.get("RelayState")!;

        const samlResponse = buildSignedSamlResponse({
          assertionId: "_assertion-mfa-ok",
          responseId: "_response-mfa-ok",
          destination: callbackUrl,
          inResponseTo: requestId,
          issuer: "https://idp.example.test/metadata",
          audience: "http://sp.example.test/metadata",
          nameId: "subject-mfa",
          email: "user@example.com",
          name: "User",
          notBefore: new Date(now.getTime() - 1000),
          notOnOrAfter: new Date(now.getTime() + 5 * 60 * 1000),
          subjectNotOnOrAfter: new Date(now.getTime() + 5 * 60 * 1000),
          authnContextClassRef: "urn:oasis:names:tc:SAML:2.0:ac:classes:TimeSyncToken",
          privateKeyPem: TEST_PRIVATE_KEY_PEM
        });

        const callback = await app.inject({
          method: "POST",
          url: `/auth/saml/${orgId}/test/callback`,
          headers: { "content-type": "application/x-www-form-urlencoded" },
          payload: new URLSearchParams({ SAMLResponse: samlResponse, RelayState: relayState }).toString()
        });
        expect(callback.statusCode).toBe(200);
      }
    } finally {
      await app.close();
      await db.end();
    }
  });
});
