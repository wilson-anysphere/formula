import crypto from "node:crypto";

export function base64UrlEncode(bytes: Buffer): string {
  return bytes
    .toString("base64")
    .replaceAll("+", "-")
    .replaceAll("/", "_")
    .replace(/=+$/g, "");
}

export function randomBase64Url(bytes = 32): string {
  return base64UrlEncode(crypto.randomBytes(bytes));
}

export function sha256Base64Url(value: string): string {
  const digest = crypto.createHash("sha256").update(value, "utf8").digest();
  return base64UrlEncode(digest);
}

