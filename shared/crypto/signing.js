const crypto = require("node:crypto");

function sha256(buffer) {
  return crypto.createHash("sha256").update(buffer).digest("hex");
}

function signBytes(bytes, privateKeyPem) {
  const signature = crypto.sign(null, bytes, privateKeyPem);
  return signature.toString("base64");
}

function verifyBytesSignature(bytes, signatureBase64, publicKeyPem) {
  try {
    const signature = Buffer.from(signatureBase64, "base64");
    return crypto.verify(null, bytes, publicKeyPem, signature);
  } catch {
    return false;
  }
}

module.exports = {
  sha256,
  signBytes,
  verifyBytesSignature,
};

