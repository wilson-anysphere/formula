const crypto = require("node:crypto");

const SIGNATURE_ALGORITHM = "ed25519";

function sha256(buffer) {
  return crypto.createHash("sha256").update(buffer).digest("hex");
}

function requireAlgorithm(options) {
  const algorithm = options?.algorithm ?? SIGNATURE_ALGORITHM;
  if (algorithm !== SIGNATURE_ALGORITHM) {
    throw new Error(`Unsupported signature algorithm: ${algorithm} (expected ${SIGNATURE_ALGORITHM})`);
  }
  return algorithm;
}

function loadEd25519PrivateKey(privateKeyPem) {
  const key = crypto.createPrivateKey(privateKeyPem);
  if (key.asymmetricKeyType !== "ed25519") {
    throw new Error(`Unsupported private key type: ${key.asymmetricKeyType} (expected ed25519)`);
  }
  return key;
}

function loadEd25519PublicKey(publicKeyPem) {
  const key = crypto.createPublicKey(publicKeyPem);
  if (key.asymmetricKeyType !== "ed25519") {
    throw new Error(`Unsupported public key type: ${key.asymmetricKeyType} (expected ed25519)`);
  }
  return key;
}

function signBytes(bytes, privateKeyPem, options = {}) {
  requireAlgorithm(options);
  const privateKey = loadEd25519PrivateKey(privateKeyPem);
  // For Ed25519, Node's crypto.sign requires `null` and ignores the algorithm string.
  const signature = crypto.sign(null, bytes, privateKey);
  return signature.toString("base64");
}

function verifyBytesSignature(bytes, signatureBase64, publicKeyPem, options = {}) {
  try {
    requireAlgorithm(options);
    const publicKey = loadEd25519PublicKey(publicKeyPem);
    const signature = Buffer.from(signatureBase64, "base64");
    // For Ed25519, Node's crypto.verify requires `null` and ignores the algorithm string.
    return crypto.verify(null, bytes, publicKey, signature);
  } catch {
    return false;
  }
}

function generateEd25519KeyPair() {
  const { publicKey, privateKey } = crypto.generateKeyPairSync("ed25519");
  return {
    publicKeyPem: publicKey.export({ type: "spki", format: "pem" }),
    privateKeyPem: privateKey.export({ type: "pkcs8", format: "pem" }),
  };
}

module.exports = {
  SIGNATURE_ALGORITHM,
  sha256,
  signBytes,
  verifyBytesSignature,
  generateEd25519KeyPair,
};
