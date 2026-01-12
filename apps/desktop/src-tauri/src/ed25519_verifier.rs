use base64::{engine::general_purpose::STANDARD, Engine as _};
use ed25519_dalek::{Signature, VerifyingKey};
use pkcs8::DecodePublicKey;

// NOTE: Keep these limits in sync with the browser verifier in
// `shared/extension-package/v2-browser.mjs`.
const MAX_SIGNATURE_PAYLOAD_BYTES: usize = 5 * 1024 * 1024; // 5MB
const MAX_PUBLIC_KEY_PEM_BYTES: usize = 64 * 1024; // 64KB
const MAX_SIGNATURE_BASE64_BYTES: usize = 1024;

pub fn verify_ed25519_signature_payload(
    payload: &[u8],
    signature_base64: &str,
    public_key_pem: &str,
) -> Result<bool, String> {
    if payload.len() > MAX_SIGNATURE_PAYLOAD_BYTES {
        return Err(format!(
            "Payload is too large (max {MAX_SIGNATURE_PAYLOAD_BYTES} bytes)"
        ));
    }

    if public_key_pem.len() > MAX_PUBLIC_KEY_PEM_BYTES {
        return Err(format!(
            "Public key PEM is too large (max {MAX_PUBLIC_KEY_PEM_BYTES} bytes)"
        ));
    }

    if signature_base64.len() > MAX_SIGNATURE_BASE64_BYTES {
        return Err(format!(
            "Signature base64 is too large (max {MAX_SIGNATURE_BASE64_BYTES} bytes)"
        ));
    }

    let signature_bytes = STANDARD
        .decode(signature_base64.trim())
        .map_err(|err| format!("Invalid signature base64: {err}"))?;

    if signature_bytes.len() != 64 {
        return Err(format!(
            "Invalid signature length: expected 64 bytes, got {}",
            signature_bytes.len()
        ));
    }

    let signature = {
        let sig: [u8; 64] = signature_bytes
            .as_slice()
            .try_into()
            .map_err(|_| "Invalid signature length".to_string())?;
        Signature::from_bytes(&sig)
    };

    let public_key = VerifyingKey::from_public_key_pem(public_key_pem.trim())
        .map_err(|err| format!("Invalid Ed25519 public key PEM: {err}"))?;

    match public_key.verify_strict(payload, &signature) {
        Ok(()) => Ok(true),
        Err(_) => Ok(false),
    }
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub fn verify_ed25519_signature(
    payload: Vec<u8>,
    signature_base64: String,
    public_key_pem: String,
) -> Result<bool, String> {
    verify_ed25519_signature_payload(&payload, &signature_base64, &public_key_pem)
}

#[cfg(test)]
mod tests {
    use super::*;
    use base64::engine::general_purpose::STANDARD;
    use ed25519_dalek::{Signer, SigningKey};
    use pkcs8::{EncodePublicKey, LineEnding};

    fn test_keypair() -> (SigningKey, String) {
        let signing_key = SigningKey::from_bytes(&[7u8; 32]);
        let public_key_pem = signing_key
            .verifying_key()
            .to_public_key_pem(LineEnding::LF)
            .expect("expected public key PEM encoding to succeed");
        (signing_key, public_key_pem)
    }

    #[test]
    fn valid_signature_returns_true() {
        let (signing_key, public_key_pem) = test_keypair();
        let payload = b"hello world".to_vec();
        let sig = signing_key.sign(&payload);
        let signature_base64 = STANDARD.encode(sig.to_bytes());

        let ok = verify_ed25519_signature_payload(&payload, &signature_base64, &public_key_pem)
            .expect("expected verification to succeed");
        assert!(ok);
    }

    #[test]
    fn tampered_payload_returns_false() {
        let (signing_key, public_key_pem) = test_keypair();
        let payload = b"hello world".to_vec();
        let sig = signing_key.sign(&payload);
        let signature_base64 = STANDARD.encode(sig.to_bytes());

        let mut tampered = payload.clone();
        tampered[0] ^= 0x01;

        let ok = verify_ed25519_signature_payload(&tampered, &signature_base64, &public_key_pem)
            .expect("expected verification to succeed");
        assert!(!ok);
    }

    #[test]
    fn invalid_base64_returns_err() {
        let (_signing_key, public_key_pem) = test_keypair();
        let payload = b"hello world".to_vec();
        let err = verify_ed25519_signature_payload(&payload, "not base64!!!", &public_key_pem)
            .expect_err("expected invalid base64 to return Err");
        assert!(
            err.to_ascii_lowercase().contains("base64"),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn invalid_pem_returns_err() {
        let (signing_key, _public_key_pem) = test_keypair();
        let payload = b"hello world".to_vec();
        let sig = signing_key.sign(&payload);
        let signature_base64 = STANDARD.encode(sig.to_bytes());

        let err = verify_ed25519_signature_payload(&payload, &signature_base64, "not a pem")
            .expect_err("expected invalid PEM to return Err");
        assert!(
            err.to_ascii_lowercase().contains("pem"),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn invalid_signature_length_returns_err() {
        let (_signing_key, public_key_pem) = test_keypair();
        let payload = b"hello world".to_vec();
        let signature_base64 = STANDARD.encode([0u8; 10]);

        let err = verify_ed25519_signature_payload(&payload, &signature_base64, &public_key_pem)
            .expect_err("expected invalid signature length to return Err");
        assert!(
            err.to_ascii_lowercase().contains("signature length"),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn oversized_payload_returns_err() {
        let payload = vec![0u8; MAX_SIGNATURE_PAYLOAD_BYTES + 1];
        let err = verify_ed25519_signature_payload(&payload, "", "")
            .expect_err("expected oversized payload to return Err");
        assert!(
            err.to_ascii_lowercase().contains("payload is too large"),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn oversized_public_key_pem_returns_err() {
        let payload = b"hello world".to_vec();
        let oversized_pem = "A".repeat(MAX_PUBLIC_KEY_PEM_BYTES + 1);

        let err = verify_ed25519_signature_payload(&payload, "AA==", &oversized_pem)
            .expect_err("expected oversized public key PEM to return Err");
        assert!(
            err.to_ascii_lowercase().contains("public key pem is too large"),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn oversized_signature_base64_returns_err() {
        let payload = b"hello world".to_vec();
        let oversized_sig = "A".repeat(MAX_SIGNATURE_BASE64_BYTES + 1);

        let err = verify_ed25519_signature_payload(&payload, &oversized_sig, "")
            .expect_err("expected oversized signature base64 to return Err");
        assert!(
            err.to_ascii_lowercase().contains("signature base64 is too large"),
            "unexpected error: {err}"
        );
    }
}
