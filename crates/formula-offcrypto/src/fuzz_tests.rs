#![allow(unexpected_cfgs)]

use proptest::prelude::*;

use super::*;
use base64::engine::general_purpose::STANDARD as BASE64;
use std::sync::OnceLock;

// Keep CI runtime bounded. Heavier fuzzing can be enabled by building with
// `RUSTFLAGS=\"--cfg fuzzing\"` (or an equivalent `cfg(fuzzing)` setup).
#[cfg(fuzzing)]
const CASES: u32 = 1024;
#[cfg(not(fuzzing))]
const CASES: u32 = 64;

#[cfg(fuzzing)]
const MAX_INPUT_LEN: usize = 256 * 1024;
#[cfg(not(fuzzing))]
const MAX_INPUT_LEN: usize = 32 * 1024;

fn parseable_agile_encryption_info() -> &'static Vec<u8> {
    static CACHE: OnceLock<Vec<u8>> = OnceLock::new();
    CACHE.get_or_init(|| {
        // A deliberately *minimal* but parseable Agile `EncryptionInfo` descriptor. The verifier
        // hash ciphertext is intentionally too short for SHA1 (16 < 20), so password verification
        // deterministically fails during decryption.
        let salt_b64 = BASE64.encode([0u8; 16]);
        let ct16_b64 = BASE64.encode([0u8; 16]);
        let xml = format!(
            r#"<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
    xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
  <keyData saltValue="{salt_b64}" hashAlgorithm="SHA1" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" blockSize="16" />
  <dataIntegrity encryptedHmacKey="{ct16_b64}" encryptedHmacValue="{ct16_b64}" />
  <keyEncryptors>
    <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
      <p:encryptedKey spinCount="0" saltValue="{salt_b64}" hashAlgorithm="SHA1"
        cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" keyBits="128"
        encryptedVerifierHashInput="{ct16_b64}"
        encryptedVerifierHashValue="{ct16_b64}"
        encryptedKeyValue="{ct16_b64}" />
    </keyEncryptor>
  </keyEncryptors>
</encryption>"#
        );

        let mut bytes = Vec::new();
        bytes.extend_from_slice(&4u16.to_le_bytes());
        bytes.extend_from_slice(&4u16.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes()); // flags
        bytes.extend_from_slice(xml.as_bytes());
        bytes
    })
}

proptest! {
    #![proptest_config(ProptestConfig {
        cases: CASES,
        max_shrink_iters: 0,
        .. ProptestConfig::default()
    })]

    #[test]
    fn parse_encryption_info_agile_is_panic_free_and_rejects_malformed_xml(tail in proptest::collection::vec(any::<u8>(), 0..=MAX_INPUT_LEN)) {
        // Ensure this is not accidentally a valid XML document (which could cause a rare `Ok` and
        // make the property test flaky). Inject a byte sequence that is never valid UTF-8.
        let mut bytes = Vec::new();
        let _ = bytes.try_reserve_exact(8 + 2 + tail.len());
        bytes.extend_from_slice(&4u16.to_le_bytes());
        bytes.extend_from_slice(&4u16.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes()); // flags
        bytes.push(b'<');
        bytes.push(0xFF);
        bytes.extend_from_slice(&tail);

        let res = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| parse_encryption_info(&bytes)));
        prop_assert!(res.is_ok(), "parse_encryption_info panicked");

        let parsed = res.unwrap();
        prop_assert!(parsed.is_err(), "expected malformed agile XML to be rejected");
    }

    #[test]
    fn decrypt_encrypted_package_agile_is_panic_free_and_rejects_garbage_ciphertext(
        len_matches in any::<bool>(),
        declared_len in any::<u64>(),
        mut ciphertext in prop::collection::vec(any::<u8>(), 0..=MAX_INPUT_LEN),
    ) {
        // Ensure ciphertext (after the 8-byte original-size header) is AES-block aligned so we
        // exercise the decrypt path rather than failing immediately on framing.
        let new_len = ciphertext.len() - (ciphertext.len() % AES_BLOCK_SIZE);
        ciphertext.truncate(new_len);

        let declared_len = if len_matches {
            // Ensure `declared_len <= ciphertext.len()` so the EncryptedPackage framing checks pass
            // and we reach password verification.
            if ciphertext.is_empty() {
                0u64
            } else {
                declared_len % (ciphertext.len() as u64 + 1)
            }
        } else {
            // Ensure `declared_len > ciphertext.len()` so we exercise the size mismatch path.
            ciphertext.len() as u64 + 1
        };

        let mut encrypted_package = Vec::new();
        let _ = encrypted_package.try_reserve_exact(8 + ciphertext.len());
        encrypted_package.extend_from_slice(&declared_len.to_le_bytes());
        encrypted_package.extend_from_slice(&ciphertext);

        let mut options = DecryptOptions::default();
        // Allow `ciphertext.len() + 1` to reach the mismatch path while still bounding allocations.
        options.limits.max_output_size = Some(MAX_INPUT_LEN as u64 + 1);

        let info = parseable_agile_encryption_info();
        let res = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            decrypt_encrypted_package(info, &encrypted_package, "pw", options)
        }));
        prop_assert!(res.is_ok(), "decrypt_encrypted_package panicked");
        prop_assert!(res.unwrap().is_err(), "garbage ciphertext should not decrypt");
    }
}
