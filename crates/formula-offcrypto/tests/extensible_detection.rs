use formula_offcrypto::{decrypt_standard_only, EncryptionType, OffcryptoError};

#[test]
fn standard_only_rejects_extensible_encryption() {
    // Extensible encryption is identified in the wild as versionMinor == 3 and
    // versionMajor ∈ {3,4}. The Standard-only decrypt entrypoint should classify this as a known
    // (but unsupported) scheme, rather than a generic version error.
    let mut encryption_info = Vec::new();
    encryption_info.extend_from_slice(&3u16.to_le_bytes()); // major
    encryption_info.extend_from_slice(&3u16.to_le_bytes()); // minor
    encryption_info.extend_from_slice(&0u32.to_le_bytes()); // flags

    let err = decrypt_standard_only(&encryption_info, &[], "pw")
        .expect_err("expected Standard-only decrypt to reject Extensible encryption");
    assert!(
        matches!(
            &err,
            OffcryptoError::UnsupportedEncryption {
                encryption_type: EncryptionType::Extensible
            }
        ),
        "expected UnsupportedEncryption(Extensible), got {err:?}"
    );
}
