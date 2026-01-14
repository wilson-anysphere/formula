use formula_xlsx::offcrypto::{decrypt_ooxml_encrypted_package, OffCryptoError};

fn standard_encryption_info_bytes(alg_id: u32, alg_id_hash: u32, key_size_bits: u32) -> Vec<u8> {
    // MS-OFFCRYPTO Standard encryption header:
    // - EncryptionVersionInfo (8 bytes): major (u16le), minor (u16le), flags (u32le)
    // - header_size (u32le)
    // - EncryptionHeader (header_size bytes): 8 DWORDs + optional CSP name (UTF-16LE)
    const MAJOR: u16 = 3;
    const MINOR: u16 = 2;
    const FLAGS: u32 = 0;
    const HEADER_SIZE: u32 = 8 * 4; // 8 DWORDs, no CSP name

    let mut out = Vec::new();
    out.extend_from_slice(&MAJOR.to_le_bytes());
    out.extend_from_slice(&MINOR.to_le_bytes());
    out.extend_from_slice(&FLAGS.to_le_bytes());
    out.extend_from_slice(&HEADER_SIZE.to_le_bytes());

    // EncryptionHeader
    out.extend_from_slice(&0u32.to_le_bytes()); // flags
    out.extend_from_slice(&0u32.to_le_bytes()); // sizeExtra
    out.extend_from_slice(&alg_id.to_le_bytes());
    out.extend_from_slice(&alg_id_hash.to_le_bytes());
    out.extend_from_slice(&key_size_bits.to_le_bytes());
    out.extend_from_slice(&0u32.to_le_bytes()); // providerType
    out.extend_from_slice(&0u32.to_le_bytes()); // reserved1
    out.extend_from_slice(&0u32.to_le_bytes()); // reserved2

    out
}

#[test]
fn unsupported_standard_alg_id_maps_to_unsupported_cipher_algorithm() {
    let encryption_info =
        standard_encryption_info_bytes(0xDEAD_BEEFu32, 0x0000_8004, 128 /* ignored */);
    let err = decrypt_ooxml_encrypted_package(&encryption_info, &[], "password")
        .expect_err("expected unsupported algorithm");
    assert!(
        matches!(err, OffCryptoError::UnsupportedCipherAlgorithm { .. }),
        "expected UnsupportedCipherAlgorithm, got {err:?}"
    );
}

#[test]
fn unsupported_standard_alg_id_hash_maps_to_unsupported_hash_algorithm() {
    // AES-128 + MD5 is invalid for Standard AES; `formula-offcrypto` reports it as an unsupported
    // algorithm (`algIdHash=...`). Ensure we surface this as an unsupported hash algorithm rather
    // than a generic "malformed Standard encryption info".
    let encryption_info = standard_encryption_info_bytes(0x0000_660Eu32, 0x0000_8003, 128);
    let err = decrypt_ooxml_encrypted_package(&encryption_info, &[], "password")
        .expect_err("expected unsupported hash algorithm");
    assert!(
        matches!(err, OffCryptoError::UnsupportedHashAlgorithm { .. }),
        "expected UnsupportedHashAlgorithm, got {err:?}"
    );
}
