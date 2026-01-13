use formula_offcrypto::{
    parse_encrypted_package_header, parse_encryption_info, validate_agile_segment_decrypt_inputs,
    validate_standard_encrypted_package_stream, OffcryptoError,
};

#[test]
fn truncated_encryption_info_less_than_8_bytes_errors() {
    // EncryptionInfo stream starts with: u16 major, u16 minor, u32 flags (8 bytes total).
    // Provide fewer than 8 bytes and ensure we get a structured error (never panic).
    let bytes = [0u8; 7];
    let err = parse_encryption_info(&bytes).unwrap_err();
    assert!(matches!(err, OffcryptoError::Truncated { .. }));
}

#[test]
fn agile_header_ok_but_xml_missing_required_attrs_errors() {
    // Minimal Agile header:
    // - major=4, minor=4, flags=0
    // - XML with <keyData> but missing most required attributes.
    let xml = br#"<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"><keyData saltSize="16"/></encryption>"#;

    let mut bytes = Vec::new();
    bytes.extend_from_slice(&4u16.to_le_bytes());
    bytes.extend_from_slice(&4u16.to_le_bytes());
    bytes.extend_from_slice(&0u32.to_le_bytes()); // flags
    bytes.extend_from_slice(xml);

    let err = parse_encryption_info(&bytes).unwrap_err();
    assert!(matches!(
        err,
        OffcryptoError::InvalidEncryptionInfo { .. }
    ));
}

#[test]
fn standard_header_encryption_header_size_larger_than_buffer_errors() {
    // Standard (3.2) header, but with a header_size that exceeds the available bytes.
    let mut bytes = Vec::new();
    bytes.extend_from_slice(&3u16.to_le_bytes());
    bytes.extend_from_slice(&2u16.to_le_bytes());
    bytes.extend_from_slice(&0u32.to_le_bytes()); // flags
    bytes.extend_from_slice(&100u32.to_le_bytes()); // header_size (too large for empty remainder)

    let err = parse_encryption_info(&bytes).unwrap_err();
    assert!(matches!(err, OffcryptoError::Truncated { .. }));
}

#[test]
fn encrypted_package_shorter_than_8_bytes_errors() {
    let err = parse_encrypted_package_header(&[0u8; 7]).unwrap_err();
    assert!(matches!(err, OffcryptoError::Truncated { .. }));
}

#[test]
fn standard_encrypted_package_ciphertext_not_multiple_of_16_errors() {
    // Ciphertext length after the 8-byte original-size prefix must be block-aligned.
    let mut encrypted_package = 0u64.to_le_bytes().to_vec();
    encrypted_package.extend_from_slice(&[0u8; 15]); // not a multiple of 16

    let err = validate_standard_encrypted_package_stream(&encrypted_package).unwrap_err();
    assert_eq!(err, OffcryptoError::InvalidCiphertextLength { len: 15 });
}

#[test]
fn agile_segment_decrypt_wrong_lengths_errors() {
    // expected_plaintext_len=17 implies at least 32 bytes of ciphertext for block alignment.
    let iv = [0u8; 16];
    let ciphertext = [0u8; 16];

    let err = validate_agile_segment_decrypt_inputs(&iv, &ciphertext, 17).unwrap_err();
    assert!(matches!(err, OffcryptoError::InvalidEncryptionInfo { .. }));
}
