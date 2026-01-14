use formula_xlsx::offcrypto::{
    parse_agile_encryption_info_stream_with_options,
    parse_agile_encryption_info_stream_with_options_and_decrypt_options, DecryptOptions,
    OffCryptoError, ParseOptions,
};

#[test]
fn encryption_info_xml_size_limit_is_enforced() {
    let opts = ParseOptions {
        max_encryption_info_xml_len: 16,
        ..ParseOptions::default()
    };

    // Minimal `EncryptionInfo` stream: 8-byte header + XML bytes.
    let mut stream = Vec::new();
    // versionMajor=4, versionMinor=4, flags=0
    stream.extend_from_slice(&4u16.to_le_bytes());
    stream.extend_from_slice(&4u16.to_le_bytes());
    stream.extend_from_slice(&0u32.to_le_bytes());
    stream.extend_from_slice(&vec![b'a'; opts.max_encryption_info_xml_len + 1]);

    let err = parse_agile_encryption_info_stream_with_options(&stream, &opts)
        .expect_err("expected size limit error");
    match err {
        OffCryptoError::EncryptionInfoTooLarge { len, max } => {
            assert_eq!(len, opts.max_encryption_info_xml_len + 1);
            assert_eq!(max, opts.max_encryption_info_xml_len);
        }
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn huge_base64_attribute_is_rejected_before_decoding() {
    // Keep the XML itself small enough to pass `max_encryption_info_xml_len`, but configure a small
    // base64 field limit so we can trigger the base64-path check deterministically.
    let opts = ParseOptions {
        max_encryption_info_xml_len: 4096,
        max_base64_field_len: 64,
        max_base64_decoded_len: 1024,
    };

    // 68 chars (multiple of 4) -> exceeds max_base64_field_len=64.
    let huge = "A".repeat(68);
    let xml = format!(
        r#"<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
            xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
              <keyData saltSize="16" blockSize="16" keyBits="128" hashSize="20"
                       cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA1"
                       saltValue="{huge}"/>
              <keyEncryptors>
                <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
                   <p:encryptedKey saltSize="16" blockSize="16" keyBits="128" hashSize="20"
                                   spinCount="1" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA1"
                                   saltValue="AA=="
                                   encryptedVerifierHashInput="AA=="
                                   encryptedVerifierHashValue="AA=="
                                   encryptedKeyValue="AA=="/>
                </keyEncryptor>
              </keyEncryptors>
            </encryption>"#
    );

    let mut stream = Vec::new();
    stream.extend_from_slice(&4u16.to_le_bytes());
    stream.extend_from_slice(&4u16.to_le_bytes());
    stream.extend_from_slice(&0u32.to_le_bytes());
    stream.extend_from_slice(xml.as_bytes());

    let err = parse_agile_encryption_info_stream_with_options(&stream, &opts)
        .expect_err("expected size limit");
    match err {
        OffCryptoError::FieldTooLarge { field, len, max } => {
            assert_eq!(field, "saltValue");
            assert!(len > opts.max_base64_field_len);
            assert_eq!(max, opts.max_base64_field_len);
        }
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn spin_count_limit_is_checked_before_decoding_password_salt() {
    // Configure a tiny base64 field limit so `encryptedKey@saltValue` would fail if we tried to
    // decode it. The parser should reject oversized `spinCount` values before touching base64
    // fields to fail fast on malicious inputs.
    let opts = ParseOptions {
        max_encryption_info_xml_len: 4096,
        max_base64_field_len: 4,
        max_base64_decoded_len: 1024,
    };
    let decrypt_opts = DecryptOptions::default();

    let too_large = u32::MAX;
    // 8 chars -> exceeds max_base64_field_len=4.
    let huge = "A".repeat(8);

    let xml = format!(
        r#"<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
            xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
              <keyData saltSize="1" blockSize="16" keyBits="128" hashSize="20"
                       cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA1"
                       saltValue="AA=="/>
              <keyEncryptors>
                <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
                   <p:encryptedKey saltSize="1" blockSize="16" keyBits="128" hashSize="20"
                                   spinCount="{too_large}" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA1"
                                   saltValue="{huge}"
                                   encryptedVerifierHashInput="AA=="
                                   encryptedVerifierHashValue="AA=="
                                   encryptedKeyValue="AA=="/>
                </keyEncryptor>
              </keyEncryptors>
            </encryption>"#
    );

    let mut stream = Vec::new();
    stream.extend_from_slice(&4u16.to_le_bytes());
    stream.extend_from_slice(&4u16.to_le_bytes());
    stream.extend_from_slice(&0u32.to_le_bytes());
    stream.extend_from_slice(xml.as_bytes());

    let err = parse_agile_encryption_info_stream_with_options_and_decrypt_options(
        &stream,
        &opts,
        &decrypt_opts,
    )
    .expect_err("expected spinCount limit error");
    match err {
        OffCryptoError::SpinCountTooLarge { spin_count, max } => {
            assert_eq!(spin_count, too_large);
            assert_eq!(max, decrypt_opts.max_spin_count);
        }
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn decoded_len_limit_allows_padded_base64_exact_size() {
    // Base64 with padding should not be rejected due to a loose decoded-length upper bound.
    let opts = ParseOptions {
        max_encryption_info_xml_len: 4096,
        max_base64_field_len: 1024,
        max_base64_decoded_len: 16,
    };

    // 16 bytes of zeros -> 24 base64 chars with "==" padding.
    let salt_16b = "AAAAAAAAAAAAAAAAAAAAAA==";

    let xml = format!(
        r#"<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
            xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
              <keyData saltSize="16" blockSize="16" keyBits="128" hashSize="20"
                       cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA1"
               saltValue="{salt_16b}"/>
                <keyEncryptors>
                <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
                  <p:encryptedKey saltSize="16" blockSize="16" keyBits="128" hashSize="20"
                                  spinCount="1" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA1"
                                  saltValue="{salt_16b}"
                                  encryptedVerifierHashInput="AA=="
                                  encryptedVerifierHashValue="AA=="
                                  encryptedKeyValue="AA=="/>
                </keyEncryptor>
              </keyEncryptors>
            </encryption>"#
    );

    let mut stream = Vec::new();
    stream.extend_from_slice(&4u16.to_le_bytes());
    stream.extend_from_slice(&4u16.to_le_bytes());
    stream.extend_from_slice(&0u32.to_le_bytes());
    stream.extend_from_slice(xml.as_bytes());

    parse_agile_encryption_info_stream_with_options(&stream, &opts)
        .expect("expected padded base64 to pass exact decoded-len limit");
}
