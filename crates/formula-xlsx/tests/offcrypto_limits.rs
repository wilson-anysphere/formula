use formula_xlsx::offcrypto::{parse_agile_encryption_info_stream_with_options, OffCryptoError, ParseOptions};

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
            assert_eq!(len, huge.len());
            assert_eq!(max, opts.max_base64_field_len);
        }
        other => panic!("unexpected error: {other:?}"),
    }
}
