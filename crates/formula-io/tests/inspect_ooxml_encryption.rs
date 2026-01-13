use std::io::{Cursor, Write};

use formula_io::inspect_ooxml_encryption;
use formula_offcrypto::{
    AgileEncryptionInfoSummary, EncryptionType, HashAlgorithm, StandardAlgId,
    StandardEncryptionInfoSummary,
};

fn ole_encrypted_ooxml_container(encryption_info: &[u8]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");

    // Use non-canonical casing to validate best-effort (case-insensitive) stream lookup.
    {
        let mut s = ole
            .create_stream("encryptioninfo")
            .expect("create EncryptionInfo stream");
        s.write_all(encryption_info)
            .expect("write EncryptionInfo bytes");
    }
    ole.create_stream("ENCRYPTEDPACKAGE")
        .expect("create EncryptedPackage stream");

    ole.into_inner().into_inner()
}

fn standard_encryption_info_bytes() -> Vec<u8> {
    // Copied from `formula-offcrypto` unit tests (`inspects_minimal_standard_encryption_info`).
    // Minimal Standard EncryptionInfo buffer sufficient for `inspect_encryption_info`:
    // - version (3.2)
    // - header size + header (AES-256 + SHA1, keySize matches algId)
    // - verifier with saltSize=16, verifierHashSize=20 (SHA1) and a 32-byte encrypted hash
    let mut bytes = Vec::new();
    bytes.extend_from_slice(&3u16.to_le_bytes());
    bytes.extend_from_slice(&2u16.to_le_bytes());
    bytes.extend_from_slice(&0u32.to_le_bytes());

    let mut header = Vec::new();
    header.extend_from_slice(&0u32.to_le_bytes()); // flags
    header.extend_from_slice(&0u32.to_le_bytes()); // sizeExtra
    header.extend_from_slice(&0x0000_6610u32.to_le_bytes()); // algId = CALG_AES_256
    header.extend_from_slice(&0x0000_8004u32.to_le_bytes()); // algIdHash = CALG_SHA1
    header.extend_from_slice(&256u32.to_le_bytes()); // keySize
    header.extend_from_slice(&0u32.to_le_bytes()); // providerType
    header.extend_from_slice(&0u32.to_le_bytes()); // reserved1
    header.extend_from_slice(&0u32.to_le_bytes()); // reserved2

    bytes.extend_from_slice(&(header.len() as u32).to_le_bytes());
    bytes.extend_from_slice(&header);

    // EncryptionVerifier
    bytes.extend_from_slice(&16u32.to_le_bytes()); // saltSize
    bytes.extend_from_slice(&[0u8; 16]); // salt
    bytes.extend_from_slice(&[0u8; 16]); // encryptedVerifier
    bytes.extend_from_slice(&20u32.to_le_bytes()); // verifierHashSize (SHA1)
    bytes.extend_from_slice(&[0u8; 32]); // encryptedVerifierHash

    bytes
}

fn agile_encryption_info_bytes() -> Vec<u8> {
    // Copied from `formula-offcrypto` unit tests (`minimal_agile_xml`).
    //
    // This is a *valid* Agile EncryptionInfo XML payload where all encrypted blobs decode to
    // AES-block-aligned lengths (multiples of 16 bytes).
    //
    // It also intentionally includes:
    // - unpadded base64
    // - embedded whitespace
    //
    // to exercise tolerant decoding behavior.
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
    xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
  <keyData saltValue="AAECA wQFBg cICQo LDA0O Dw" hashAlgorithm="SHA256" blockSize="16"/>
  <dataIntegrity encryptedHmacKey="EBESE xQVFh cYGRo bHB0e HyAhI iMkJS YnKCk qKywt Li8" encryptedHmacValue="oKGio 6Slpq eoqaq rrK2u r7Cxs rO0tb a3uLm 6u7y9 vr8"/>
  <keyEncryptors>
    <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
      <p:encryptedKey spinCount="100000" saltValue="AQIDB AUGBw gJCgs MDQ4P EA" hashAlgorithm="SHA512" keyBits="256"
        encryptedKeyValue="ICEiI yQlJi coKSo rLC0u LzAxM jM0NT Y3ODk 6Ozw9 Pj8"
        encryptedVerifierHashInput="MDEyM zQ1Nj c4OTo 7PD0+ P0BBQ kNERU ZHSEl KS0xN Tk8"
        encryptedVerifierHashValue="QEFCQ 0RFRk dISUp LTE1O T1BRU lNUVV ZXWFl aW1xd Xl8"/>
    </keyEncryptor>
  </keyEncryptors>
</encryption>
"#;

    let mut bytes = Vec::new();
    bytes.extend_from_slice(&4u16.to_le_bytes());
    bytes.extend_from_slice(&4u16.to_le_bytes());
    bytes.extend_from_slice(&0u32.to_le_bytes());
    bytes.extend_from_slice(xml.as_bytes());
    bytes
}

#[test]
fn inspect_ooxml_encryption_returns_none_for_non_ole() {
    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("not-ole.bin");
    std::fs::write(&path, b"not an ole file").expect("write bytes");

    let res = inspect_ooxml_encryption(&path).expect("inspect");
    assert!(res.is_none());
}

#[test]
fn inspect_ooxml_encryption_parses_standard_encryption_info() {
    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("standard.xlsx");

    let bytes = ole_encrypted_ooxml_container(&standard_encryption_info_bytes());
    std::fs::write(&path, bytes).expect("write ole bytes");

    let summary = inspect_ooxml_encryption(&path)
        .expect("inspect")
        .expect("expected encrypted OOXML container");

    assert_eq!(summary.encryption_type, EncryptionType::Standard);
    assert!(summary.agile.is_none());
    assert_eq!(
        summary.standard,
        Some(StandardEncryptionInfoSummary {
            alg_id: StandardAlgId::Aes256,
            key_size: 256,
        })
    );
}

#[test]
fn inspect_ooxml_encryption_parses_agile_encryption_info() {
    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("agile.xlsx");

    let bytes = ole_encrypted_ooxml_container(&agile_encryption_info_bytes());
    std::fs::write(&path, bytes).expect("write ole bytes");

    let summary = inspect_ooxml_encryption(&path)
        .expect("inspect")
        .expect("expected encrypted OOXML container");

    assert_eq!(summary.encryption_type, EncryptionType::Agile);
    assert!(summary.standard.is_none());
    assert_eq!(
        summary.agile,
        Some(AgileEncryptionInfoSummary {
            hash_algorithm: HashAlgorithm::Sha512,
            spin_count: 100_000,
            key_bits: 256,
        })
    );
}
