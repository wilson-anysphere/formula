use formula_vba::extract_vba_signature_signed_digest;

fn wrap_in_digsig_info_serialized(pkcs7: &[u8]) -> Vec<u8> {
    // Synthetic DigSigInfoSerialized-like blob:
    // [cbSignature, cbSigningCertStore, cchProjectName] (LE u32)
    // [projectName UTF-16LE] [certStore bytes] [signature bytes]
    let project_name_utf16: Vec<u16> = "VBAProject\0".encode_utf16().collect();
    let mut project_name_bytes = Vec::new();
    for ch in &project_name_utf16 {
        project_name_bytes.extend_from_slice(&ch.to_le_bytes());
    }
    let cert_store = vec![0xAA, 0xBB, 0xCC, 0xDD];

    let cb_signature = pkcs7.len() as u32;
    let cb_cert_store = cert_store.len() as u32;
    let cch_project = project_name_utf16.len() as u32;

    let mut out = Vec::new();
    out.extend_from_slice(&cb_signature.to_le_bytes());
    out.extend_from_slice(&cb_cert_store.to_le_bytes());
    out.extend_from_slice(&cch_project.to_le_bytes());
    out.extend_from_slice(&project_name_bytes);
    out.extend_from_slice(&cert_store);
    out.extend_from_slice(pkcs7);
    out
}

#[test]
fn extracts_spc_indirect_data_digest_from_ber_indefinite_cms_at_offset_0() {
    let cms = include_bytes!("fixtures/cms_indefinite.der");

    let digest_info = extract_vba_signature_signed_digest(cms)
        .expect("extract should succeed")
        .expect("digest info should be present");

    assert_eq!(digest_info.digest_algorithm_oid, "2.16.840.1.101.3.4.2.1"); // SHA-256
    assert_eq!(digest_info.digest, (0u8..0x20).collect::<Vec<_>>());
}

#[test]
fn extracts_spc_indirect_data_digest_from_ber_indefinite_cms() {
    let cms = include_bytes!("fixtures/cms_indefinite.der");

    // Real-world VBA signature streams sometimes include a small prefix/header before the CMS blob.
    let mut stream = b"VBA\0SIG\0".to_vec();
    stream.extend_from_slice(cms);

    let digest_info = extract_vba_signature_signed_digest(&stream)
        .expect("extract should succeed")
        .expect("digest info should be present");

    assert_eq!(digest_info.digest_algorithm_oid, "2.16.840.1.101.3.4.2.1"); // SHA-256
    assert_eq!(digest_info.digest, (0u8..0x20).collect::<Vec<_>>());
}

#[test]
fn extracts_spc_indirect_data_digest_from_ber_indefinite_cms_wrapped_in_digsig_info_serialized() {
    let cms = include_bytes!("fixtures/cms_indefinite.der");
    let stream = wrap_in_digsig_info_serialized(cms);

    let digest_info = extract_vba_signature_signed_digest(&stream)
        .expect("extract should succeed")
        .expect("digest info should be present");

    assert_eq!(digest_info.digest_algorithm_oid, "2.16.840.1.101.3.4.2.1"); // SHA-256
    assert_eq!(digest_info.digest, (0u8..0x20).collect::<Vec<_>>());
}

#[test]
fn extracts_sigdata_v1_source_hash_from_ber_indefinite_cms_with_constructed_octet_string_econtent() {
    // Minimal SpcIndirectDataContentV2-like payload:
    // SEQUENCE { NULL, OCTET STRING(sigDataV1Serialized) }
    //
    // SigDataV1Serialized (binary-ish) is:
    // [version u32 LE = 1][cbSourceHash u32 LE = 16][sourceHash bytes]
    let source_hash = (0u8..16).collect::<Vec<_>>();
    let mut sigdata = Vec::new();
    sigdata.extend_from_slice(&1u32.to_le_bytes());
    sigdata.extend_from_slice(&(source_hash.len() as u32).to_le_bytes());
    sigdata.extend_from_slice(&source_hash);

    let mut spc_v2 = vec![
        0x30, 0x1C, // SEQUENCE (28 bytes)
        0x05, 0x00, // NULL
        0x04, 0x18, // OCTET STRING (24 bytes)
    ];
    spc_v2.extend_from_slice(&sigdata);
    assert_eq!(spc_v2.len(), 30);

    // Minimal BER-indefinite CMS ContentInfo/SignedData wrapper with:
    // - BER indefinite lengths at multiple levels
    // - eContent encoded as an *indefinite-length constructed OCTET STRING* (0x24 0x80)
    let split = 10;
    let (part1, part2) = spc_v2.split_at(split);

    let mut pkcs7 = vec![
        0x30, 0x80, // ContentInfo SEQUENCE (indefinite)
        0x06, 0x09, 0x2A, 0x86, 0x48, 0x86, 0xF7, 0x0D, 0x01, 0x07, 0x02, // OID 1.2.840.113549.1.7.2 (signedData)
        0xA0, 0x80, // [0] EXPLICIT (indefinite)
        0x30, 0x80, // SignedData SEQUENCE (indefinite)
        0x02, 0x01, 0x03, // version INTEGER 3
        0x31, 0x00, // digestAlgorithms SET (empty; we don't validate it here)
        0x30, 0x80, // encapContentInfo SEQUENCE (indefinite)
        0x06, 0x0A, 0x2B, 0x06, 0x01, 0x04, 0x01, 0x82, 0x37, 0x02, 0x01, 0x04, // OID 1.3.6.1.4.1.311.2.1.4 (SpcIndirectDataContent; not validated)
        0xA0, 0x80, // eContent [0] EXPLICIT (indefinite)
        0x24, 0x80, // OCTET STRING (constructed, indefinite)
        0x04,
        part1.len() as u8,
    ];
    pkcs7.extend_from_slice(part1);
    pkcs7.extend_from_slice(&[0x04, part2.len() as u8]);
    pkcs7.extend_from_slice(part2);
    pkcs7.extend_from_slice(&[
        0x00, 0x00, // EOC for constructed OCTET STRING
        0x00, 0x00, // EOC for eContent [0]
        0x00, 0x00, // EOC for encapContentInfo
        0x00, 0x00, // EOC for SignedData
        0x00, 0x00, // EOC for [0] EXPLICIT
        0x00, 0x00, // EOC for ContentInfo
    ]);

    let digest_info = extract_vba_signature_signed_digest(&pkcs7)
        .expect("extract should succeed")
        .expect("digest info should be present");

    assert_eq!(digest_info.digest_algorithm_oid, "1.2.840.113549.2.5"); // MD5
    assert_eq!(digest_info.digest, source_hash);
}

#[test]
fn extracts_spc_indirect_data_digest_from_ber_indefinite_cms_with_definite_length_constructed_octet_string_econtent(
) {
    // SpcIndirectDataContent (DER) whose digest is 0..31.
    let mut spc = vec![
        0x30, 0x41, // SEQUENCE
        0x30, 0x0c, // SEQUENCE
        0x06, 0x0a, 0x2b, 0x06, 0x01, 0x04, 0x01, 0x82, 0x37, 0x02, 0x01, 0x0f, // 1.3.6.1.4.1.311.2.1.15
        0x30, 0x31, // SEQUENCE
        0x30, 0x0d, // SEQUENCE
        0x06, 0x09, 0x60, 0x86, 0x48, 0x01, 0x65, 0x03, 0x04, 0x02, 0x01, // 2.16.840.1.101.3.4.2.1 (sha256)
        0x05, 0x00, // NULL
        0x04, 0x20, // OCTET STRING (32 bytes)
    ];
    spc.extend(0u8..0x20);

    // Constructed OCTET STRING with *definite length* (BER) containing two primitive segments.
    let split = 10;
    let (part1, part2) = spc.split_at(split);
    assert!(part1.len() < 128 && part2.len() < 128);

    // Content length is the sum of the child OCTET STRING TLVs.
    let constructed_content_len = (2 + part1.len()) + (2 + part2.len());
    assert!(constructed_content_len < 128);

    let mut pkcs7 = vec![
        0x30, 0x80, // ContentInfo SEQUENCE (indefinite)
        0x06, 0x09, 0x2A, 0x86, 0x48, 0x86, 0xF7, 0x0D, 0x01, 0x07, 0x02, // OID 1.2.840.113549.1.7.2 (signedData)
        0xA0, 0x80, // [0] EXPLICIT (indefinite)
        0x30, 0x80, // SignedData SEQUENCE (indefinite)
        0x02, 0x01, 0x03, // version INTEGER 3
        0x31, 0x00, // digestAlgorithms SET (empty; we don't validate it here)
        0x30, 0x80, // encapContentInfo SEQUENCE (indefinite)
        0x06, 0x0A, 0x2B, 0x06, 0x01, 0x04, 0x01, 0x82, 0x37, 0x02, 0x01, 0x04, // OID 1.3.6.1.4.1.311.2.1.4 (SpcIndirectDataContent)
        0xA0, 0x80, // eContent [0] EXPLICIT (indefinite)
        0x24,
        constructed_content_len as u8, // constructed OCTET STRING (definite length)
        0x04,
        part1.len() as u8,
    ];
    pkcs7.extend_from_slice(part1);
    pkcs7.extend_from_slice(&[0x04, part2.len() as u8]);
    pkcs7.extend_from_slice(part2);
    pkcs7.extend_from_slice(&[
        0x00, 0x00, // EOC for eContent [0]
        0x00, 0x00, // EOC for encapContentInfo
        0x00, 0x00, // EOC for SignedData
        0x00, 0x00, // EOC for [0] EXPLICIT
        0x00, 0x00, // EOC for ContentInfo
    ]);

    let digest_info = extract_vba_signature_signed_digest(&pkcs7)
        .expect("extract should succeed")
        .expect("digest info should be present");
    assert_eq!(digest_info.digest_algorithm_oid, "2.16.840.1.101.3.4.2.1"); // SHA-256
    assert_eq!(digest_info.digest, (0u8..0x20).collect::<Vec<_>>());
}

#[test]
fn extracts_spc_indirect_data_digest_from_ber_indefinite_cms_with_nested_constructed_octet_string_econtent(
) {
    // SpcIndirectDataContent (DER) whose digest is 0..31.
    let mut spc = vec![
        0x30, 0x41, // SEQUENCE
        0x30, 0x0c, // SEQUENCE
        0x06, 0x0a, 0x2b, 0x06, 0x01, 0x04, 0x01, 0x82, 0x37, 0x02, 0x01, 0x0f, // 1.3.6.1.4.1.311.2.1.15
        0x30, 0x31, // SEQUENCE
        0x30, 0x0d, // SEQUENCE
        0x06, 0x09, 0x60, 0x86, 0x48, 0x01, 0x65, 0x03, 0x04, 0x02, 0x01, // 2.16.840.1.101.3.4.2.1 (sha256)
        0x05, 0x00, // NULL
        0x04, 0x20, // OCTET STRING (32 bytes)
    ];
    spc.extend(0u8..0x20);

    // eContent is encoded as an *indefinite-length constructed OCTET STRING* (0x24 0x80) whose
    // children include another *definite-length constructed OCTET STRING*. This exercises nested
    // constructed OCTET STRING parsing + concatenation.
    let split1 = 10;
    let split2 = 30;
    let (part1, rest) = spc.split_at(split1);
    let (part2, part3) = rest.split_at(split2 - split1);
    assert!(part1.len() < 128 && part2.len() < 128 && part3.len() < 128);

    // Inner constructed OCTET STRING (definite length) contains part2 + part3.
    let inner_content_len = (2 + part2.len()) + (2 + part3.len());
    assert!(inner_content_len < 128);

    let mut pkcs7 = vec![
        0x30, 0x80, // ContentInfo SEQUENCE (indefinite)
        0x06, 0x09, 0x2A, 0x86, 0x48, 0x86, 0xF7, 0x0D, 0x01, 0x07, 0x02, // OID 1.2.840.113549.1.7.2 (signedData)
        0xA0, 0x80, // [0] EXPLICIT (indefinite)
        0x30, 0x80, // SignedData SEQUENCE (indefinite)
        0x02, 0x01, 0x03, // version INTEGER 3
        0x31, 0x00, // digestAlgorithms SET (empty; we don't validate it here)
        0x30, 0x80, // encapContentInfo SEQUENCE (indefinite)
        0x06, 0x0A, 0x2B, 0x06, 0x01, 0x04, 0x01, 0x82, 0x37, 0x02, 0x01, 0x04, // OID 1.3.6.1.4.1.311.2.1.4 (SpcIndirectDataContent)
        0xA0, 0x80, // eContent [0] EXPLICIT (indefinite)
        0x24, 0x80, // OCTET STRING (constructed, indefinite)
        // Child 1: primitive OCTET STRING(part1)
        0x04,
        part1.len() as u8,
    ];
    pkcs7.extend_from_slice(part1);

    // Child 2: constructed OCTET STRING(definite) containing two primitive segments.
    pkcs7.extend_from_slice(&[0x24, inner_content_len as u8, 0x04, part2.len() as u8]);
    pkcs7.extend_from_slice(part2);
    pkcs7.extend_from_slice(&[0x04, part3.len() as u8]);
    pkcs7.extend_from_slice(part3);

    pkcs7.extend_from_slice(&[
        0x00, 0x00, // EOC for outer constructed OCTET STRING
        0x00, 0x00, // EOC for eContent [0]
        0x00, 0x00, // EOC for encapContentInfo
        0x00, 0x00, // EOC for SignedData
        0x00, 0x00, // EOC for [0] EXPLICIT
        0x00, 0x00, // EOC for ContentInfo
    ]);

    let digest_info = extract_vba_signature_signed_digest(&pkcs7)
        .expect("extract should succeed")
        .expect("digest info should be present");
    assert_eq!(digest_info.digest_algorithm_oid, "2.16.840.1.101.3.4.2.1"); // SHA-256
    assert_eq!(digest_info.digest, (0u8..0x20).collect::<Vec<_>>());
}

#[test]
fn extracts_digest_from_indefinite_length_detached_signeddata_with_indefinite_encap_content_info() {
    // SpcIndirectDataContent (DER) whose digest is 0..31.
    let mut spc = vec![
        0x30, 0x41, // SEQUENCE
        0x30, 0x0c, // SEQUENCE
        0x06, 0x0a, 0x2b, 0x06, 0x01, 0x04, 0x01, 0x82, 0x37, 0x02, 0x01, 0x0f, // 1.3.6.1.4.1.311.2.1.15
        0x30, 0x31, // SEQUENCE
        0x30, 0x0d, // SEQUENCE
        0x06, 0x09, 0x60, 0x86, 0x48, 0x01, 0x65, 0x03, 0x04, 0x02, 0x01, // 2.16.840.1.101.3.4.2.1 (sha256)
        0x05, 0x00, // NULL
        0x04, 0x20, // OCTET STRING (32 bytes)
    ];
    spc.extend(0u8..0x20);

    // Minimal BER-indefinite CMS ContentInfo/SignedData wrapper with:
    // - signedData ContentInfo
    // - SignedData.encapContentInfo encoded as an *indefinite-length* SEQUENCE with no eContent
    //
    // This exercises the \"optional eContent\" handling where the next bytes in an indefinite-length
    // SEQUENCE are the EOC marker.
    let pkcs7 = vec![
        0x30, 0x80, // ContentInfo SEQUENCE (indefinite)
        0x06, 0x09, 0x2A, 0x86, 0x48, 0x86, 0xF7, 0x0D, 0x01, 0x07, 0x02, // OID 1.2.840.113549.1.7.2 (signedData)
        0xA0, 0x80, // [0] EXPLICIT (indefinite)
        0x30, 0x80, // SignedData SEQUENCE (indefinite)
        0x02, 0x01, 0x03, // version INTEGER 3
        0x31, 0x00, // digestAlgorithms SET (empty; we don't validate it here)
        0x30, 0x80, // encapContentInfo SEQUENCE (indefinite)
        0x06, 0x0A, 0x2B, 0x06, 0x01, 0x04, 0x01, 0x82, 0x37, 0x02, 0x01, 0x04, // OID 1.3.6.1.4.1.311.2.1.4 (SpcIndirectDataContent)
        0x00, 0x00, // EOC for encapContentInfo
        0x00, 0x00, // EOC for SignedData
        0x00, 0x00, // EOC for [0] EXPLICIT
        0x00, 0x00, // EOC for ContentInfo
    ];

    // Detached signature stream layout: signed content prefix + CMS signature.
    let mut stream = spc.clone();
    stream.extend_from_slice(&pkcs7);

    let digest_info = extract_vba_signature_signed_digest(&stream)
        .expect("extract should succeed")
        .expect("digest info should be present");
    assert_eq!(digest_info.digest_algorithm_oid, "2.16.840.1.101.3.4.2.1"); // SHA-256
    assert_eq!(digest_info.digest, (0u8..0x20).collect::<Vec<_>>());
}
